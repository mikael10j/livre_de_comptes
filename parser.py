"""PDF parsing module for extracting financial transactions."""
import re
import logging
from pathlib import Path
from typing import Tuple, Optional
import pandas as pd
import camelot

from models import Transaction, ParsedPDFResult
from config import INCOME_RULES, EXPENSE_RULES


logger = logging.getLogger(__name__)


def parse_amount(amount_str: str) -> float:
    """Parse an amount from a string, handles different formats."""
    if not amount_str or amount_str.strip() == "":
        return 0.0
    
    amount_clean = amount_str.strip().replace(" ", "").replace(",", ".")
    
    try:
        return float(amount_clean)
    except ValueError:
        import string
        allowed = string.digits + ".,-+"
        amount_clean = "".join(c for c in amount_clean if c in allowed)
        amount_clean = amount_clean.replace(",", ".")
        try:
            return float(amount_clean)
        except ValueError:
            logger.warning(f"Could not parse amount: {amount_str}")
            return 0.0


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    """Find the header row containing 'Libellé', 'Débit', 'Crédit'."""
    for idx, row in df.iterrows():
        row_text = " ".join(str(cell) for cell in row.values).upper()
        if "LIBELLÉ" in row_text and ("DÉBIT" in row_text or "CRÉDIT" in row_text):
            return idx
    return None


def create_named_dataframe(df: pd.DataFrame, header_row_idx: int) -> Optional[pd.DataFrame]:
    """Create a DataFrame with column names based on the header row."""
    if header_row_idx is None:
        return None
        
    headers = df.iloc[header_row_idx].values
    
    column_names = []
    for i, header in enumerate(headers):
        header_clean = str(header).strip().upper()
        if "LIBELLÉ" in header_clean or "LIBELLE" in header_clean:
            column_names.append("LIBELLE")
        elif "DÉBIT" in header_clean or "DEBIT" in header_clean:
            column_names.append("DEBIT")
        elif "CRÉDIT" in header_clean or "CREDIT" in header_clean:
            column_names.append("CREDIT")
        elif "DATE" in header_clean and "OPÉ" in header_clean:
            column_names.append("DATE_OPE")
        elif "DATE" in header_clean and ("VAL" in header_clean or "VALEUR" in header_clean):
            column_names.append("DATE_VAL")
        else:
            column_names.append(f"COL_{i}")
    
    data_rows = df.iloc[header_row_idx + 1:].copy()
    data_rows.columns = column_names
    
    return data_rows


def categorize_transaction(label: str, rules: list, default: str = "À préciser") -> str:
    """Categorize a transaction based on predefined rules."""
    for pattern, category in rules:
        if re.search(pattern, label.lower()):
            return category
    return default


def extract_operations_pdf_camelot(pdf_path: Path, year: int) -> ParsedPDFResult:
    """Extract operations from a PDF using camelot and pandas."""
    income_transactions = []
    expense_transactions = []
    errors = []

    parameters_camelot = [
        {"flavor": "stream", "pages": "all", "edge_tol": 500},
        {"flavor": "stream", "pages": "all"},
        {"flavor": "stream", "pages": "all", "row_tol": 10},
    ]
    
    table_operations = None
    
    for i, params in enumerate(parameters_camelot):
        try:
            logger.info(f"Testing parameter set {i+1}: {params}")
            tables_test = camelot.read_pdf(str(pdf_path), strip_text="\n", **params)
            
            for table in tables_test:
                df = table.df
                header_row_idx = find_header_row(df)
                
                if header_row_idx is not None:
                    table_operations = create_named_dataframe(df, header_row_idx)
                    logger.info(f"Success with parameter set {i+1}: Headers found at row {header_row_idx}")
                    break
            
            if table_operations is not None:
                break
                
            logger.info(f"Parameter set {i+1}: no table with headers found")
                
        except Exception as e:
            logger.error(f"Error with parameter set {i+1}: {e}")
            continue
    
    if table_operations is None:
        error_msg = f"No operations table with headers found in {pdf_path.name}"
        logger.error(error_msg)
        errors.append(error_msg)
        return ParsedPDFResult([], [], errors)
        
    logger.info(f"Operations table found with columns: {list(table_operations.columns)}")

    operations_found = 0
    
    for idx, row in table_operations.iterrows():
        date_ope = str(row.iloc[0]).strip() if len(row) > 0 else ""
        date_val = str(row.iloc[1]).strip() if len(row) > 1 else ""
        
        if not re.match(r'\d{2}\.\d{2}', date_ope):
            continue
            
        libelle = ""
        if "LIBELLE" in table_operations.columns:
            libelle = str(row["LIBELLE"]).strip()
        else:
            libelle = str(row.iloc[2]).strip() if len(row) > 2 else ""
        
        montant = 0.0
        est_recette = False
        
        debit_value = ""
        credit_value = ""
        
        if "DEBIT" in table_operations.columns:
            debit_value = str(row["DEBIT"]).strip()
        if "CREDIT" in table_operations.columns:
            credit_value = str(row["CREDIT"]).strip()
        
        debit_clean = re.sub(r'[¨\s]', '', debit_value) if debit_value else ""
        credit_clean = re.sub(r'[¨\s]', '', credit_value) if credit_value else ""
        
        if debit_clean and debit_clean not in ["", "nan", "0"]:
            montant = parse_amount(debit_clean)
            est_recette = False
        elif credit_clean and credit_clean not in ["", "nan", "0"]:
            montant = parse_amount(credit_clean)
            est_recette = True
        else:
            for col_name in table_operations.columns:
                if col_name not in ["LIBELLE", "DATE_OPE", "DATE_VAL"]:
                    val_str = str(row[col_name]).strip()
                    val_clean = re.sub(r'[¨\s]', '', val_str)
                    
                    if val_clean and val_clean not in ["", "nan"]:
                        if re.match(r'^\d{1,4}(?:\s?\d{3})*,\d{2}$', val_clean):
                            montant = parse_amount(val_clean)
                            if montant > 0:
                                libelle_upper = libelle.upper()
                                mots_recette = ["REM CHQ", "REMISE", "VERSEMENT", "DEPOT", "VIREMENT", "REGUL", "VIR RECU"]
                                est_recette = any(mot in libelle_upper for mot in mots_recette)
                                break
        
        if montant <= 0:
            continue
            
        operations_found += 1
        
        date_ope_formatted = date_ope.replace('.', '/')
        date_val_formatted = date_val.replace('.', '/')
        
        transaction = Transaction(
            date_operation=f"{date_ope_formatted}/{year}",
            date_valeur=f"{date_val_formatted}/{year}",
            libelle=libelle.strip(),
            reference="",
            montant=montant,
            categorie="",  # Will be filled after categorization
            source=pdf_path.stem,
            is_income=est_recette
        )

        if est_recette:
            transaction.categorie = categorize_transaction(libelle, INCOME_RULES)
            income_transactions.append(transaction)
        else:
            transaction.categorie = categorize_transaction(libelle, EXPENSE_RULES)
            expense_transactions.append(transaction)

    logger.info(f"Extracted {operations_found} operations from {pdf_path.name}")
    return ParsedPDFResult(income_transactions, expense_transactions, errors)