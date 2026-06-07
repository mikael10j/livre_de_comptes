"""
Copyright 2026 mikael10j

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

"""Main application for generating bookkeeping reports from bank statements."""
import argparse
import logging
from pathlib import Path
from datetime import datetime
import sys

from parser import extract_operations_pdf_camelot
from excel_writer import generate_or_update_excel
from models import Transaction


# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def sort_transactions(transactions: list[Transaction]) -> list[Transaction]:
    """Sort transactions by date."""
    def sort_key(op):
        try:
            return datetime.strptime(op.date_operation, "%d/%m/%Y")
        except ValueError:
            return datetime.min

    return sorted(transactions, key=sort_key)


def main():
    parser = argparse.ArgumentParser(description="Génère un livre de comptes Excel depuis des relevés PDF.")
    parser.add_argument("pdfs", nargs="+", help="Fichiers PDF à traiter")
    parser.add_argument("-o", "--output", default="livre_comptes.xlsx",
                        help="Fichier Excel de sortie (défaut: livre_comptes.xlsx)")
    parser.add_argument("-a", "--annee", type=int, default=datetime.now().year,
                        help="Année des relevés (défaut: année en cours)")
    parser.add_argument("-f", "--force", action="store_true",
                        help="Remplacer l'onglet existant sans confirmation")
    args = parser.parse_args()
    args.output = Path(args.output)

    logger.info(f"Bookkeeping - Year {args.annee}")
    print("=" * 50)

    valid_pdfs = []
    for p in args.pdfs:
        path = Path(p)
        if not path.exists():
            logger.warning(f"File not found: {p}")
        elif path.suffix.lower() != ".pdf":
            logger.warning(f"Not a PDF: {p}")
        else:
            valid_pdfs.append(path)

    if not valid_pdfs:
        logger.error("No valid PDF files found.")
        sys.exit(1)

    logger.info(f"Extracting from {len(valid_pdfs)} file(s)...")
    all_income_transactions = []
    all_expense_transactions = []
    
    for pdf in sorted(valid_pdfs):
        logger.info(f"Processing {pdf.name}...")
        try:
            result = extract_operations_pdf_camelot(pdf, args.annee)
            logger.info(f"  -> {len(result.income_transactions)} income(s), {len(result.expense_transactions)} expense(s)")
            
            # Add any errors to the log
            for error in result.errors:
                logger.error(f"Error processing {pdf.name}: {error}")
            
            all_income_transactions.extend(result.income_transactions)
            all_expense_transactions.extend(result.expense_transactions)
        except Exception as e:
            logger.error(f"Error extracting from PDF {pdf.name}: {e}")
            continue

    # Sort transactions by date
    all_income_transactions = sort_transactions(all_income_transactions)
    all_expense_transactions = sort_transactions(all_expense_transactions)

    logger.info(f"Total: {len(all_income_transactions)} income transactions, {len(all_expense_transactions)} expense transactions")

    generate_or_update_excel(all_income_transactions, all_expense_transactions, args.output, args.annee, args.force)


if __name__ == "__main__":
    main()