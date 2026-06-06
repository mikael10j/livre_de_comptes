"""Data models for the bookkeeping application."""
from dataclasses import dataclass
from typing import Optional


@dataclass
class Transaction:
    """Represents a financial transaction."""
    date_operation: str
    date_valeur: str
    libelle: str
    reference: str
    montant: float
    categorie: str
    source: str
    is_income: bool = False


@dataclass
class ParsedPDFResult:
    """Result of PDF parsing operation."""
    income_transactions: list[Transaction]
    expense_transactions: list[Transaction]
    errors: list[str]