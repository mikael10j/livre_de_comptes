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