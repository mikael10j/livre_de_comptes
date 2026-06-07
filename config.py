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

"""Configuration file for the bookkeeping application."""

# Colors for Excel formatting
COLORS = {
    "titre": "1F3864", 
    "recette": "1E8449", 
    "recette_clair": "D5F5E3",
    "depense": "922B21", 
    "depense_clair": "FADBD8", 
    "recap": "1A5276", 
    "recap_clair": "D6EAF8", 
    "synthese": "4A235A", 
    "synthese_clair": "E8DAEF",
    "alternance": "F2F3F4", 
    "blanc": "FFFFFF", 
    "or": "D4AC0D",
}

# Categories
INCOME_CATEGORIES = [
    "Cotisations", 
    "Buvette/Tournoi", 
    "Sponsoring/Dons",
    "Subventions", 
    "Autres", 
    "À préciser"
]

EXPENSE_CATEGORIES = [
    "Matériel sportif", 
    "Assurance",
    "Frais bancaires", 
    "Buvette/Tournoi", 
    "Location salle",
    "Déplacements", 
    "Autres", 
    "À préciser", 
    "Don", 
    "WE Vétérans"
]

# Categorization rules
INCOME_RULES = [
    (r"cotisation|adhesion|licence", "Cotisations"),
    (r"rem chq|remise.*chq|remise de cheque", "Buvette/Tournoi"),
    (r"regul.*verst|regul.*gab", "Buvette/Tournoi"),
    (r"loomis|gab|versement esp", "Buvette/Tournoi"),
    (r"tournoi|buvette|snack", "Buvette/Tournoi"),
    (r"sponsor|parrain|don", "Sponsoring/Dons"),
    (r"subvention|aide|grant", "Subventions"),
    (r"virement.*recu|vir.*recu", "Sponsoring/Dons"),
]

EXPENSE_RULES = [
    (r"retrait|gab|distributeur", "Buvette/Tournoi"),
    (r"cheque|chq(?!.*rem)", "Autres"),
    (r"smacl|maif|assurance|insurance", "Assurance"),
    (r"intérêts|interets|commission|frais bancaires|cotisation bancaire", "Frais bancaires"),
    (r"décathlon|decathlon|intersport|sport 2000|go sport|matériel", "Matériel sportif"),
    (r"location|salle|gymnase|terrain", "Location salle"),
    (r"transport|essence|carburant|péage|deplacement", "Déplacements"),
    (r"tournoi|buvette|restauration", "Buvette/Tournoi"),
]