# 📒 Livre de Comptes

Script automatisé pour générer un livre de comptes Excel à partir de relevés bancaires PDF.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Copyright 2026 mikael10j

## ✨ Fonctionnalités

- **Extraction robuste** : Utilise camelot et pandas avec détection automatique des colonnes
- **Multi-format** : Compatible avec différents formats de relevés (2019, 2023+)
- **Classification intelligente** : Catégorisation automatique des opérations (recettes/dépenses)
- **Synthèse pluriannuelle** : Génère automatiquement des graphiques et statistiques
- **Interface Excel complète** : Mise en forme professionnelle avec validations

## 🔧 Installation

```bash
# Créer un environnement virtuel
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# ou .venv\Scripts\activate  # Windows

# Installer les dépendances
pip install -r requirements.txt
```

## 🚀 Utilisation

### Première année
```bash
python livre_comptes.py *.pdf -a 2019 -o comptes_asso.xlsx
```
→ Crée `comptes_asso.xlsx` avec l'onglet "Comptes 2019" et la synthèse

### Année suivante  
```bash
python livre_comptes.py "Relevé 2023.pdf" -a 2023 -o comptes_asso.xlsx
```
→ Ajoute l'onglet "Comptes 2023" et met à jour la synthèse

### Retraitement (avec confirmation)
```bash
python livre_comptes.py *.pdf -a 2019 -f
```
→ Remplace l'onglet existant sans confirmation (`-f` = force)

## 📊 Formats supportés

- ✅ Relevés Crédit Agricole 2019
- ✅ Relevés Crédit Agricole 2023+  
- ✅ Auto-détection des colonnes Débit/Crédit
- ✅ Gestion des caractères spéciaux (¨, espaces)

## 🎯 Robustesse

Le script utilise une **approche intelligente** :

1. **Détection automatique** des en-têtes ("Libellé", "Débit", "Crédit")
2. **Mapping par nom de colonne** (plus robuste que par position)
3. **Fallback adaptatif** pour les formats atypiques
4. **Validation des montants** et classification selon le libellé

## 🏗️ Architecture du projet

Le code a été réorganisé en modules pour améliorer la maintenabilité :

- `main.py` - Point d'entrée principal de l'application
- `config.py` - Configuration des couleurs, catégories et règles
- `models.py` - Modèles de données (Transaction, etc.)
- `parser.py` - Extraction et traitement des données PDF
- `excel_writer.py` - Génération des fichiers Excel
- `livre_comptes.py` - Ancien fichier principal, maintenant point d'entrée pour compatibilité
- `requirements.txt` - Dépendances du projet

## 🧪 Tests

Le projet inclut une suite de tests complète organisée dans le répertoire `tests/` :

```bash
# Exécuter tous les tests
python run_tests.py

# Tests spécifiques
python tests/test_unitaires.py      # Tests unitaires
python tests/test_performance.py    # Tests de performance  
python tests/test_integration.py    # Tests avec vrais PDFs

# Avec pytest (si installé)
pytest -v
```

### Types de tests inclus

- **Tests unitaires** (18 tests) : Fonctions individuelles, parsing, catégorisation
- **Tests de performance** (6 tests) : Vitesse d'exécution, robustesse, limites mémoire  
- **Tests d'intégration** (6 tests) : Extraction avec vrais PDFs, cohérence des données

## 📁 Structure du projet

```
livre de comptes/
├── livre_comptes.py          # Script principal (point d'entrée)
├── main.py                  # Programme principal
├── config.py                # Configuration
├── models.py                # Modèles de données
├── parser.py                # Extraction des données PDF
├── excel_writer.py          # Génération Excel
├── requirements.txt         # Dépendances
├── livre_comptes.xlsx        # Fichier de sortie
├── run_tests.py             # Lanceur de tous les tests
├── pytest.ini              # Configuration pytest
├── tests/                   # Répertoire des tests
│   ├── __init__.py         # Package tests
│   ├── test_unitaires.py   # Tests unitaires (18 tests)
│   ├── test_performance.py # Tests de performance (6 tests)
│   └── test_integration.py # Tests d'intégration (6 tests)
├── *.pdf                    # Relevés bancaires
└── README.md               # Cette documentation
```

*Développé à l'aide de Claude Sonnet 4.6 et Qwen3 Coder Plus*