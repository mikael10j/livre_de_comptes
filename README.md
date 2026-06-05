# 📒 Livre de Comptes - Les Vétérans du Basket

Script automatisé pour générer un livre de comptes Excel à partir de relevés bancaires PDF.

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
pip install camelot-py[cv] pandas openpyxl
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

## 📁 Structure du projet

```
livre de comptes/
├── livre_comptes.py          # Script principal
├── livre_comptes.xlsx        # Fichier de sortie
├── *.pdf                     # Relevés bancaires
└── README.md                # Cette documentation
```