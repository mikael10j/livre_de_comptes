#!/usr/bin/env python3
"""
Génère un livre de comptes Excel à partir de relevés bancaires PDF.

Usage:
    python livre_comptes.py releve1.pdf releve2.pdf [...] [-o sortie.xlsx] [-a ANNEE] [-f]
"""

import argparse
import re
import sys
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import camelot
    import pandas as pd
except ImportError:
    print("❌ Installez les dépendances : pip install camelot-py[cv] pandas openpyxl")
    sys.exit(1)


# Couleurs pour la mise en forme Excel
COULEURS = {
    "titre": "1F3864", "recette": "1E8449", "recette_clair": "D5F5E3",
    "depense": "922B21", "depense_clair": "FADBD8", "recap": "1A5276", 
    "recap_clair": "D6EAF8", "synthese": "4A235A", "synthese_clair": "E8DAEF",
    "alternance": "F2F3F4", "blanc": "FFFFFF", "or": "D4AC0D",
}

# Catégories de recettes et dépenses
CATEGORIES_RECETTES = [
    "Cotisations", "Buvette/Tournoi", "Sponsoring/Dons",
    "Subventions", "Autres", "À préciser"
]
CATEGORIES_DEPENSES = [
    "Matériel sportif", "Assurance",
    "Frais bancaires", "Buvette/Tournoi", "Location salle",
    "Déplacements", "Autres", "À préciser", "Don", "WE Vétérans"
]

# Règles de catégorisation automatique
REGLES_RECETTES = [
    (r"cotisation|adhesion|licence", "Cotisations"),
    (r"rem chq|remise.*chq|remise de cheque", "Buvette/Tournoi"),
    (r"regul.*verst|regul.*gab", "Buvette/Tournoi"),
    (r"loomis|gab|versement esp", "Buvette/Tournoi"),
    (r"tournoi|buvette|snack", "Buvette/Tournoi"),
    (r"sponsor|parrain|don", "Sponsoring/Dons"),
    (r"subvention|aide|grant", "Subventions"),
    (r"virement.*recu|vir.*recu", "Sponsoring/Dons"),
]
REGLES_DEPENSES = [
    (r"retrait|gab|distributeur", "Buvette/Tournoi"),
    (r"cheque|chq(?!.*rem)", "Autres"),
    (r"smacl|maif|assurance|insurance", "Assurance"),
    (r"intérêts|interets|commission|frais bancaires|cotisation bancaire", "Frais bancaires"),
    (r"décathlon|decathlon|intersport|sport 2000|go sport|matériel", "Matériel sportif"),
    (r"location|salle|gymnase|terrain", "Location salle"),
    (r"transport|essence|carburant|péage|deplacement", "Déplacements"),
    (r"tournoi|buvette|restauration", "Buvette/Tournoi"),
]


def style_bordure(fin=False):
    s = "thin" if not fin else "hair"
    return Border(left=Side(style=s), right=Side(style=s), top=Side(style=s), bottom=Side(style=s))

def fill(couleur):
    return PatternFill("solid", fgColor=couleur)

def police(gras=False, taille=11, couleur="000000", italique=False):
    return Font(bold=gras, size=taille, color=couleur, italic=italique)

def centrer(horizontal="center", vertical="center", wrap=False):
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)

def categoriser(libelle: str, regles: list, defaut="À préciser") -> str:
    for pattern, cat in regles:
        if re.search(pattern, libelle.lower()):
            return cat
    return defaut

def parse_montant(montant_str: str) -> float:
    """Parse un montant depuis une chaîne, gère différents formats."""
    if not montant_str or montant_str.strip() == "":
        return 0.0
    
    montant_clean = montant_str.strip().replace(" ", "").replace(",", ".")
    
    try:
        return float(montant_clean)
    except ValueError:
        import string
        allowed = string.digits + ".,-+"
        montant_clean = "".join(c for c in montant_clean if c in allowed)
        montant_clean = montant_clean.replace(",", ".")
        try:
            return float(montant_clean)
        except ValueError:
            return 0.0

def find_header_row(df):
    """Trouve la ligne d'en-têtes contenant 'Libellé', 'Débit', 'Crédit'."""
    for idx, row in df.iterrows():
        row_text = " ".join(str(cell) for cell in row.values).upper()
        if "LIBELLÉ" in row_text and ("DÉBIT" in row_text or "CRÉDIT" in row_text):
            return idx
    return None

def create_named_dataframe(df, header_row_idx):
    """Crée un DataFrame avec des noms de colonnes basés sur la ligne d'en-têtes."""
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

def extraire_operations_pdf_camelot(pdf_path: Path, annee: int) -> tuple[list, list]:
    """Extrait les opérations depuis un PDF en utilisant camelot et pandas avec détection automatique des colonnes."""
    recettes, depenses = [], []

    paramètres_camelot = [
        {"flavor": "stream", "pages": "all", "edge_tol": 500},
        {"flavor": "stream", "pages": "all"},
        {"flavor": "stream", "pages": "all", "row_tol": 10},
    ]
    
    table_operations = None
    
    for i, params in enumerate(paramètres_camelot):
        try:
            print(f"    🔧 Test paramètre {i+1}: {params}")
            tables_test = camelot.read_pdf(str(pdf_path), strip_text="\n", **params)
            
            for table in tables_test:
                df = table.df
                header_row_idx = find_header_row(df)
                
                if header_row_idx is not None:
                    table_operations = create_named_dataframe(df, header_row_idx)
                    print(f"    ✅ Succès avec paramètre {i+1}: En-têtes trouvés à la ligne {header_row_idx}")
                    break
            
            if table_operations is not None:
                break
                
            print(f"    ❌ Paramètre {i+1}: aucune table avec en-têtes trouvée")
                
        except Exception as e:
            print(f"    ❌ Erreur avec paramètre {i+1}: {e}")
            continue
    
    if table_operations is None:
        print(f"❌ Aucune table d'opérations avec en-têtes trouvée dans {pdf_path.name}")
        return [], []
        
    print(f"    📊 Table d'opérations trouvée avec colonnes: {list(table_operations.columns)}")

    operations_trouvees = 0
    
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
            montant = parse_montant(debit_clean)
            est_recette = False
        elif credit_clean and credit_clean not in ["", "nan", "0"]:
            montant = parse_montant(credit_clean)
            est_recette = True
        else:
            for col_name in table_operations.columns:
                if col_name not in ["LIBELLE", "DATE_OPE", "DATE_VAL"]:
                    val_str = str(row[col_name]).strip()
                    val_clean = re.sub(r'[¨\s]', '', val_str)
                    
                    if val_clean and val_clean not in ["", "nan"]:
                        if re.match(r'^\d{1,4}(?:\s?\d{3})*,\d{2}$', val_clean):
                            montant = parse_montant(val_clean)
                            if montant > 0:
                                libelle_upper = libelle.upper()
                                mots_recette = ["REM CHQ", "REMISE", "VERSEMENT", "DEPOT", "VIREMENT", "REGUL", "VIR RECU"]
                                est_recette = any(mot in libelle_upper for mot in mots_recette)
                                break
        
        if montant <= 0:
            continue
            
        operations_trouvees += 1
        
        date_ope_formatted = date_ope.replace('.', '/')
        date_val_formatted = date_val.replace('.', '/')
        
        op = [
            f"{date_ope_formatted}/{annee}",
            f"{date_val_formatted}/{annee}",
            libelle.strip(),
            "",
            montant,
            "",
            pdf_path.stem,
        ]

        if est_recette:
            op[5] = categoriser(libelle, REGLES_RECETTES)
            recettes.append(op)
        else:
            op[5] = categoriser(libelle, REGLES_DEPENSES)
            depenses.append(op)

    print(f"    📋 {operations_trouvees} opérations extraites")
    return recettes, depenses

def ecrire_onglet_annuel(ws, recettes: list, depenses: list, annee: int):
    """Remplit un onglet avec recettes, dépenses, récap et ventilation."""

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    largeurs = {"A": 14, "B": 14, "C": 45, "D": 18, "E": 14, "F": 22, "G": 20}
    for col, larg in largeurs.items():
        ws.column_dimensions[col].width = larg

    entetes = ["Date opération", "Date valeur", "Libellé", "Référence",
               "Montant (€)", "Catégorie", "Source"]

    ligne = 1

    # Titre général
    ws.merge_cells(f"A{ligne}:G{ligne}")
    c = ws.cell(ligne, 1, f"📒  LIVRE DE COMPTES {annee}")
    c.font = police(gras=True, taille=16, couleur="FFFFFF")
    c.fill = fill(COULEURS["titre"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 32
    ligne += 2

    # Section RECETTES
    ws.merge_cells(f"A{ligne}:G{ligne}")
    c = ws.cell(ligne, 1, "💰  RECETTES")
    c.font = police(gras=True, taille=13, couleur="FFFFFF")
    c.fill = fill(COULEURS["recette"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 26
    ligne += 1

    # En-têtes recettes
    ligne_entete_rec = ligne
    for col, titre in enumerate(entetes, 1):
        c = ws.cell(ligne, col, titre)
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(COULEURS["recette"])
        c.alignment = centrer()
        c.border = style_bordure()
    ws.row_dimensions[ligne].height = 20
    ligne += 1

    # Données recettes
    ligne_debut_rec = ligne
    cats_rec_dv = ",".join(CATEGORIES_RECETTES)
    dv_rec = DataValidation(type="list", formula1=f'"{cats_rec_dv}"', showDropDown=False)
    ws.add_data_validation(dv_rec)

    for i, op in enumerate(recettes):
        bg = COULEURS["blanc"] if i % 2 == 0 else COULEURS["alternance"]
        for col, val in enumerate(op, 1):
            c = ws.cell(ligne, col, val)
            c.fill = fill(bg)
            c.border = style_bordure(fin=True)
            c.alignment = centrer("left") if col == 3 else centrer()
            if col == 5:
                c.number_format = '#,##0.00 €'
                c.font = police(couleur=COULEURS["recette"])
        dv_rec.add(ws.cell(ligne, 6))
        ligne += 1

    ligne_fin_rec = ligne - 1

    # Total recettes
    ws.merge_cells(f"A{ligne}:D{ligne}")
    c = ws.cell(ligne, 1, "TOTAL RECETTES")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["recette"])
    c.alignment = centrer("right")
    c.border = style_bordure()

    col_total_rec = 5
    cell_total_rec = ws.cell(ligne, col_total_rec)
    if recettes:
        cell_total_rec.value = f"=SUM(E{ligne_debut_rec}:E{ligne_fin_rec})"
    else:
        cell_total_rec.value = 0
    cell_total_rec.font = police(gras=True, couleur="FFFFFF")
    cell_total_rec.fill = fill(COULEURS["recette"])
    cell_total_rec.number_format = '#,##0.00 €'
    cell_total_rec.alignment = centrer()
    cell_total_rec.border = style_bordure()
    ref_total_rec = f"E{ligne}"
    ws.row_dimensions[ligne].height = 22
    ligne += 2

    # Section DÉPENSES
    ws.merge_cells(f"A{ligne}:G{ligne}")
    c = ws.cell(ligne, 1, "💸  DÉPENSES")
    c.font = police(gras=True, taille=13, couleur="FFFFFF")
    c.fill = fill(COULEURS["depense"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 26
    ligne += 1

    # En-têtes dépenses
    for col, titre in enumerate(entetes, 1):
        c = ws.cell(ligne, col, titre)
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(COULEURS["depense"])
        c.alignment = centrer()
        c.border = style_bordure()
    ws.row_dimensions[ligne].height = 20
    ligne += 1

    # Données dépenses
    ligne_debut_dep = ligne
    cats_dep_dv = ",".join(CATEGORIES_DEPENSES)
    dv_dep = DataValidation(type="list", formula1=f'"{cats_dep_dv}"', showDropDown=False)
    ws.add_data_validation(dv_dep)

    for i, op in enumerate(depenses):
        bg = COULEURS["blanc"] if i % 2 == 0 else COULEURS["alternance"]
        for col, val in enumerate(op, 1):
            c = ws.cell(ligne, col, val)
            c.fill = fill(bg)
            c.border = style_bordure(fin=True)
            c.alignment = centrer("left") if col == 3 else centrer()
            if col == 5:
                c.number_format = '#,##0.00 €'
                c.font = police(couleur=COULEURS["depense"])
        dv_dep.add(ws.cell(ligne, 6))
        ligne += 1

    ligne_fin_dep = ligne - 1

    # Total dépenses
    ws.merge_cells(f"A{ligne}:D{ligne}")
    c = ws.cell(ligne, 1, "TOTAL DÉPENSES")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["depense"])
    c.alignment = centrer("right")
    c.border = style_bordure()

    cell_total_dep = ws.cell(ligne, 5)
    if depenses:
        cell_total_dep.value = f"=SUM(E{ligne_debut_dep}:E{ligne_fin_dep})"
    else:
        cell_total_dep.value = 0
    cell_total_dep.font = police(gras=True, couleur="FFFFFF")
    cell_total_dep.fill = fill(COULEURS["depense"])
    cell_total_dep.number_format = '#,##0.00 €'
    cell_total_dep.alignment = centrer()
    cell_total_dep.border = style_bordure()
    ref_total_dep = f"E{ligne}"
    ws.row_dimensions[ligne].height = 22
    ligne += 2

    # Récapitulatif annuel
    ws.merge_cells(f"A{ligne}:G{ligne}")
    c = ws.cell(ligne, 1, "📊  RÉCAPITULATIF ANNUEL")
    c.font = police(gras=True, taille=13, couleur="FFFFFF")
    c.fill = fill(COULEURS["recap"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 26
    ligne += 1

    recap_items = [
        ("Total recettes", f"={ref_total_rec}", COULEURS["recette_clair"], COULEURS["recette"]),
        ("Total dépenses", f"={ref_total_dep}", COULEURS["depense_clair"], COULEURS["depense"]),
        ("Solde de l'année", f"={ref_total_rec}-{ref_total_dep}", COULEURS["recap_clair"], COULEURS["recap"]),
    ]

    for label, formule, bg_clair, bg_fort in recap_items:
        ws.merge_cells(f"A{ligne}:D{ligne}")
        c = ws.cell(ligne, 1, label)
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(bg_fort)
        c.alignment = centrer("right")
        c.border = style_bordure()

        c2 = ws.cell(ligne, 5, formule)
        c2.font = police(gras=True)
        c2.fill = fill(bg_clair)
        c2.number_format = '#,##0.00 €'
        c2.alignment = centrer()
        c2.border = style_bordure()
        ws.row_dimensions[ligne].height = 22
        ligne += 1

    ligne += 1

    # Ventilation par catégorie
    ws.merge_cells(f"A{ligne}:G{ligne}")
    c = ws.cell(ligne, 1, "🎯  VENTILATION PAR CATÉGORIE")
    c.font = police(gras=True, taille=13, couleur="FFFFFF")
    c.fill = fill(COULEURS["recap"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 26
    ligne += 1

    # Recettes par catégorie
    ws.merge_cells(f"A{ligne}:C{ligne}")
    c = ws.cell(ligne, 1, "Catégorie recette")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["recette"])
    c.alignment = centrer()
    c.border = style_bordure()

    ws.merge_cells(f"D{ligne}:E{ligne}")
    c = ws.cell(ligne, 4, "Montant (€)")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["recette"])
    c.alignment = centrer()
    c.border = style_bordure()
    ligne += 1

    for cat in CATEGORIES_RECETTES:
        if recettes:
            formule = (f'=SUMIF(F{ligne_debut_rec}:F{ligne_fin_rec},'
                       f'"{cat}",E{ligne_debut_rec}:E{ligne_fin_rec})')
        else:
            formule = 0
        ws.merge_cells(f"A{ligne}:C{ligne}")
        c = ws.cell(ligne, 1, cat)
        c.fill = fill(COULEURS["recette_clair"])
        c.border = style_bordure(fin=True)
        c.alignment = centrer("left")

        ws.merge_cells(f"D{ligne}:E{ligne}")
        c2 = ws.cell(ligne, 4, formule)
        c2.number_format = '#,##0.00 €'
        c2.fill = fill(COULEURS["recette_clair"])
        c2.border = style_bordure(fin=True)
        c2.alignment = centrer()
        ligne += 1

    ligne += 1

    # Dépenses par catégorie
    ws.merge_cells(f"A{ligne}:C{ligne}")
    c = ws.cell(ligne, 1, "Catégorie dépense")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["depense"])
    c.alignment = centrer()
    c.border = style_bordure()

    ws.merge_cells(f"D{ligne}:E{ligne}")
    c = ws.cell(ligne, 4, "Montant (€)")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["depense"])
    c.alignment = centrer()
    c.border = style_bordure()
    ligne += 1

    for cat in CATEGORIES_DEPENSES:
        if depenses:
            formule = (f'=SUMIF(F{ligne_debut_dep}:F{ligne_fin_dep},'
                       f'"{cat}",E{ligne_debut_dep}:E{ligne_fin_dep})')
        else:
            formule = 0
        ws.merge_cells(f"A{ligne}:C{ligne}")
        c = ws.cell(ligne, 1, cat)
        c.fill = fill(COULEURS["depense_clair"])
        c.border = style_bordure(fin=True)
        c.alignment = centrer("left")

        ws.merge_cells(f"D{ligne}:E{ligne}")
        c2 = ws.cell(ligne, 4, formule)
        c2.number_format = '#,##0.00 €'
        c2.fill = fill(COULEURS["depense_clair"])
        c2.border = style_bordure(fin=True)
        c2.alignment = centrer()
        ligne += 1

def ecrire_onglet_synthese(wb: openpyxl.Workbook):
    """Recrée l'onglet '📈 Synthèse' en lisant les totaux de chaque onglet annuel."""

    NOM_SYNTHESE = "📈 Synthèse"
    if NOM_SYNTHESE in wb.sheetnames:
        del wb[NOM_SYNTHESE]

    ws = wb.create_sheet(NOM_SYNTHESE, 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    onglets_annuels = sorted(
        [n for n in wb.sheetnames if n.startswith("Comptes ")],
        key=lambda n: int(n.split()[-1])
    )
    annees = [int(n.split()[-1]) for n in onglets_annuels]

    def trouver_cellules_recap(ws_annuel):
        """Retourne (cell_total_rec, cell_total_dep) en cherchant les libellés."""
        cell_rec = cell_dep = None
        for row in ws_annuel.iter_rows():
            for cell in row:
                if cell.value == "TOTAL RECETTES":
                    cell_rec = ws_annuel.cell(cell.row, 5)
                if cell.value == "TOTAL DÉPENSES":
                    cell_dep = ws_annuel.cell(cell.row, 5)
        return cell_rec, cell_dep

    # Largeurs
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 16

    ligne = 1

    # Titre
    ws.merge_cells(f"A{ligne}:F{ligne}")
    c = ws.cell(ligne, 1, "📈  SYNTHÈSE PLURIANNUELLE")
    c.font = police(gras=True, taille=16, couleur="FFFFFF")
    c.fill = fill(COULEURS["synthese"])
    c.alignment = centrer()
    c.border = style_bordure()
    ws.row_dimensions[ligne].height = 34
    ligne += 1

    # Sous-titre date de mise à jour
    ws.merge_cells(f"A{ligne}:F{ligne}")
    c = ws.cell(ligne, 1, f"Mis à jour le {datetime.now().strftime('%d/%m/%Y à %H:%M')}")
    c.font = police(italique=True, taille=10, couleur="FFFFFF")
    c.fill = fill(COULEURS["synthese"])
    c.alignment = centrer()
    ws.row_dimensions[ligne].height = 16
    ligne += 1

    # En-têtes
    entetes_syn = ["Année", "Total recettes", "Total dépenses", "Solde annuel", "Solde cumulé", "Évolution"]
    couleurs_ent = [
        COULEURS["synthese"], COULEURS["recette"], COULEURS["depense"],
        COULEURS["recap"], COULEURS["recap"], COULEURS["titre"],
    ]
    for col, (titre, coul) in enumerate(zip(entetes_syn, couleurs_ent), 1):
        c = ws.cell(ligne, col, titre)
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(coul)
        c.alignment = centrer()
        c.border = style_bordure()
    ws.row_dimensions[ligne].height = 22
    ligne += 1

    # Données par année
    ligne_debut_donnees = ligne

    for idx, (annee, nom_onglet) in enumerate(zip(annees, onglets_annuels)):
        ws_ann = wb[nom_onglet]
        cell_rec, cell_dep = trouver_cellules_recap(ws_ann)

        nom_safe = nom_onglet.replace("'", "''")
        if cell_rec and cell_dep:
            ref_rec = f"'{nom_safe}'!E{cell_rec.row}"
            ref_dep = f"'{nom_safe}'!E{cell_dep.row}"
        else:
            ref_rec = ref_dep = "0"

        bg = COULEURS["blanc"] if idx % 2 == 0 else COULEURS["alternance"]

        # Année
        c = ws.cell(ligne, 1, annee)
        c.font = police(gras=True)
        c.fill = fill(COULEURS["synthese_clair"])
        c.alignment = centrer()
        c.border = style_bordure()

        # Total recettes
        c = ws.cell(ligne, 2, f"={ref_rec}")
        c.number_format = '#,##0.00 €'
        c.font = police(couleur=COULEURS["recette"])
        c.fill = fill(bg)
        c.alignment = centrer()
        c.border = style_bordure(fin=True)

        # Total dépenses
        c = ws.cell(ligne, 3, f"={ref_dep}")
        c.number_format = '#,##0.00 €'
        c.font = police(couleur=COULEURS["depense"])
        c.fill = fill(bg)
        c.alignment = centrer()
        c.border = style_bordure(fin=True)

        # Solde annuel
        c = ws.cell(ligne, 4, f"=B{ligne}-C{ligne}")
        c.number_format = '#,##0.00 €'
        c.font = police(gras=True)
        c.fill = fill(bg)
        c.alignment = centrer()
        c.border = style_bordure(fin=True)

        # Solde cumulé
        if ligne == ligne_debut_donnees:
            formule_cumul = f"=D{ligne}"
        else:
            formule_cumul = f"=E{ligne-1}+D{ligne}"
        c = ws.cell(ligne, 5, formule_cumul)
        c.number_format = '#,##0.00 €'
        c.font = police(gras=True)
        c.fill = fill(bg)
        c.alignment = centrer()
        c.border = style_bordure(fin=True)

        # Évolution (vs année précédente)
        if ligne == ligne_debut_donnees:
            c = ws.cell(ligne, 6, "—")
            c.alignment = centrer()
            c.fill = fill(bg)
            c.border = style_bordure(fin=True)
        else:
            c = ws.cell(ligne, 6, f'=IFERROR((D{ligne}-D{ligne-1})/ABS(D{ligne-1}),"")') 
            c.number_format = '+0.0%;-0.0%;0.0%'
            c.font = police(gras=True)
            c.fill = fill(bg)
            c.alignment = centrer()
            c.border = style_bordure(fin=True)

        ws.row_dimensions[ligne].height = 20
        ligne += 1

    ligne_fin_donnees = ligne - 1

    # Ligne totaux / moyennes
    ligne += 1
    ws.merge_cells(f"A{ligne}:A{ligne}")
    c = ws.cell(ligne, 1, "TOTAUX")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["synthese"])
    c.alignment = centrer()
    c.border = style_bordure()

    for col, formule in [
        (2, f"=SUM(B{ligne_debut_donnees}:B{ligne_fin_donnees})"),
        (3, f"=SUM(C{ligne_debut_donnees}:C{ligne_fin_donnees})"),
        (4, f"=SUM(D{ligne_debut_donnees}:D{ligne_fin_donnees})"),
        (5, f"=E{ligne_fin_donnees}"),
    ]:
        c = ws.cell(ligne, col, formule)
        c.number_format = '#,##0.00 €'
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(COULEURS["synthese"])
        c.alignment = centrer()
        c.border = style_bordure()

    ws.row_dimensions[ligne].height = 22
    ligne += 1

    # Moyennes
    ws.merge_cells(f"A{ligne}:A{ligne}")
    c = ws.cell(ligne, 1, "MOYENNES")
    c.font = police(gras=True, couleur="FFFFFF")
    c.fill = fill(COULEURS["titre"])
    c.alignment = centrer()
    c.border = style_bordure()

    for col, formule in [
        (2, f"=AVERAGE(B{ligne_debut_donnees}:B{ligne_fin_donnees})"),
        (3, f"=AVERAGE(C{ligne_debut_donnees}:C{ligne_fin_donnees})"),
        (4, f"=AVERAGE(D{ligne_debut_donnees}:D{ligne_fin_donnees})"),
    ]:
        c = ws.cell(ligne, col, formule)
        c.number_format = '#,##0.00 €'
        c.font = police(gras=True, couleur="FFFFFF")
        c.fill = fill(COULEURS["titre"])
        c.alignment = centrer()
        c.border = style_bordure()

    ws.row_dimensions[ligne].height = 22
    ligne += 2

    # Graphique
    if annees:
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "clustered"
        bar.title = "Recettes & Dépenses par année"
        bar.y_axis.title = "Montant (€)"
        bar.x_axis.title = "Année"
        bar.style = 10
        bar.width = 20
        bar.height = 12

        data_rec = Reference(ws, min_col=2, max_col=2, 
                           min_row=ligne_debut_donnees - 1, max_row=ligne_fin_donnees)
        data_dep = Reference(ws, min_col=3, max_col=3,
                           min_row=ligne_debut_donnees - 1, max_row=ligne_fin_donnees)
        cats = Reference(ws, min_col=1,
                        min_row=ligne_debut_donnees, max_row=ligne_fin_donnees)

        bar.add_data(data_rec, titles_from_data=True)
        bar.add_data(data_dep, titles_from_data=True)
        bar.set_categories(cats)
        bar.series[0].graphicalProperties.solidFill = COULEURS["recette"]
        bar.series[1].graphicalProperties.solidFill = COULEURS["depense"]

        ws.add_chart(bar, f"A{ligne}")

        # LineChart solde cumulé
        line = LineChart()
        line.title = "Évolution du solde cumulé"
        line.y_axis.title = "Solde cumulé (€)"
        line.x_axis.title = "Année"
        line.style = 10
        line.width = 20
        line.height = 12

        data_cumul = Reference(ws, min_col=5, max_col=5,
                             min_row=ligne_debut_donnees - 1, max_row=ligne_fin_donnees)
        line.add_data(data_cumul, titles_from_data=True)
        line.set_categories(cats)
        line.series[0].graphicalProperties.line.solidFill = COULEURS["synthese"]
        line.series[0].graphicalProperties.line.width = 25000

        ws.add_chart(line, f"D{ligne}")

def generer_ou_mettre_a_jour(recettes, depenses, output: Path, annee: int, force: bool):
    NOM_ONGLET = f"Comptes {annee}"
    NOM_SYNTHESE = "📈 Synthèse"

    if output.exists():
        print(f"📂 Fichier existant détecté : {output.name}")
        wb = openpyxl.load_workbook(output)

        if NOM_ONGLET in wb.sheetnames:
            if not force:
                rep = input(f"⚠️  L'onglet '{NOM_ONGLET}' existe déjà. Le remplacer ? [o/N] ").strip().lower()
                if rep not in ("o", "oui", "y", "yes"):
                    print("❌ Opération annulée.")
                    return
            del wb[NOM_ONGLET]
            print(f"🗑️  Onglet '{NOM_ONGLET}' supprimé.")
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        print(f"🆕 Création du fichier : {output.name}")

    ws = wb.create_sheet(NOM_ONGLET)
    ecrire_onglet_annuel(ws, recettes, depenses, annee)
    print(f"✅ Onglet '{NOM_ONGLET}' créé.")

    onglets_annuels = sorted([n for n in wb.sheetnames if n.startswith("Comptes ")],
                            key=lambda n: int(n.split()[-1]))
    ordre = [NOM_SYNTHESE] + onglets_annuels if NOM_SYNTHESE in wb.sheetnames else onglets_annuels
    wb._sheets.sort(key=lambda s: ordre.index(s.title) if s.title in ordre else 99)

    try:
        ecrire_onglet_synthese(wb)
        print(f"📈 Onglet '{NOM_SYNTHESE}' mis à jour.")
    except Exception as e:
        print(f"⚠️ Erreur lors de la création de la synthèse: {e}")
        print("Le fichier sera sauvegardé sans la synthèse.")

    try:
        wb.save(output)
        print(f"💾 Fichier sauvegardé : {output}")
    except Exception as e:
        print(f"❌ Erreur lors de la sauvegarde: {e}")
        sys.exit(1)

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

    print(f"\n📒 Livre de comptes — Année {args.annee}")
    print("=" * 50)

    pdfs_valides = []
    for p in args.pdfs:
        path = Path(p)
        if not path.exists():
            print(f"⚠️  Fichier introuvable : {p}")
        elif path.suffix.lower() != ".pdf":
            print(f"⚠️  Pas un PDF : {p}")
        else:
            pdfs_valides.append(path)

    if not pdfs_valides:
        print("❌ Aucun fichier PDF valide trouvé.")
        sys.exit(1)

    print(f"\n🔍 Extraction depuis {len(pdfs_valides)} fichier(s)...")
    all_recettes, all_depenses = [], []
    for pdf in sorted(pdfs_valides):
        print(f"  📄 {pdf.name}...")
        try:
            rec, dep = extraire_operations_pdf_camelot(pdf, args.annee)
            print(f"     → {len(rec)} recette(s), {len(dep)} dépense(s)")
            all_recettes.extend(rec)
            all_depenses.extend(dep)
        except Exception as e:
            print(f"     ❌ Erreur lors de l'extraction du PDF {pdf.name}: {e}")
            continue

    def cle_tri(op):
        try:
            return datetime.strptime(op[0], "%d/%m/%Y")
        except ValueError:
            return datetime.min

    all_recettes.sort(key=cle_tri)
    all_depenses.sort(key=cle_tri)

    print(f"\n📊 Total : {len(all_recettes)} recettes, {len(all_depenses)} dépenses")

    generer_ou_mettre_a_jour(all_recettes, all_depenses, args.output, args.annee, args.force)

if __name__ == "__main__":
    main()