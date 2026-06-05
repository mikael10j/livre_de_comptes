#!/usr/bin/env python3
"""
Tests d'intégration avec les vrais fichiers PDF.
Ces tests utilisent les fichiers PDF réels présents dans le répertoire parent.
"""

import unittest
import tempfile
import shutil
import sys
from pathlib import Path
import openpyxl

# Ajouter le répertoire parent au path pour l'import
sys.path.append(str(Path(__file__).parent.parent))
import livre_comptes as lc


class TestIntegrationReal(unittest.TestCase):
    """Tests d'intégration avec les vrais fichiers PDF."""

    def setUp(self):
        """Configuration des tests d'intégration."""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
        # Trouver les fichiers PDF réels dans le répertoire parent
        parent_dir = Path(__file__).parent.parent
        self.pdf_files = list(parent_dir.glob("*.pdf"))
        
    def tearDown(self):
        """Nettoyage après tests."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_extraction_tous_pdfs_reels(self):
        """Test d'extraction avec tous les PDFs réels disponibles."""
        if not self.pdf_files:
            self.skipTest("Aucun fichier PDF trouvé pour les tests")
        
        total_operations = 0
        total_recettes = 0
        total_depenses = 0
        
        for pdf_file in self.pdf_files:
            with self.subTest(pdf=pdf_file.name):
                # Déterminer l'année selon le nom du fichier
                annee = 2023 if "2023" in pdf_file.name else 2019
                
                recettes, depenses = lc.extraire_operations_pdf_camelot(pdf_file, annee)
                
                # Vérifications de base
                self.assertIsInstance(recettes, list)
                self.assertIsInstance(depenses, list)
                
                # Chaque opération doit avoir la bonne structure
                for recette in recettes:
                    self.assertEqual(len(recette), 7)  # 7 champs par opération
                    self.assertIsInstance(recette[4], (int, float))  # Montant numérique
                    self.assertGreater(recette[4], 0)  # Montant positif
                
                for depense in depenses:
                    self.assertEqual(len(depense), 7)
                    self.assertIsInstance(depense[4], (int, float))
                    self.assertGreater(depense[4], 0)
                
                total_operations += len(recettes) + len(depenses)
                total_recettes += len(recettes)
                total_depenses += len(depenses)
        
        print(f"\n📊 Extraction réelle: {total_operations} opérations "
              f"({total_recettes} recettes, {total_depenses} dépenses)")
        
        # Vérifier qu'au moins quelques opérations ont été extraites
        self.assertGreater(total_operations, 0, "Aucune opération extraite des PDFs réels")

    def test_generation_excel_complete(self):
        """Test de génération complète d'un fichier Excel avec tous les PDFs."""
        if not self.pdf_files:
            self.skipTest("Aucun fichier PDF trouvé pour les tests")
        
        output_path = self.temp_path / "test_integration_complete.xlsx"
        
        # Traiter les PDFs par année
        annees_traitees = set()
        
        for pdf_file in self.pdf_files:
            annee = 2023 if "2023" in pdf_file.name else 2019
            
            if annee in annees_traitees:
                continue  # Éviter de traiter plusieurs fois la même année
                
            recettes, depenses = lc.extraire_operations_pdf_camelot(pdf_file, annee)
            
            if recettes or depenses:  # Seulement si des opérations ont été trouvées
                lc.generer_ou_mettre_a_jour(recettes, depenses, output_path, annee, force=True)
                annees_traitees.add(annee)
        
        # Vérifier le fichier généré
        self.assertTrue(output_path.exists())
        
        # Vérifier le contenu Excel
        wb = openpyxl.load_workbook(output_path)
        
        # Doit contenir au moins un onglet annuel
        onglets_comptes = [sheet for sheet in wb.sheetnames if sheet.startswith("Comptes ")]
        self.assertGreater(len(onglets_comptes), 0, "Aucun onglet de comptes créé")
        
        # Doit contenir l'onglet synthèse
        self.assertIn("📈 Synthèse", wb.sheetnames)
        
        print(f"\n📊 Excel généré avec {len(onglets_comptes)} onglet(s) : {onglets_comptes}")

    def test_coherence_donnees_extraites(self):
        """Test de cohérence des données extraites."""
        if not self.pdf_files:
            self.skipTest("Aucun fichier PDF trouvé pour les tests")
        
        for pdf_file in self.pdf_files:
            with self.subTest(pdf=pdf_file.name):
                annee = 2023 if "2023" in pdf_file.name else 2019
                recettes, depenses = lc.extraire_operations_pdf_camelot(pdf_file, annee)
                
                # Vérifier la cohérence des dates
                for operation in recettes + depenses:
                    date_ope, date_val, libelle, ref, montant, categorie, source = operation
                    
                    # Les dates doivent contenir l'année
                    self.assertIn(str(annee), date_ope)
                    self.assertIn(str(annee), date_val)
                    
                    # Le libellé ne doit pas être vide
                    self.assertGreater(len(libelle.strip()), 0)
                    
                    # Le montant doit être positif
                    self.assertGreater(montant, 0)
                    
                    # La catégorie doit être valide
                    if operation in recettes:
                        self.assertIn(categorie, lc.CATEGORIES_RECETTES)
                    else:
                        self.assertIn(categorie, lc.CATEGORIES_DEPENSES)
                    
                    # La source doit correspondre au fichier
                    self.assertEqual(source, pdf_file.stem)

    def test_reproductibilite(self):
        """Test de reproductibilité : même PDF doit donner mêmes résultats."""
        if not self.pdf_files:
            self.skipTest("Aucun fichier PDF trouvé pour les tests")
        
        pdf_test = self.pdf_files[0]  # Prendre le premier PDF
        annee = 2023 if "2023" in pdf_test.name else 2019
        
        # Extraire deux fois
        recettes1, depenses1 = lc.extraire_operations_pdf_camelot(pdf_test, annee)
        recettes2, depenses2 = lc.extraire_operations_pdf_camelot(pdf_test, annee)
        
        # Les résultats doivent être identiques
        self.assertEqual(len(recettes1), len(recettes2))
        self.assertEqual(len(depenses1), len(depenses2))
        
        # Comparer operation par operation
        for op1, op2 in zip(recettes1, recettes2):
            self.assertEqual(op1, op2)
            
        for op1, op2 in zip(depenses1, depenses2):
            self.assertEqual(op1, op2)
        
        print(f"✅ Reproductibilité confirmée pour {pdf_test.name}")

    def test_gestion_erreurs_fichiers(self):
        """Test de gestion d'erreurs avec des fichiers problématiques."""
        # Tester avec un fichier inexistant
        pdf_inexistant = Path("fichier_inexistant.pdf")
        recettes, depenses = lc.extraire_operations_pdf_camelot(pdf_inexistant, 2023)
        
        # Doit retourner des listes vides sans lever d'exception
        self.assertEqual(recettes, [])
        self.assertEqual(depenses, [])
        
        # Tester avec un fichier vide
        fichier_vide = self.temp_path / "vide.pdf"
        fichier_vide.touch()
        
        recettes, depenses = lc.extraire_operations_pdf_camelot(fichier_vide, 2023)
        self.assertEqual(recettes, [])
        self.assertEqual(depenses, [])


def run_integration_tests():
    """Exécute les tests d'intégration avec les vrais fichiers."""
    print("🔗 TESTS D'INTÉGRATION")
    print("=" * 50)
    
    # Vérifier la présence de fichiers PDF
    parent_dir = Path(__file__).parent.parent
    pdf_files = list(parent_dir.glob("*.pdf"))
    if not pdf_files:
        print("⚠️  Aucun fichier PDF trouvé. Les tests d'intégration seront ignorés.")
        return True
    
    print(f"📄 {len(pdf_files)} fichier(s) PDF trouvé(s): {[f.name for f in pdf_files]}")
    
    suite = unittest.TestLoader().loadTestsFromTestCase(TestIntegrationReal)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print(f"\n📊 Résultats: {result.testsRun - len(result.failures) - len(result.errors)}/{result.testsRun} tests réussis")
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_integration_tests()
    exit(0 if success else 1)