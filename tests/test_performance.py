#!/usr/bin/env python3
"""
Tests de performance pour le script livre_comptes.py
"""

import unittest
import time
import tempfile
import shutil
import sys
from pathlib import Path
from unittest.mock import Mock, patch
import pandas as pd

# Ajouter le répertoire parent au path pour l'import
sys.path.append(str(Path(__file__).parent.parent))
import livre_comptes as lc


class TestPerformance(unittest.TestCase):
    """Tests de performance et de robustesse."""

    def setUp(self):
        """Configuration des tests de performance."""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)

    def tearDown(self):
        """Nettoyage après tests."""
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_performance_parse_montant(self):
        """Test de performance du parsing de montants."""
        test_values = ["123,45"] * 10000
        
        start_time = time.time()
        results = [lc.parse_montant(val) for val in test_values]
        end_time = time.time()
        
        # Doit traiter 10k valeurs en moins d'1 seconde
        execution_time = end_time - start_time
        self.assertLess(execution_time, 1.0, 
                       f"Parsing trop lent: {execution_time:.3f}s pour 10k valeurs")
        
        # Vérifier la cohérence des résultats
        self.assertEqual(len(set(results)), 1)  # Tous identiques
        self.assertEqual(results[0], 123.45)

    def test_performance_categorisation(self):
        """Test de performance de la catégorisation."""
        test_libelles = [
            "Cotisation membre 2023", "Rem Chq 123456", "Retrait GAB",
            "Assurance SMACL", "Arbitre match", "Location salle"
        ] * 1000
        
        start_time = time.time()
        results_rec = [lc.categoriser(lib, lc.REGLES_RECETTES) for lib in test_libelles]
        results_dep = [lc.categoriser(lib, lc.REGLES_DEPENSES) for lib in test_libelles]
        end_time = time.time()
        
        execution_time = end_time - start_time
        self.assertLess(execution_time, 1.0,
                       f"Catégorisation trop lente: {execution_time:.3f}s pour 6k opérations")

    @patch('livre_comptes.camelot')
    def test_performance_extraction_large_dataset(self, mock_camelot):
        """Test avec un grand dataset simulé."""
        # Créer un grand DataFrame (1000 lignes d'opérations)
        large_data = []
        large_data.extend([["", "", "Libellé des opérations", "", "Débit", "Crédit"]])
        large_data.extend([["", "", "", "", "", ""]])
        
        for i in range(1000):
            if i % 2 == 0:
                large_data.append([f"{i%30+1:02d}.01", f"{i%30+1:02d}.01", f"Opération {i}", "", f"{i+100},00", ""])
            else:
                large_data.append([f"{i%30+1:02d}.01", f"{i%30+1:02d}.01", f"Recette {i}", "", "", f"{i+50},50"])
        
        df_large = pd.DataFrame(large_data)
        
        mock_table = Mock()
        mock_table.df = df_large
        mock_camelot.read_pdf.return_value = [mock_table]
        
        pdf_path = Path("large_test.pdf")
        
        start_time = time.time()
        recettes, depenses = lc.extraire_operations_pdf_camelot(pdf_path, 2023)
        end_time = time.time()
        
        execution_time = end_time - start_time
        self.assertLess(execution_time, 5.0,
                       f"Extraction trop lente: {execution_time:.3f}s pour 1000 opérations")
        
        # Vérifier que toutes les opérations ont été extraites
        total_operations = len(recettes) + len(depenses)
        self.assertEqual(total_operations, 1000)

    def test_robustesse_donnees_corrompues(self):
        """Test de robustesse avec des données corrompues."""
        test_cases_corrompus = [
            "999999999999999999,99",  # Montant énorme
            "abc,def",                # Lettres au lieu de chiffres
            "123,456,789",           # Format invalide
            "123.456.789,00",        # Trop de points
            "",                      # Chaîne vide
            "   ",                   # Espaces seulement
            "NULL",                  # Valeur NULL
            "undefined",             # Valeur undefined
        ]
        
        for test_case in test_cases_corrompus:
            with self.subTest(input_corrompu=test_case):
                # Ne doit pas lever d'exception
                try:
                    result = lc.parse_montant(test_case)
                    # Doit retourner 0.0 ou une valeur numérique valide
                    self.assertTrue(isinstance(result, (int, float)))
                except Exception as e:
                    self.fail(f"Exception inattendue pour '{test_case}': {e}")

    def test_robustesse_regex_complexes(self):
        """Test de robustesse des expressions régulières."""
        libelles_complexes = [
            "Très long libellé avec beaucoup de mots et de caractères spéciaux éàùêç!@#$%^&*()",
            "Libellé\navec\nsauts\nde\nligne",
            "Libellé\tavec\ttabulations",
            "Libellé avec émojis 🏀⚽🎾",
            "LibelléAvecAccentsÀÉÈÊÇÙÛÎÔ",
            "LIBELLÉ EN MAJUSCULES AVEC ESPACES   MULTIPLES",
            "libellé en minuscules",
            "",  # Libellé vide
        ]
        
        for libelle in libelles_complexes:
            with self.subTest(libelle=libelle[:30] + "..."):
                # Ne doit pas lever d'exception
                try:
                    cat_rec = lc.categoriser(libelle, lc.REGLES_RECETTES)
                    cat_dep = lc.categoriser(libelle, lc.REGLES_DEPENSES) 
                    
                    # Doit retourner des catégories valides
                    self.assertIn(cat_rec, lc.CATEGORIES_RECETTES)
                    self.assertIn(cat_dep, lc.CATEGORIES_DEPENSES)
                except Exception as e:
                    self.fail(f"Exception inattendue pour libellé complexe: {e}")

    def test_limites_memoire(self):
        """Test des limites mémoire avec de grandes structures."""
        # Créer de grandes listes d'opérations
        grandes_recettes = []
        grandes_depenses = []
        
        for i in range(10000):
            recette = [f"01/01/2023", f"01/01/2023", f"Recette {i}", f"REF{i:06d}", 
                      float(i), "Cotisations", "test.pdf"]
            depense = [f"02/01/2023", f"02/01/2023", f"Dépense {i}", f"REF{i+10000:06d}", 
                      float(i), "Autres", "test.pdf"]
            
            grandes_recettes.append(recette)
            grandes_depenses.append(depense)
        
        # Tester la génération Excel avec ces grandes listes
        output_path = self.temp_path / "test_memoire.xlsx"
        
        start_time = time.time()
        try:
            lc.generer_ou_mettre_a_jour(grandes_recettes, grandes_depenses, 
                                        output_path, 2023, force=True)
            end_time = time.time()
            
            # Vérifier que le fichier a été créé
            self.assertTrue(output_path.exists())
            
            execution_time = end_time - start_time
            print(f"Génération de 20k opérations: {execution_time:.2f}s")
            
        except MemoryError:
            self.skipTest("Pas assez de mémoire pour ce test")
        except Exception as e:
            self.fail(f"Erreur lors du test de mémoire: {e}")


def run_performance_tests():
    """Exécute les tests de performance."""
    print("🏃‍♂️ TESTS DE PERFORMANCE")
    print("=" * 50)
    
    suite = unittest.TestLoader().loadTestsFromTestCase(TestPerformance)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print(f"\n📊 Résultats: {result.testsRun - len(result.failures) - len(result.errors)}/{result.testsRun} tests réussis")
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_performance_tests()
    exit(0 if success else 1)