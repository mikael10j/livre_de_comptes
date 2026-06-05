#!/usr/bin/env python3
"""
Script principal pour exécuter tous les tests du projet livre de comptes.
Les tests sont maintenant organisés dans le répertoire tests/.
"""

import sys
import time
from pathlib import Path

# Ajouter le répertoire des tests au path
tests_dir = Path(__file__).parent / "tests"
sys.path.insert(0, str(tests_dir))

# Import des modules de test
try:
    import test_unitaires
    import test_performance  
    import test_integration
except ImportError as e:
    print(f"❌ Erreur d'import des modules de test: {e}")
    print("💡 Vérifiez que le répertoire tests/ existe et contient les fichiers de test")
    sys.exit(1)


def print_header(title):
    """Affiche un en-tête formaté."""
    print(f"\n{'='*60}")
    print(f"🧪 {title}")
    print('='*60)


def print_summary(results):
    """Affiche un résumé des résultats de tests."""
    total_tests = sum(results.values())
    passed_tests = results.get('passed', 0)
    failed_tests = results.get('failed', 0)
    
    print(f"\n🎯 RÉSUMÉ GLOBAL DES TESTS")
    print('='*60)
    print(f"✅ Tests réussis:    {passed_tests:3d}")
    print(f"❌ Tests échoués:    {failed_tests:3d}")
    print(f"📊 Total:            {total_tests:3d}")
    
    success_rate = (passed_tests / total_tests * 100) if total_tests > 0 else 0
    print(f"📈 Taux de réussite: {success_rate:5.1f}%")
    
    if failed_tests == 0:
        print("\n🎉 TOUS LES TESTS SONT RÉUSSIS !")
    else:
        print(f"\n⚠️  {failed_tests} test(s) nécessitent une attention")


def main():
    """Fonction principale pour exécuter tous les tests."""
    print("🧪 SUITE DE TESTS COMPLÈTE - LIVRE DE COMPTES")
    print("=" * 60)
    print("Tests organisés dans le répertoire tests/")
    print("• Tests unitaires (fonctions individuelles)")  
    print("• Tests de performance (vitesse et robustesse)")
    print("• Tests d'intégration (avec vrais fichiers PDF)")
    print()
    
    # Vérifier la présence du module principal
    try:
        import livre_comptes
        print("✅ Module livre_comptes importé avec succès")
    except ImportError as e:
        print(f"❌ Impossible d'importer livre_comptes: {e}")
        return False
    
    # Vérifier la présence des dépendances
    missing_deps = []
    try:
        import pandas
        import openpyxl 
        import camelot
    except ImportError as e:
        missing_deps.append(str(e))
    
    if missing_deps:
        print("❌ Dépendances manquantes:")
        for dep in missing_deps:
            print(f"   • {dep}")
        print("💡 Installez avec: pip install camelot-py[cv] pandas openpyxl")
        return False
    
    print("✅ Toutes les dépendances sont disponibles")
    
    # Statistiques globales
    total_start_time = time.time()
    results = {'passed': 0, 'failed': 0}
    
    # 1. Tests unitaires
    print_header("TESTS UNITAIRES")
    try:
        unit_success = test_unitaires.run_tests()
        if unit_success:
            results['passed'] += 18  # Nombre de tests unitaires
            print("✅ Tests unitaires: RÉUSSIS")
        else:
            results['failed'] += 18
            print("❌ Tests unitaires: ÉCHECS")
    except Exception as e:
        print(f"💥 Erreur lors des tests unitaires: {e}")
        results['failed'] += 18
    
    # 2. Tests de performance
    print_header("TESTS DE PERFORMANCE")
    try:
        perf_success = test_performance.run_performance_tests()
        if perf_success:
            results['passed'] += 6  # Nombre de tests de performance
            print("✅ Tests de performance: RÉUSSIS")
        else:
            results['failed'] += 6
            print("❌ Tests de performance: ÉCHECS")
    except Exception as e:
        print(f"💥 Erreur lors des tests de performance: {e}")
        results['failed'] += 6
    
    # 3. Tests d'intégration avec fichiers réels
    print_header("TESTS D'INTÉGRATION")
    pdf_files = list(Path(".").glob("*.pdf"))
    if pdf_files:
        try:
            integration_success = test_integration.run_integration_tests()
            if integration_success:
                results['passed'] += 6  # Nombre de tests d'intégration
                print("✅ Tests d'intégration: RÉUSSIS")
            else:
                results['failed'] += 6
                print("❌ Tests d'intégration: ÉCHECS")
        except Exception as e:
            print(f"💥 Erreur lors des tests d'intégration: {e}")
            results['failed'] += 6
    else:
        print("⚠️  Aucun fichier PDF trouvé - tests d'intégration ignorés")
        print("💡 Ajoutez des fichiers *.pdf pour tester l'extraction réelle")
    
    # Calcul du temps total
    total_end_time = time.time()
    total_duration = total_end_time - total_start_time
    
    # Résumé final
    print_summary(results)
    print(f"⏱️  Temps d'exécution total: {total_duration:.2f} secondes")
    
    # Information sur la structure des tests
    print(f"\n📁 Structure des tests:")
    print(f"   tests/")
    print(f"   ├── __init__.py")
    print(f"   ├── test_unitaires.py      (18 tests)")
    print(f"   ├── test_performance.py    (6 tests)")
    print(f"   └── test_integration.py    (6 tests)")
    
    # Code de sortie
    all_success = results['failed'] == 0
    if all_success:
        print("\n🚀 Le projet est prêt pour la production !")
    else:
        print("\n🔧 Quelques améliorations sont nécessaires.")
    
    return all_success


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)