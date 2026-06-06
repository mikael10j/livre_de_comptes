#!/usr/bin/env python3
"""
Génère un livre de comptes Excel à partir de relevés bancaires PDF.

Usage:
    python livre_comptes.py releve1.pdf releve2.pdf [...] [-o sortie.xlsx] [-a ANNEE] [-f]
"""

# This file now serves as a simple entry point to the modular application
# All functionality has been moved to separate modules
from main import main

if __name__ == "__main__":
    main()