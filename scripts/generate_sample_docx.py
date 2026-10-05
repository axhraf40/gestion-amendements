"""Génère un document d'amendements fictif pour essayer l'application.

Usage :
    python scripts/generate_sample_docx.py [sortie.docx]

Tout le contenu est inventé (voir amendements/sample_docx.py).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from amendements.sample_docx import enregistrer  # noqa: E402

if __name__ == "__main__":
    sortie = sys.argv[1] if len(sys.argv) > 1 else "exemple_amendements.docx"
    enregistrer(sortie)
    print(f"Document fictif créé : {sortie}")
