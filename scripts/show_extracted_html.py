import os
import django
import sys

# Initialisation Django
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'charge_amendements.settings')
django.setup()

from amendements.models import AmendementFichier

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/show_extracted_html.py <fichier_id>")
        sys.exit(1)
    fichier_id = int(sys.argv[1])
    try:
        fichier_obj = AmendementFichier.objects.get(id=fichier_id)
        print("Contenu extrait (extracted_content) :\n")
        print(fichier_obj.extracted_content)
    except AmendementFichier.DoesNotExist:
        print(f"Aucun fichier avec l'id {fichier_id}") 