from django.core.management.base import BaseCommand
from amendements.models import Amendement

class Command(BaseCommand):
    help = 'Supprime tous les amendements non liés à un fichier.'

    def handle(self, *args, **kwargs):
        count = Amendement.objects.filter(fichier__isnull=True).delete()[0]
        self.stdout.write(self.style.SUCCESS(f'{count} amendements non liés supprimés.')) 