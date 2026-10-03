from django import forms
from .models import Amendement
from .models import AmendementFichier

from django.core.validators import FileExtensionValidator

DOCX_ONLY = [FileExtensionValidator(allowed_extensions=['docx'])]


class UploadWordForm(forms.Form):
    fichier = forms.FileField(label="Fichier Word (.docx)", validators=DOCX_ONLY)

class AmendementForm(forms.ModelForm):
    class Meta:
        model = Amendement
        exclude = ['date_insertion', 'utilisateur', 'fichier']

class AmendementFichierForm(forms.ModelForm):
    class Meta:
        model = AmendementFichier
        fields = ['fichier']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['fichier'].validators += DOCX_ONLY 