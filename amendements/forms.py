from django import forms
from .models import Amendement
from .models import AmendementFichier

class UploadWordForm(forms.Form):
    fichier = forms.FileField(label="Fichier Word (.docx)")

class AmendementForm(forms.ModelForm):
    class Meta:
        model = Amendement
        exclude = ['date_insertion', 'utilisateur', 'fichier']

class AmendementFichierForm(forms.ModelForm):
    class Meta:
        model = AmendementFichier
        fields = ['fichier'] 