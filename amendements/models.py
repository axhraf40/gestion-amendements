from django.db import models
from django.contrib.auth.models import User

class Amendement(models.Model):
    numero_modification = models.CharField(max_length=50)
    type_modification = models.CharField(max_length=255, blank=True, null=True)
    texte_amendement = models.TextField()
    numero_article = models.CharField(max_length=50)
    alinea = models.TextField(blank=True, null=True)
    article = models.CharField(max_length=50)
    nom_loi = models.CharField(max_length=255)
    texte_original = models.TextField()
    texte_modifie = models.TextField()
    justification = models.TextField()
    date_insertion = models.DateTimeField(auto_now_add=True)
    utilisateur = models.ForeignKey(User, on_delete=models.CASCADE)
    fichier = models.ForeignKey('AmendementFichier', on_delete=models.CASCADE, null=True, blank=True)

    def __str__(self):
        return f"Amendement {self.numero_modification} - Article {self.numero_article}"

class AmendementFichier(models.Model):
    utilisateur = models.ForeignKey(User, on_delete=models.CASCADE)
    fichier = models.FileField(upload_to='amendements_uploads/')
    date_upload = models.DateTimeField(auto_now_add=True)
    extracted_content = models.TextField(blank=True, null=True)  # Contenu extrait du fichier Word (HTML)

    def __str__(self):
        return f"{self.fichier.name} ({self.utilisateur.username})"
