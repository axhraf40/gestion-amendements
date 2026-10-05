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
    texte_original = models.TextField(blank=True, null=True)
    texte_modifie = models.TextField(blank=True, null=True)
    justification = models.TextField(blank=True, null=True)
    texte_entete = models.TextField(blank=True, null=True, help_text="Texte d'en-tête brut de l'amendement (sert à la détection IA)")
    numero_groupe = models.CharField(max_length=100, blank=True, null=True)
    date_insertion = models.DateTimeField(auto_now_add=True)
    utilisateur = models.ForeignKey(User, on_delete=models.CASCADE)
    fichier = models.ForeignKey('AmendementFichier', on_delete=models.CASCADE, null=True, blank=True)
    champs_detectes = models.JSONField(blank=True, null=True, help_text="Champs détectés automatiquement (IA + règles)")

    def __str__(self):
        return f"Amendement {self.numero_modification} - Article {self.numero_article}"

class AmendementFichier(models.Model):
    utilisateur = models.ForeignKey(User, on_delete=models.CASCADE)
    fichier = models.FileField(upload_to='amendements_uploads/')
    date_upload = models.DateTimeField(auto_now_add=True)
    extracted_content = models.TextField(blank=True, null=True)  # Contenu extrait du fichier Word (HTML)
    nom_personnalise = models.CharField(max_length=255, blank=True, null=True, help_text="Nom personnalisé du fichier (optionnel)")

    def __str__(self):
        return f"{self.nom_personnalise or self.fichier.name} ({self.utilisateur.username})"
