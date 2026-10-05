from django.contrib import admin
from .models import AmendementFichier, Amendement
from django.utils.safestring import mark_safe
import csv
from django.http import HttpResponse

# Suppression de l'inline Amendement et retour à la configuration précédente

@admin.register(AmendementFichier)
class AmendementFichierAdmin(admin.ModelAdmin):
    list_display = ('id', 'fichier', 'utilisateur', 'date_upload')
    search_fields = ('fichier', 'utilisateur__username')
    list_filter = ('utilisateur',)

@admin.register(Amendement)
class AmendementAdmin(admin.ModelAdmin):
    list_display = (
        'numero_modification', 'type_modification', 'nom_loi', 'article',
        'alinea', 'numero_article', 'texte_original', 'texte_modifie',
        'justification', 'numero_groupe', 'fichier', 'utilisateur', 'has_detected_fields'
    )
    list_filter = ('fichier', 'nom_loi', 'numero_modification', 'numero_groupe')
    search_fields = ('numero_modification', 'texte_modifie', 'texte_original', 'justification', 'numero_groupe')
    list_editable = ('type_modification', 'nom_loi', 'article', 'alinea', 'numero_article', 'texte_original', 'texte_modifie', 'justification', 'numero_groupe')
    readonly_fields = ('champs_detectes',)
    actions = ['export_as_csv', 're_detect_fields']
    
    def has_detected_fields(self, obj):
        """Affiche si des champs ont été détectés"""
        if obj.champs_detectes:
            return mark_safe('<span style="color: green;">✓</span>')
        return mark_safe('<span style="color: red;">✗</span>')
    has_detected_fields.short_description = 'Champs détectés'
    
    def re_detect_fields(self, request, queryset):
        """Re-détecte les champs pour les amendements sélectionnés"""
        from .views import _appliquer_detection

        updated_count = 0
        for amendement in queryset:
            _appliquer_detection(amendement)
            updated_count += 1
        
        self.message_user(request, f"{updated_count} amendement(s) mis à jour avec les champs détectés.")
    re_detect_fields.short_description = "Re-détecter les champs avec l'IA"

    def export_as_csv(self, request, queryset):
        meta = self.model._meta
        field_names = [field.name for field in meta.fields]
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename=amendements.csv'
        writer = csv.writer(response)
        writer.writerow(field_names)
        for obj in queryset:
            writer.writerow([getattr(obj, field) for field in field_names])
        return response
    export_as_csv.short_description = "Exporter la sélection en CSV"
