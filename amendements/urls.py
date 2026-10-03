from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_amendement, name='upload_amendement'),
    path('valider/', views.valider_amendements, name='valider_amendements'),
    path('fichiers/', views.fichiers_amendements, name='fichiers_amendements'),
    path('fichiers/upload/', views.upload_fichier_amendement, name='upload_fichier_amendement'),
    path('fichiers/<int:fichier_id>/supprimer/', views.supprimer_fichier_amendement, name='supprimer_fichier_amendement'),
    path('fichiers/<int:fichier_id>/detail/', views.detail_fichier_amendement, name='detail_fichier_amendement'),
    path('fichiers/<int:fichier_id>/modifier/', views.modifier_fichier_extrait, name='modifier_fichier_extrait'),
    path('fichier_mammoth/<int:fichier_id>/', views.afficher_fichier_mammoth, name='fichier_mammoth'),
    path('fichiers/<int:fichier_id>/export_pdf/', views.export_pdf_amendement, name='export_pdf_amendement'),
] 