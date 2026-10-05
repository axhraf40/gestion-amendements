from django.urls import path
from . import views
from .views import telecharger_pdf_libreoffice

urlpatterns = [
    path('upload/', views.upload_amendement, name='upload_amendement'),
    path('valider/', views.valider_amendements, name='valider_amendements'),
    path('fichiers/', views.fichiers_amendements, name='fichiers_amendements'),
    path('fichiers/upload/', views.upload_fichier_amendement, name='upload_fichier_amendement'),
    path('fichiers/<int:fichier_id>/supprimer/', views.supprimer_fichier_amendement, name='supprimer_fichier_amendement'),
    path('fichiers/<int:fichier_id>/detail/', views.detail_fichier_amendement, name='detail_fichier_amendement'),
    path('fichiers/<int:fichier_id>/modifier/', views.modifier_fichier_extrait, name='modifier_fichier_extrait'),
    path('fichier_mammoth/<int:fichier_id>/', views.afficher_fichier_mammoth, name='fichier_mammoth'),
    path('fichiers/<int:fichier_id>/apercu_pdf/', views.afficher_pdf_amendement, name='apercu_pdf_amendement'),
    path('fichiers/<int:fichier_id>/renommer/', views.renommer_fichier_amendement, name='renommer_fichier_amendement'),
    path('fichiers/supprimer_tous/', views.supprimer_tous_fichiers_amendement, name='supprimer_tous_fichiers_amendement'),
    path('amendements/<int:fichier_id>/telecharger_pdf_libreoffice/', telecharger_pdf_libreoffice, name='telecharger_pdf_libreoffice'),
    # Nouvelles URLs pour HuggingFace
    path('amendements/<int:amendement_id>/champs-detectes/', views.afficher_champs_detectes, name='afficher_champs_detectes'),
    path('amendements/<int:amendement_id>/re-detecter-champs/', views.re_detecter_champs, name='re_detecter_champs'),
    path('fichiers/<int:fichier_id>/statistiques-detection/', views.statistiques_detection, name='statistiques_detection'),
] 