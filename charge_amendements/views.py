from django.shortcuts import render

# Vue d'accueil accessible à tous

def home(request):
    return render(request, 'charge_amendements/home.html') 