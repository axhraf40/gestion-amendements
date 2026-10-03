from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from amendements.models import AmendementFichier

@login_required
def dashboard(request):
    total_fichiers = AmendementFichier.objects.filter(utilisateur=request.user).count()
    return render(request, 'stats/dashboard.html', {
        'total_fichiers': total_fichiers,
    })
