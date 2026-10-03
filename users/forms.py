from django import forms
from django.contrib.auth.models import User

class UserUpdateForm(forms.ModelForm):
    class Meta:
        model = User
        fields = ['first_name', 'last_name', 'email']

class PasswordChangeForm(forms.Form):
    old_password = forms.CharField(widget=forms.PasswordInput, label='Ancien mot de passe')
    new_password1 = forms.CharField(widget=forms.PasswordInput, label='Nouveau mot de passe')
    new_password2 = forms.CharField(widget=forms.PasswordInput, label='Confirmer le nouveau mot de passe') 