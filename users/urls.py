from django.urls import path
from . import views
 
urlpatterns = [
    path('profile/', views.profile, name='profile'),
    path('change-password/', views.change_password, name='change_password'),
] 