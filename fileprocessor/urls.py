from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('charges/', views.charges, name='charges'),
    path('surgeries/', views.surgeries, name='surgeries'),
]
