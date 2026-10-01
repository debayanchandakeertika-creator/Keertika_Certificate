from django.urls import path
from django.conf.urls.static import static
from django.conf import settings
from . import views

urlpatterns = [
    path('aadhar/', views.index, name='index'),
    path('aadhar/api.aspx', views.api, name='api'),
    
]+ static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
