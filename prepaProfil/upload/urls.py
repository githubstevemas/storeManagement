from django.urls import path
from .views import process_upload

urlpatterns = [
    path('', process_upload, name='process_upload'),
]
