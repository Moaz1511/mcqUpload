from django.urls import path
from . import views

urlpatterns = [
    # path('', views.upload_file, name='home'),  # Map the root URL to the upload_file view
    path('', views.file_upload, name='home'),  # Map the root URL to the upload_file view
    path('upload/', views.file_upload, name='file_upload'),
    # path('upload/success/', views.upload_success, name='success_url'),
    path('upload/success/', views.upload_success, name='success_url'),
    path('upload/export-worksheet/', views.export_worksheet, name='export_worksheet'),
    # path('upload/', views.upload_file, name='file_upload'),
]
