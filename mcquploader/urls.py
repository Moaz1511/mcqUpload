from django.urls import path
from . import views

urlpatterns = [
    # path('', views.upload_file, name='home'),  # Map the root URL to the upload_file view
    # Home URL
    path('', views.home, name='home'),

    # Upload MCQ URL for mcquploader
    path('upload/', views.file_upload, name='file_upload'),

    # Success URL after upload
    path('upload/success/', views.upload_success, name='success_url'),

    # Export worksheet URL
    path('upload/export-worksheet/', views.export_worksheet, name='export_worksheet'),

    # Download lecture slide for mcqdownloader
    path('download-lecture-slide/', views.download_lecture_slide, name='download_lecture_slide'),
]
