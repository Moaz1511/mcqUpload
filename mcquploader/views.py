# mcquploader/views.py
import os
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.conf import settings
from .forms import UploadFileForm
from .utils import process_pptx, export_worksheet_as_excel  # Assume you've adapted your script and placed it in utils.py

def home(request):
    return HttpResponse('Welcome to the homepage!')
def upload_success(request):
    # return HttpResponse('File successfully uploaded!')
    return render(request, 'mcquploader/upload_success.html')

def file_upload(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            # Process the uploaded file
            file = request.FILES['file']
            file_path = os.path.join(settings.MEDIA_ROOT, file.name)
            with open(file_path, 'wb+') as destination:
                for chunk in file.chunks():
                    destination.write(chunk)
            process_pptx(file_path)  # Call your adapted processing function
            return redirect('success_url')  # Redirect or indicate success
    else:
        form = UploadFileForm()
    return render(request, 'mcquploader/upload.html', {'form': form})

'''def export_worksheet(request):
    # You might want to get the spreadsheet ID and worksheet title dynamically
    spreadsheet_id = '1G4WPQi9tlRjvS1SRGM721QjoYDCpJ1MCLzlPgx9L250'
    worksheet_title = 'Sheet1'
    
    # Ensure you handle GET or POST request appropriately and check user permissions if necessary
    if request.method == 'POST':
        return export_worksheet_as_excel(spreadsheet_id, worksheet_title)
    
    # If it's not a POST request, or if you need to display a form, render a template
    return render(request, 'mcquploader/upload.html')'''

def export_worksheet(request):
    # Example spreadsheet ID and worksheet title
    spreadsheet_id = '1G4WPQi9tlRjvS1SRGM721QjoYDCpJ1MCLzlPgx9L250'
    worksheet_title = 'Sheet1'
    
    # Call function to create Excel file
    excel_file_path = export_worksheet_as_excel(spreadsheet_id, worksheet_title)

    # Create HTTP response with Excel file
    with open(excel_file_path, 'rb') as excel_file:
        response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="{worksheet_title}.xlsx"'
    
    # Clean up: remove the temporary Excel file
    os.remove(excel_file_path)

    return response


