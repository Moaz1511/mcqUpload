import re
import pandas as pd
import gspread
from pptx import Presentation
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2.service_account import Credentials
import tempfile
import os
import shutil
from django.http import HttpResponse
from copy import deepcopy

def extract_mcq_info(text):
    # Define the regular expression pattern to match questions, references, options, answers, and explanations
    question_regex = r"(\d+।)\s*(.*?)\s*(?:\[(.*?)\])?\s+\(ক\)\s*(.*?)\s+\(খ\)\s*(.*?)\s+\(গ\)\s*(.*?)\s+\(ঘ\)\s*(.*?)\s*(?:\s+উত্তর:\s+(.*?))?(?:\s+ব্যাখ্যা:\s+(.*?))?(?=\d+।|$)"

    match = re.finditer(question_regex, text, re.DOTALL)
    mcq_list = []
    for m in match:
        mcq_list.append(m.groups())
    return mcq_list

def process_pptx(file_path):
    # Load PowerPoint presentation
    presentation = Presentation(file_path)

    # Authenticate and open Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    credentials_path = "E:/Developer/Website/Django/TwigTech/mcquploader/credentials.json"
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)

    client = gspread.authorize(credentials)
    sheet = client.open("Automation | PowerPoint to Google Sheet to PowerPoint").worksheet('UploadToGoogleSheet')  # Access the specified worksheet

    row = 2  # Start from row 2 to avoid header

    batch_updates = []

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                mcq_info_list = extract_mcq_info(text)
                for mcq_info in mcq_info_list:
                    # Define the mapping of columns in the Google Sheet
                    # Change the column numbers as needed
                    serial_number_col = 1  # Column for serial number
                    question_col = 2  # Column for questions
                    reference_col = 3  # Column for reference
                    option_k_col = 4  # Column for option (ক)
                    option_kh_col = 5  # Column for option (খ)
                    option_g_col = 6  # Column for option (গ)
                    option_gh_col = 7  # Column for option (ঘ)
                    answer_col = 8  # Column for answers
                    explanation_col = 9  # Column for ব্যাখ্যা

                    # Initialize variables for each column value
                    serial_number, question, reference, option_k, option_kh, option_g, option_gh, answer, explanation = "", "", "", "", "", "", "", "", ""

                    # Extract all available data
                    if mcq_info:
                        serial_number, question, reference, option_k, option_kh, option_g, option_gh, answer, explanation = mcq_info

                    # Check if answer and explanation are empty and handle them
                    if not answer:
                        answer = ""
                    if not explanation:
                        explanation = ""

                    # Append the update as a list of values with columns
                    update_values = [
                        serial_number,
                        question,
                        reference,
                        option_k,
                        option_kh,
                        option_g,
                        option_gh,
                        answer,
                        explanation,
                    ]

                    batch_updates.append(update_values)

    # Update the Google Sheet in batches
    if batch_updates:
        sheet.update(f'A{row}:I{row + len(batch_updates) - 1}', batch_updates)


def export_worksheet_as_excel(spreadsheet_id, worksheet_title):
    # Authenticate with Google Sheets
    credentials = Credentials.from_service_account_file('E:/Developer/Website/Django/TwigTech/mcquploader/credentials.json', scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(credentials)

    # Open the spreadsheet and worksheet
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet(worksheet_title)

    # Get all values in the worksheet
    data = worksheet.get_all_values()
    # Convert to a pandas DataFrame
    df = pd.DataFrame(data)
    df.columns = df.iloc[0] # Set first row as column names
    df = df.iloc[1:] # Remove first row

    # Create a Pandas Excel writer using XlsxWriter as the engine
    excel_file = f"{worksheet_title}.xlsx"
    sheet_name = 'Sheet1'

    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    return excel_file

# Set up Google Sheets API
def google_sheets_login():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('E:/Developer/Website/Django/TwigTech/mcquploader/credentials.json', scope)
    client = gspread.authorize(creds)
    return client


# Function to fetch data from Google Sheets
def get_google_sheet_data(sheet_name):
    client = google_sheets_login()
    sheet = client.open(sheet_name).worksheet('DownloadLectureSlide')  # Access the specified worksheet
    data = sheet.get_all_values()  # Fetch all data from the sheet
    return data

# Function to duplicate a slide
def duplicate_slide(prs, index):
    """Duplicate the slide at the given index."""
    slide = prs.slides[index]
    slide_layout = slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            new_shape = deepcopy(shape)
            new_slide.shapes._spTree.insert_element_before(new_shape._element, 'p:extLst')

    return new_slide


# Function to copy data into PowerPoint and download lecture slides
def copy_data_to_ppt(slide_number, sheet_data, template_path, output_path):
    prs = Presentation(template_path)

    for i in range(3, slide_number + 3):
        # Add or duplicate slides
        if i < len(prs.slides):
            slide = prs.slides[i]
        else:
            slide = duplicate_slide(prs, 3)

        # Extract data from the Google Sheet
        n = sheet_data[i - 3][0]  # Number
        q = sheet_data[i - 3][1]  # Question
        r = sheet_data[i - 3][2]  # Reference
        o1 = sheet_data[i - 3][3]  # Option A
        o2 = sheet_data[i - 3][4]  # Option B
        o3 = sheet_data[i - 3][5]  # Option C
        o4 = sheet_data[i - 3][6]  # Option D
        a = sheet_data[i - 3][7]  # Answer

        # Check if the slide has enough shapes to add data
        if len(slide.shapes) >= 8:
            slide.shapes[0].text = n
            slide.shapes[1].text = q
            slide.shapes[2].text = r
            slide.shapes[3].text = o1
            slide.shapes[4].text = o2
            slide.shapes[5].text = o3
            slide.shapes[6].text = o4
            slide.shapes[7].text = a

    # Save the PowerPoint presentation
    prs.save(output_path)