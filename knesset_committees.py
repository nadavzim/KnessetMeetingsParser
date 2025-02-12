import re
import win32com.client
import pandas as pd
import textract
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import os


def clean_text(original_text):
    """get rid of unwanted characters from the text."""
    # Keep Hebrew, English, digits, spaces, and common punctuation
    cleaned_text = re.sub(r"[^a-zA-Z0-9\u0590-\u05FF\s.,!?;:()\"'-]", "", original_text)
    return cleaned_text


def extract_text_from_doc(file_path):
    """Extracts text from a .doc file using the Word application. returns the text."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in background
    try:
        doc = word.Documents.Open(file_path)
    except Exception as e:
        print(f"Error: {e}")
        print("Make sure the file exists and is a valid .doc file.")
        exit(1)

    text = doc.Content.Text  # Extract all text
    print(f"\nText extracted from '{file_path}' successfully!\n")
    doc.Close(False)  # Close the document
    word.Quit()  # Close Word application
    return text


def save_to_excel(meetings, excel_filepath):
    """Saves the meeting details to an Excel file."""
    columns = ["יום", "שעה", "וועדה", "נושא", "הערות"]  # Define column headers (matching your Excel file)
    df = pd.DataFrame(meetings, columns=columns)  # Create a DataFrame
    df.to_excel(excel_filepath, index=False)  # Save to Excel file


def format_excel(excel_filepath, col, row):
    wb = load_workbook(excel_filepath)
    ws = wb.active

    ws.sheet_view.rightToLeft = True

    ws.column_dimensions['A'].width = 8.42
    ws.column_dimensions['B'].width = 7
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 85
    ws.column_dimensions['E'].width = 25

    for i in range(1, row + 1):
        for j in range(1, col + 1):
            if j == 1:
                ws.cell(row=i, column=j).fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
            elif j == 2:
                ws.cell(row=i, column=j).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
            ws.cell(row=i, column=j).alignment = Alignment(horizontal="right", vertical="center", readingOrder=2, wrap_text=True)
            ws.cell(row=i, column=j).font = Font(name='Calibri', size=10)  # Set font
            ws.cell(row=i, column=j).border = Border(left=Side(style='thin'),
                                                     right=Side(style='thin'),
                                                     top=Side(style='thin'),
                                                     bottom=Side(style='thin'))
            # make header bold
            for cell in ws[1]:
                cell.font = Font(size=12, bold=True)

    wb.save(excel_filepath)
    print(f"Excel file '{excel_filepath}' formatted successfully!")


def parse_doc_to_excel(txt, output_excel_filepath, delimiter="יום"):
    """
    Parses the text  and saves the data to an Excel file.
       :param delimiter: the character separating data fields.
       :param output_excel_filepath: the path to save the Excel file.
       :param txt: the text extracted from the .doc file.
    """
    cleand_text = clean_text(txt)
    valid_lines = filter_valid_lines(cleand_text)
    meetings = extract_meeting_details(valid_lines, delimiter)
    save_to_excel(meetings, output_excel_filepath)
    format_excel(output_excel_filepath, 5, len(meetings))


def extract_meeting_details(text_lines, delimiter):
    # Split the text into individual meetings
    meetings = []
    i = 0
    while i < len(text_lines):
        line = text_lines[i]

        if line.startswith(delimiter):
            words = line.split(" ")
            date = " ".join(words[:6])[:-2]
            time = words[-1]
            title = text_lines[i + 1]  # the next line is the title
            subject = text_lines[i + 2]
            # while not text_lines[i].startswith(delimiter):
            #     subject += text_lines[i]
            i += 3
            meetings.append((date, time, title, subject, ""))
        else:
            i += 1
    return meetings


def filter_valid_lines(text_):
    lines = text_.split("\r\r")  # Split into lines
    valid_lines = [line.strip() for line in lines if re.match(r'^\s*[a-zA-Zא-ת0-9]', line)]
    return valid_lines  # Join back into a cleaned text


base_dir = os.getcwd()  # Gets the current working directory
doc_file = os.path.join(base_dir, "Committee.doc")  # Full path to the .doc file
excel_file = os.path.join(base_dir, "output.xlsx")  # Full path to the Excel file

text_content = extract_text_from_doc(doc_file)
parse_doc_to_excel(text_content, excel_file)
