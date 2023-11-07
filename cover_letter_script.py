import re
from docx import Document
from docx.shared import Pt
import PySimpleGUI as gui

FILE_SRC = r"C:\Users\kaush\Desktop\Jobs\cover_letter_editor\template_cover_letter\0_Cover letter_Kaushik Moudgalya_v2.docx"
#get company name from the user:
COMPANY = gui.PopupGetText("Enter the company name: ")
POSITION = gui.PopupGetText("Enter the position name: ")
# TO DO: we need a dialog box that will ask me to enter the company name
FILE_DEST = r"C:\Users\kaush\Desktop\Jobs\cover_letter_editor\generated_cover_letters\\" + str(COMPANY) + r"_Cover letter_Kaushik Moudgalya_v2.docx"

# replace the "Position" and the "Company" strings.
document = Document(FILE_SRC)
# some stuff for the font
style = document.styles['Normal']
font = style.font
font.name, font.size = 'Arial', Pt(10.5)
for index, para in enumerate(document.paragraphs):
    if "Position" in para.text and "Company" in para.text:
        runs = para.runs
        for run in runs:
            if "Position" in run.text:
                # Find and replace "Position" with the user input
                new_text = run.text.replace("Position", POSITION)
                run.text = new_text
                run.bold = True  # Set the bold formatting for "POSITION"
            if "Company" in run.text:
                # Find and replace "Company" with the user input
                new_text = run.text.replace("Company", COMPANY)
                run.text = new_text
                run.bold = True  # Set the bold formatting for "COMPANY"
        para.style = style
        break
document.save(FILE_DEST)
