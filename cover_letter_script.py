import re
from docx import Document
from docx.shared import Pt
import PySimpleGUI as gui

FILE_SRC = r"C:\Users\kaush\Desktop\Resumes\Cover Letters\Script\0_Cover letter_Kaushik Moudgalya_v2.docx"
#get company name from the user:
COMPANY = gui.PopupGetText("Enter the company name: ")
POSITION = gui.PopupGetText("Enter the position name: ")
# TO DO: we need a dialog box that will ask me to enter the company name
FILE_DEST = r"C:\Users\kaush\Desktop\Resumes\Cover Letters\generated_cover_letters\\" + str(COMPANY) + r"_Cover letter_Kaushik Moudgalya_v2.docx"

# replace the "Position" and the "Company" strings.
document = Document(FILE_SRC)
# some stuff for the font
style = document.styles['Normal']
font = style.font
font.name, font.size = 'Arial', Pt(10.5)
# font.size = Pt(10.5)
# loop through the paragraphs and replace the strings
for index, para in enumerate(document.paragraphs):
    # print(index, para.text)
    if "Position" and "Company" in para.text:
        # print(para.text)
        new_text = re.sub("Position", POSITION, para.text)
        new_text = re.sub("Company", COMPANY, new_text)
        para.text = new_text
        para.style = style
        break
# following code tests if the strings are replaced
# for index, para in enumerate(document.paragraphs):
#     print(index, para.text)
document.save(FILE_DEST)
