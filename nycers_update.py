#import document template from docx
from docx import Document
#import highlight colors
from docx.enum.text import WD_COLOR_INDEX
#import os for splitting file types
import os

#init original document
file = input("Enter file here: ")
document = docx.Document(file)

#remove .docx file extension and save as UPDATED.docx
file = os.path.splitext(file)[0] + '-UPDATE.docx'
document.save(file)

"""
    parses the document to update highlights

    red => purple
    green => blue
    purple => delete
    blue => unhighlight

"""

def textUpdate(filename):
    doc = docx.Document(filename)
    for chars in doc.paragraphs:
        highlight = chars.text.highlight_color
        if highlight == WD_COLOR_INDEX.RED:
            highlight = WD_COLOR_INDEX.VIOLET
        elif highlight == WD_COLOR_INDEX.GREEN:
            highlight = WD_COLOR_INDEX.BLUE
        elif highlight == WD_COLOR_INDEX.VIOLET:
            highlight = delete #delete
        else:
            highlight = WD_COLOR_INDEX.WHITE
