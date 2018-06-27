#import docx
from docx import *
#import os for splitting file types
import os

def main():
    #init original document
    file = test.docx #input("Enter file here: ")
    document = docx.Document(file)

    open(document)
    line = document.readlines
    print(line)
    

    """
        parses the document to update highlights

        red => purple
        green => blue
        purple => delete
        blue => unhighlight

    """
'''
    def textUpdate(filename):
        doc = docx.Document(filename)
        for chars in doc.paragraphs:
            for run in chars.runs:
                highlight = chars.text.highlight_color
                if highlight == WD_COLOR_INDEX.RED:
                    highlight = WD_COLOR_INDEX.VIOLET
                elif highlight == WD_COLOR_INDEX.GREEN:
                    highlight = WD_COLOR_INDEX.BLUE
                elif highlight == WD_COLOR_INDEX.VIOLET:
                    highlight = delete #delete
                else:
                    highlight = None

    #driver
    textUpdate(document)
    #remove .docx file extension and save as UPDATED.docx
    file = os.path.splitext(file)[0] + '-UPDATE.docx'
    document.save(file)
'''
main()
