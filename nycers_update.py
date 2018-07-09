#import docx
from docx import *
#import os for splitting file types
import os


def main():
    #init original document
    file = input("Enter file here: ")
    document = Document(file)

    open(document)
    line = document.readlines
    print(line)
    
    #add every paragraph to a map, 
    #key = paragraph number 
    #value = every run in paragraph as nested key, style for run as nested value
    #maybe have a nested map? Map<para, map<runs,style>>
   
    para = {} # paragraph key
    def runLoader(document):
        for p in document.paragraphs:
            #add paragraph no. to key
            vals = {} #nested key
            para[p] = vals
            for r in p.runs:
                #add runs & styles to vals mapper
                vals[r] = r.style
                if next(r) == None:
                    break
        return

    """
        parses the document to update highlights
        red => purple
        green => blue
        purple => delete
        blue => unhighlight
    """
    #edit so it can access key-value pairs in para map rather than just the doc
    def textUpdate(document):
        doc = document
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
    
    #assemble the updated docx back together
    def assemble(para):
        #take in para and assemble it back together on new document
        return

    #driver
    def driver(document):
        runLoader(document)
        textUpdate(document)
        assemble(document)
        #remove .docx file extension and save as UPDATED.docx
        file = os.path.splitext(file)[0] + '-UPDATE.docx'
        document.save(file)
    
    driver(document)    

main()
