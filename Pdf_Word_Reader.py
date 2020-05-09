'''
Fully-baked code snippets/functions for the Medium blog 
'''
####========================================================
#### Section 1: Python-docx → work with MS Word .docx files
####========================================================
'''
### read in the .docx file extention
Install the docx module in anaconda:
    conda install -c conda-forge python-docx
'''
import docx
doc = docx.Document('..\\Sample_File_DOCX.docx')
paras = [p.text for p in doc.paragraphs if p.text]    
print(f'=== Output type is a {type(paras)} of {type(paras[1])} \ntotal length is {len(paras)} ===') 

   
####========================================================
#### Section 2: Win32com → work with MS Word .doc files
####========================================================   
'''
### read in the .doc file extention
''' 
import os
import win32com
from win32com import client
def docReader(doc_file_name):    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False 
    
    _ = word.Documents.Open(doc_file_name)
    
    doc = word.ActiveDocument
    paras = doc.Range().text    
    doc.Close()
    word.Quit()
    
    return paras    

cur_file_full_name = os.getcwd() + '\\' + 'Sample_File_DOC.doc'
doc_out = docReader(cur_file_full_name)
print(f'=== Output type is a {type(doc_out)} ===') 
### Processing the text paragraph into a list of substrings 
[i for i in doc_out.replace('\x07', '\r').split('\r') if i]


####==============================================================
#### Section 3: Pdfminer (in lieu of PyPDF2) → work with PDF text
####==============================================================
'''
PyPDF2 to extract PDF text
'''
from PyPDF2 import PdfFileReader
def text_extractor(file_name):
    with open(file_name, 'rb') as f:
        pdf = PdfFileReader(f)
        # Get the first page
        page = pdf.getPage(0)
        paras = page.extractText()        
    f.close()
        
    return paras
            
if __name__ == '__main__':
    cur_file = 'Sample_File_PDF_Text.pdf'  
    out1 = text_extractor(cur_file)    
    print(out1.split('\n'))


'''
### Read in PDF complicated text using Pdfminer
'''      
from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

def pdf_text_reader(pdf_file_name, pages=None):
    if pages:
        pagenums = set(pages)
    else:
        pagenums = set()

    ## 1) Initiate the Pdf text converter and interpreter
    textOutput = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, textOutput, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    ## 2) Extract text from file using the interpreter
    infile = open(pdf_file_name, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)        
    infile.close()
    
    ## 3) Extract the paragraphs and close the connections
    paras = textOutput.getvalue()   
    converter.close()
    textOutput.close
    
    return paras

if __name__ == '__main__':    
    out = pdf_text_reader('Sample_File_PDF_Text.pdf', pages=[0])   
    print([i for i in out.split('\n') if i])

    
    
####=====================================================================
#### Section 4: Pdf2image + Pytesseract → work with PDF scanned-in images
####=====================================================================
import pandas as pd
import pytesseract as pt
import pdf2image
import os

'''
install pytesseract and pdf2image.
When running the func pdf_image_reader, if you got the error(s) below,
    1) error msg: PDFInfoNotInstalledError: Unable to get page count. Is poppler installed and in PATH?
        solution: 
            install the poppler first, in conda: conda install -c conda-forge poppler
    
    2) error msg: TesseractNotFoundError: tesseract is not installed or it's not in your PATH
        solution by following the steps below:
            Step 1: Download and install the Tesseract OCR here.
            Step 2: After installing, find the folder 'Tesseract-OCR', and then find the tesseract.exe in this folder.
            Step 3: Copy the file location to the tesseract.exe.
            Step 4: Set your tesseract_cmd to this location, as shown below, 
                        pt.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
                        or 
                        pt.pytesseract.tesseract_cmd = "PATH TO YOUR tesseract.exe"
'''
pt.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

def pdf_image_reader(pdf_file_name, image_folder_name='images'):
    name = pdf_file_name.split('.pdf')[0]
    ### 1) Initiate to store the converted images
    pages = pdf2image.convert_from_path(pdf_path=pdf_file_name, dpi=200, size=(1654,2340))

    ### 2) Make the dir to store the images (if not already exists)
    try:
        os.makedirs(image_folder_name)
    except:
        pass
    
    ### 3) Save each page as one image        
    saved_image_name = f'{image_folder_name}\\{name}'
    for i in range(len(pages)):
        pages[i].save(saved_image_name + str(i) + '.jpg')
    
    ### 4) Extract the content by converting images to a list of strings 
    content = ''    
    for i in range(len(pages)):
        content += pt.image_to_string(pages[i]) 
    
    return content 

if __name__ == '__main__':     
    out = pdf_image_reader('Sample_File_PDF_Image.pdf', image_folder_name='images')
    print(' '.join([i for i in out.split('\n') if i]))
    




    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

