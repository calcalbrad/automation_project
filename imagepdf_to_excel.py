import pytesseract
import xlwings as xw

from pdf2image import convert_from_path
from text_to_excel import text_to_excel

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    
    custom_config = r'--psm 6 --oem 1'
    
    text = ''
    for image in images:
        text += pytesseract.image_to_string(image, config=custom_config)
        
    return text

        
def append_text_to_workbook(text, excel_path):    
    with xw.App(visible=True) as app:
        workbook = app.books.open(excel_path)
        text_to_excel(workbook.sheets[0], text)
        workbook.save()
    
    
def main(pdf_path, excel_path):
    text = extract_text_from_pdf(pdf_path)    
    append_text_to_workbook(text, excel_path)

if __name__ == "__main__":
    main("data/KNZ151 Quote.pdf", "test.xlsx")