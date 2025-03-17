import pytesseract
import pandas as pd
import openpyxl

from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    
    custom_config = r'--psm 6 --oem 1'
    
    text = ''
    for image in images:
        text += pytesseract.image_to_string(image, config=custom_config)
        
    return text

        
def append_text_to_workbook(text, excel_path):    
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    
    # insert text_to_excel reference here
    
    
def main(pdf_path, excel_path):
    text = extract_text_from_pdf(pdf_path)    
    workbook = append_text_to_workbook(text, excel_path)
    workbook.save(excel_path)

if __name__ == "__main__":
    main("data/KNZ151 Quote.pdf", "test.xlsm")