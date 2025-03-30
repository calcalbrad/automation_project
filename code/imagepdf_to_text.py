import pytesseract
import xlwings as xw

from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    
    custom_config = r'--psm 6 --oem 1'
    
    text = ''
    for image in images:
        text += pytesseract.image_to_string(image, config=custom_config)
        
    return text