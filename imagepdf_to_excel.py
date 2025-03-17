import pytesseract
import pandas as pd

from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    
    custom_config = r'--psm 6 --oem 1'
    
    text = ''
    for image in images:
        text += pytesseract.image_to_string(image, config=custom_config)
        
    return text

    
def process_text_to_dataframe(text, excel_path):
    df = pd.read_excel()
    

def save_to_excel(df, output_file):
    df.to_excel(output_file, index=False) 
    

def main(pdf_path, excel_path): # add output_path to parameters
    text = extract_text_from_pdf(pdf_path)    
    df = process_text_to_dataframe(text, excel_path)
    # save_to_excel (df, output_path)

if __name__ == "__main__":
    main("data/KNZ151 Quote.pdf", "", "data/2025 versoin 1 -MASTER SPREADSHEET - MOTUKA (003).xlsm") # add output file