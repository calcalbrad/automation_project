import re
import xlwings as xw

categories = ["RR", "Repair", "Paint", "Part", "Sublet"]
current_category = None
category_pattern = re.compile(r"^(RR|Repair|Paint|Part|Sublet):")

current_row = 1


def append_to_cell(cell, value):
    global_sheet.range(cell).value = value
        
    
def parse_repairer():
    append_to_cell("H2", "")
    
    
def parse_owner():
    append_to_cell("H8", "")
    
    
def parse_vehicle_info():
    append_to_cell("C3", "")
    append_to_cell("C4", "")
    append_to_cell("C5", "")
    
    
def parse_registation_info():
    append_to_cell("C6", "")
    

def parse_claim_number():
    append_to_cell("C2", "")
    
    
def extract_description(row):
    pattern = r"^\d+\.\s(.*?)(?:\s[@$]|\s\d|\s\(|$)"
    
    # need to add handler for "SubTotal $"
    
    match = re.search(pattern, row)
    if match:
        return match.group(1).strip()
    else:
        print("Returning None")
        return None
    
    
def print_row(row):    
    global current_row
    description = extract_description(row)
    
    if description:
        append_to_cell("A"+str(current_row), current_category)
        append_to_cell("B"+str(current_row), description)
        print(description)
        
        current_row += 1
    
    
def text_to_excel(sheet: xw.Sheet, text): 
    global global_sheet
    global global_text
    
    global_sheet = sheet
    global_text = text
    
    for line in text.splitlines():
        if line:
            line.strip()
            
            print(line)
              
            match = category_pattern.match(line)
            if match:
                global current_category
                current_category = match.group(1)
            elif current_category:
                print_row(line)
    