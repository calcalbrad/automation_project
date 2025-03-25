import re
import xlwings as xw

from macros import *

categories = ["RR", "Repair", "Paint", "Part", "Sublet"]
current_category = None
category_pattern = re.compile(r"^(RR|Repair|Paint|Part|Sublet):")

current_total_padding = 0

total_labour_items = 0
total_paint_items = 0
total_part_items = 0
total_sublet_items = 0


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
    
    
def handle_labour_category(row):
    global current_total_padding
    global total_labour_items
    
    if total_labour_items > 5:
        # call_add_labour_detail_line_macro()
        current_total_padding += 1
        
    row_to_append = 17 + total_labour_items
    print("Row to Append: "+str(row_to_append)+"; Total Labour Items: "+str(total_labour_items)+"; Total Padding: "+str(current_total_padding)+";")
        
    print_row(row, "B"+str(row_to_append), "C"+str(row_to_append))
    total_labour_items += 1
    

def handle_paint_category(row):
    global current_total_padding
    global total_paint_items
    
    if total_paint_items > 5:
        # call_add_paint_detail_line_macro()
        current_total_padding += 1
        
    row_to_append = 26 + total_paint_items + current_total_padding
    print("Row to Append: "+str(row_to_append)+"; Total Paint Items: "+str(total_paint_items)+"; Total Padding: "+str(current_total_padding)+";")
        
    print_row(row, "B"+str(row_to_append), "C"+str(row_to_append))
    total_paint_items += 1
    
    
def handle_parts_category(row):
    global current_total_padding
    global total_part_items
    
    if total_part_items > 5:
        # call_add_paint_detail_line_macro()
        current_total_padding += 1
        
    row_to_append = 36 + total_part_items + current_total_padding
    print("Row to Append: "+str(row_to_append)+"; Total Part Items: "+str(total_part_items)+"; Total Padding: "+str(current_total_padding)+";")
        
    print_row(row, "B"+str(row_to_append), "D"+str(row_to_append))
    total_part_items += 1
    
    
def handle_sublet_category(row):
    global current_total_padding
    global total_sublet_items
    
    if total_sublet_items > 5:
        # call_add_paint_detail_line_macro()
        current_total_padding += 1
        
    row_to_append = 45 + total_sublet_items + current_total_padding
    print("Row to Append: "+str(row_to_append)+"; Total Sublet Items: "+str(total_sublet_items)+"; Total Padding: "+str(current_total_padding)+";")
        
    print_row(row, "B"+str(row_to_append), "D"+str(row_to_append)) # Need to remove category from this one
    total_sublet_items += 1
    
        
def sort_row_into_categories(row):
    global current_category
    
    description = extract_description(row)
    
    if description:
        match current_category:
            case "RR" | "Repair":
                handle_labour_category(row)
            case "Paint":
                handle_paint_category(row)
            case "Part":
                handle_parts_category(row)
            case "Sublet":
                handle_sublet_category(row)
            
            
def extract_description(row):
    pattern = r"^\d+\.\s([^\d@$]*(?:\d*\w+[^@$]*))(?:\s[@$]|\s\d|\s\(|$)"
    
    # need to add handler for "(not required)"
    
    # need to add handler for "SubTotal $"
    
    match = re.search(pattern, row)
    if match:
        return match.group(1).strip()
    else:
        return None
    
    
def print_row(row, cell1, cell2):    
    global current_row
    description = extract_description(row)
    
    if description:
        append_to_cell(cell1, current_category)
        append_to_cell(cell2, description)
    
    
def text_to_excel(workbook: xw.Book, text): 
    global global_sheet
    global global_text
    
    global_sheet = workbook.sheets[0]
    global_text = text
    
    for line in text.splitlines():
        if line:
            # based on error case I found with OCR
            line = re.sub(r"^(\d+),", r"\1.", line)
            line.strip()
              
            match = category_pattern.match(line)
            if match:
                global current_category
                current_category = match.group(1)
            elif current_category:
                sort_row_into_categories(line)