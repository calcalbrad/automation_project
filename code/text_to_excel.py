import re
import xlwings as xw

from macros_handler import *

current_category = None

start_rows = {
    "Labour": 17,
    "Repair": 17,
    "Paint": 26,
    "Part": 36,
    "Sublet": 45
}

total_labour_items = 0
total_repair_items = 0
total_paint_items = 0
total_part_items = 0
total_sublet_items = 0


def append_to_cell(cell, value):
    global_sheet.range(cell).value = value
        
    
def append_owner(owner):
    append_to_cell("H8", owner)
    
    
def append_vehicle_info(vehicle):
    vehicle_split = re.split(r"[- ]", vehicle)
    
    append_to_cell("C3", vehicle_split[0])
    append_to_cell("C4", vehicle_split[1])
    append_to_cell("C5", vehicle_split[2])
    
    
def append_registation_info(registration):
    append_to_cell("C6", registration)
    

def append_claim_number(claim_number):
    append_to_cell("C2", claim_number)
    
    
def handle_labour_category(row):
    global global_workbook
    global start_rows
    global total_labour_items
    
    if total_labour_items >= 5:
        call_add_labour_detail_line_macro(global_workbook)
        start_rows["Paint"] += 1
        start_rows["Part"] += 1
        start_rows["Sublet"] += 1
        
    row_to_append = start_rows["Labour"] + total_labour_items        
    print_row(row, "C", row_to_append)
    total_labour_items += 1
    
    start_rows["Repair"] += 1
    
    
def handle_repair_category(row):
    global global_workbook
    global start_rows
    global total_labour_items
    global total_repair_items
    
    if (total_labour_items + total_repair_items) >= 5:
        call_add_labour_detail_line_macro(global_workbook)
        start_rows["Paint"] += 1
        start_rows["Part"] += 1
        start_rows["Sublet"] += 1
        
    row_to_append = start_rows["Repair"] + total_repair_items        
    print_row(row, "C", row_to_append)
    total_repair_items += 1


def handle_paint_category(row):
    global global_workbook
    global start_rows
    global total_paint_items
    
    if total_paint_items >= 5:
        call_add_paint_detail_line_macro(global_workbook)
        start_rows["Part"] += 1
        start_rows["Sublet"] += 1
        
    row_to_append = start_rows["Paint"] + total_paint_items        
    print_row(row, "C", row_to_append)
    total_paint_items += 1
    
    
def handle_parts_category(row):
    global global_workbook
    global start_rows
    global total_part_items
    
    if total_part_items >= 5:
        call_add_parts_detail_line_macro(global_workbook)
        start_rows["Sublet"] += 1
        
    row_to_append = start_rows["Part"] + total_part_items        
    print_row(row, "D", row_to_append)
    total_part_items += 1
    
    
def handle_sublet_category(row):
    global global_workbook
    global start_rows
    global total_sublet_items
        
    if total_sublet_items >= 5:
        call_add_sublet_detail_line_macro(global_workbook)
        
    row_to_append = start_rows["Sublet"] + total_sublet_items        
    print_row(row, "E", row_to_append)
    total_sublet_items += 1
        

def handle_claim_information(line):
    match line: 
        case line if line.startswith("Owner"):
            owner = extract_information("Owner", line)
            if owner:
                append_owner(owner)
        case line if line.startswith("Vehicle"):
            vehicle = extract_information("Vehicle", line)
            if vehicle:
                append_vehicle_info(vehicle)
        case line if line.startswith("Reg No"):
            reg_no = extract_information("Reg No", line)
            if reg_no:
                append_registation_info(reg_no)
        case line if line.startswith("Claim #"):
            claim_number = extract_information("Claim #", line)
            if claim_number:
                append_claim_number(claim_number)
            

# Captures the value after a specific keyword
def extract_information(label, text):
    match = re.search(f"{label}\s+(.+)", text)
    if match:
        return match.group(1).strip()
    return None
        
        
def extract_description(row): 
    if "not required" not in row:
        pattern = r"^\d+\.\s([^\d@$]*(?:\d*\w+[^@$]*))(?:\s[@$]|\s\d|\s\(|$)"
        
        match = re.search(pattern, row)
        if match:
            return match.group(1).strip()
        else:
            return None


def check_for_subtotal(row):
    return row.startswith("SubTotal $")


def extract_subtotal(subtotal):
    match = re.search(r"\$([\d,.]+) > \$([\d,.]+)", subtotal)

    if match:
        first_amount = float(match.group(1).replace(",", "")) 
        second_amount = float(match.group(2).replace(",", ""))
        return(first_amount, second_amount) 
    
    
def print_row(row, description_column_letter, row_number):    
    global current_category
    description = extract_description(row)
    
    if description:
        if current_category in ["RR", "Repair", "Paint"]:
            append_to_cell("B"+str(row_number), current_category)
        append_to_cell(description_column_letter+str(row_number), description)
  
  
def get_subtotal_row():
    global current_category
    global start_rows
        
    match current_category:
        case "RR":
            return start_rows["Labour"]
        case "Repair":
            return start_rows["Repair"]
        case "Paint":
            return start_rows["Paint"]
        case "Part":
            return start_rows["Part"]
        case "Sublet":
            return start_rows["Sublet"]
            
            
def print_subtotal(subtotal):
    global current_category
    
    if current_category != "Sublet":
        subtotal_float = extract_subtotal(subtotal)
        row_number = get_subtotal_row()
        
        append_to_cell("J"+str(row_number), subtotal_float[0])
        append_to_cell("K"+str(row_number), subtotal_float[1])
        
        
def sort_row_into_category(row):
    global current_category
    
    subtotal = check_for_subtotal(row)
    description = extract_description(row)
    
    if subtotal:
        print_subtotal(row)
    elif description:
        match current_category:
            case "RR":
                handle_labour_category(row)
            case "Repair":
                handle_repair_category(row)
            case "Paint":
                handle_paint_category(row)
            case "Part":
                handle_parts_category(row)
            case "Sublet":
                handle_sublet_category(row)
            
          
def read_through_text(text):
    category_pattern = re.compile(r"^(RR|Repair|Paint|Part|Sublet):")
    
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
                sort_row_into_category(line)
            else:
                handle_claim_information(line)
    
    
def text_to_excel(workbook: xw.Book, text): 
    global global_workbook
    global global_sheet
    global global_text
    
    global_workbook = workbook
    global_sheet = workbook.sheets[0]
    global_text = text
    
    read_through_text(text)