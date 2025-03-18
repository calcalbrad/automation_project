import re
import xlwings as xw

categories = ["RR", "Repair", "Paint", "Part", "Sublet"]

def append_to_cell(sheet: xw.Sheet, cell, text):
    sheet[cell] = text
    
    
def text_to_excel(sheet: xw.Sheet, text):
    headers = ["Category", "Description", "Quoted Amount", "Assessed Amount"]
    sheet.range("A1").value = headers

    row_num = 2  # Starting row for the data

    # We will keep track of whether we've added the quoted/assessed amounts for the first line in each category
    category_quoted_assessed_added = {category: False for category in categories}

    # Use regular expressions to match the data for each category
    for category in categories:
        pattern = re.compile(rf"{category}:([\s\S]*?)(?=(\n[A-Z][a-z]+:|\n$))", re.MULTILINE)
        matches = re.findall(pattern, text)

        # First, check for the SubTotal $ lines for quoted and assessed amounts
        subtotals_pattern = re.compile(rf"SubTotal \$ ([\d\.,]+) > ([\d\.,]+)", re.MULTILINE)
        subtotal_matches = re.findall(subtotals_pattern, text)
        quoted_amount = assessed_amount = None
        
        # Find the first subtotal for quoted and assessed amounts (skip Sublet category)
        if category != "Sublet" and subtotal_matches:
            quoted_amount, assessed_amount = subtotal_matches[0]

        # Now parse each match (data for this category)
        for match in matches:
            lines = match[0].strip().split('\n')
            
            # Loop through each line in the category
            for line in lines:
                print(line)
                # Example pattern to extract description and amount from each line
                match_data = re.match(r"([\d\.\w\s\(\)\-]+)\s+([\d\.\,]+)\s*>\s*([\d\.\,]+)", line)
                
                print(match_data)
                # ERROR HERE - match_data constantly returning none
                
                if match_data:
                    description = match_data.group(1).strip()
                    quoted = match_data.group(3).strip()
                    assessed = match_data.group(4).strip()

                    # If we're adding quoted/assessed amounts for this category, put them in the first line
                    if not category_quoted_assessed_added[category]:
                        # Add the category, description, and the quoted/assessed amounts to the first line
                        sheet.range(f"A{row_num}").value = category
                        sheet.range(f"B{row_num}").value = description
                        sheet.range(f"C{row_num}").value = quoted_amount
                        sheet.range(f"D{row_num}").value = assessed_amount
                        category_quoted_assessed_added[category] = True
                    else:
                        # Otherwise, just add the category and description without quoted/assessed
                        sheet.range(f"A{row_num}").value = category
                        sheet.range(f"B{row_num}").value = description
                        sheet.range(f"C{row_num}").value = ""  # No quoted amount
                        sheet.range(f"D{row_num}").value = ""  # No assessed amount
                    
                    # Move to the next row
                    row_num += 1