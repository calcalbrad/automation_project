import xlwings as xw

# TODO - add error handling

def call_add_labour_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddLabourDetailLine')
  
    
def call_add_paint_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddPaintDetailLine')
    
    
def call_add_parts_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddPartsDetailLine')


def call_add_sublet_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddSubletDetailLine')