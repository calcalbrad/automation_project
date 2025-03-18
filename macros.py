import xlwings as xw

def call_macro(marco, workbook: xw.Book):
    #add logic to know which macro to call here
    workbook.save()

def call_add_labour_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddLabourDetailLine')
  
    
def call_add_paint_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddPaintDetailLine')
    
    
def call_add_parts_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddPartsDetailLine')


def call_add_sublet_detail_line_macro(workbook: xw.Book):
    workbook.macro('AddSubletDetailLine')