import xlwings as xw

def append_to_cell(sheet: xw.Sheet, cell, text):
    sheet[cell] = text