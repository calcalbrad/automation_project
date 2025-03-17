import openpyxl

def append_to_cell(sheet: openpyxl.worksheet, cell, text):
    sheet[cell] = text