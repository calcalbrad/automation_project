import re
import xlwings as xw

categories = ["RR", "Repair", "Paint", "Part", "Sublet"]

def __init__(self, sheet: xw.sheet, text):
    self.sheet = sheet
    self.text = text


def append_to_cell(self, cell, value):
    self.sheet[cell] = value
        
    
def parse_repairer(self):
    append_to_cell(self, "H2", "")
    
    
def parse_owner(self):
    append_to_cell(self, "H8", "")
    
    
def parse_vehicle_info(self):
    append_to_cell(self, "C3", "")
    append_to_cell(self, "C4", "")
    append_to_cell(self, "C5", "")
    
    
def parse_registation_info(self):
    append_to_cell(self, "C6", "")
    

def parse_claim_number(self):
    append_to_cell(self, "C2", "")
    
    
def text_to_excel(self): 
    print()
    