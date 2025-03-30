import xlwings as xw


def call_add_labour_detail_line_macro(workbook: xw.Book):
    try:
        macro = workbook.macro('AddLabourDetailLine')
        macro()
    except AttributeError:
        print("Error: Macro not found in the workbook. Check the macro name.")
    except xw.XlwingsError as e:
        print(f"Xlwings error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")
  
    
def call_add_paint_detail_line_macro(workbook: xw.Book):
    try:
        macro = workbook.macro('AddPaintDetailLine')
        macro()
    except AttributeError:
        print("Error: Macro not found in the workbook. Check the macro name.")
    except xw.XlwingsError as e:
        print(f"Xlwings error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")
    
    
def call_add_parts_detail_line_macro(workbook: xw.Book):
    try:
        macro = workbook.macro('AddPartsDetailLine')
        macro()
    except AttributeError:
        print("Error: Macro not found in the workbook. Check the macro name.")
    except xw.XlwingsError as e:
        print(f"Xlwings error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


def call_add_sublet_detail_line_macro(workbook: xw.Book):
    try:
        macro = workbook.macro('AddSubletDetailLine')
        macro()
    except AttributeError:
        print("Error: Macro not found in the workbook. Check the macro name.")
    except xw.XlwingsError as e:
        print(f"Xlwings error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")