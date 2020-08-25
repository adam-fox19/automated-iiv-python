# -----------------------------------------------------------
# ms_excel 
#
# Defines utility functions to use when updating excel spreadsheets reg orgs & sales tracker.
# Extends functionality of external openpyxl module.
# 
# -----------------------------------------------------------


# External modules

import openpyxl


# -----------------------------------------------------------
# Utility functions 
# -----------------------------------------------------------

# Returns openpyxl workbook and sheet objects
      
def set_up_sheet_object(filepath, sheet_name):
        sandbox_workbook = openpyxl.load_workbook(filepath)
        return [sandbox_workbook, sandbox_workbook.get_sheet_by_name(sheet_name)]


# Wipes given range of output row of sheet_name of any existing data 

def clear_sheet_row(sheet_name, range_start, range_finish):
    for cell in range(range_start, range_finish):
        sheet_name.cell(row = 4, column = cell).value = None


# Updates specified cell in sheet with new value 

def update_cell(sheet_name, cell_address, new_value):
    try:
        sheet_name[cell_address].value = new_value
    except Exception as e:
        print('Unable to update cell ' + cell_address + ' with ' + new_value)
        print(str(e))
        

# Saves entire workbook (not just individual sheets) to given filepath 

def save_workbook(workbook, filepath):
    workbook.save(filepath)


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


