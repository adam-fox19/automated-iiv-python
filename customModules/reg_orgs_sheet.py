# -----------------------------------------------------------
# reg_orgs_sheet
#
# Draws on ms_excel functions to update reg orgs sheet in sandbox workbook.
# -----------------------------------------------------------


# Custom modules

from customModules.MSOffice.ms_excel import set_up_sheet_object, clear_sheet_row, update_cell, save_workbook


# -----------------------------------------------------------
# Updates reg_orgs sheet
# -----------------------------------------------------------

def update_reg_orgs(filepath, sheet_name, org_variables, other_variables, assessors):
    
    [workbook, sandbox_reg_orgs] = set_up_sheet_object(filepath, sheet_name)
    
    clear_sheet_row(sandbox_reg_orgs, 1, 22)
    
    update_cell(sandbox_reg_orgs, 'A4', org_variables['org_name'])

    if other_variables['sale_type'] == 'Renewal':
        update_cell(sandbox_reg_orgs, 'C4', other_variables['sale_type'] + ' (' + other_variables['renewal_number'] + ')')
    else:
        update_cell(sandbox_reg_orgs, 'C4', other_variables['sale_type'])

    update_cell(sandbox_reg_orgs, 'D4', org_variables['org_members'])   
    update_cell(sandbox_reg_orgs, 'E4', other_variables['number_of_days'])
    update_cell(sandbox_reg_orgs, 'F4', org_variables['org_area_of_work'])   
    update_cell(sandbox_reg_orgs, 'H4', org_variables['org_twitter'])
    update_cell(sandbox_reg_orgs, 'I4', org_variables['org_head_office_region'])
    update_cell(sandbox_reg_orgs, 'J4', other_variables['uk_wide'])
    update_cell(sandbox_reg_orgs, 'K4', org_variables['person_name'])
    update_cell(sandbox_reg_orgs, 'L4', org_variables['person_telephone'])
    update_cell(sandbox_reg_orgs, 'M4', org_variables['person_email'])
    update_cell(sandbox_reg_orgs, 'O4', other_variables['assessor_name'])
    
    # returns allocated lead assessor's name (each assessor is a key with their lead assessor as the corresponding value)
    
    update_cell(sandbox_reg_orgs, 'P4', assessors[other_variables['assessor_name']])
    
    update_cell(sandbox_reg_orgs, 'Q4', other_variables['date_today'])
    update_cell(sandbox_reg_orgs, 'U4', other_variables['sale_value'])
    update_cell(sandbox_reg_orgs, 'V4', other_variables['discount_value'])
    
    save_workbook(workbook, filepath)

    print('reg orgs sheet updated')


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


# -----------------------------------------------------------
# For testing
# -----------------------------------------------------------


#import os, json
#from dotenv import load_dotenv
#load_dotenv()


#org_variables = {
#                'org_name':'Test org',
#                'org_postcode': 'N0 123',
#                'org_area_of_work': 'Children and young people',
#                'org_head_office_region': 'London',
#                'org_turnover_band': 'Â£500,000 to Â£1 million',
#                'org_no_of_volunteers': '158',
#                'org_no_of_vol_roles': '5',
#                'org_no_of_sites': '1',
#                'person_telephone': '0123 456789',
#                'person_name': 'Adam Fox',
#                'person_email': 'test@testemail.org',
#                'org_twitter': '@test',
#                'org_members': 'Yes',
#                'org_sector': 'Public'
#                 }

#other_variables = {
#                'date_today': '08/01/2020',
#                'current_month': '01/2020',
#                'sale_type': 'New',
#                'number_of_days': '4',
#                'assessor_name': 'Adam Fox',
#                'sale_value': '999',
#                'discount_value': 999,
#                'uk_wide': ''
#                 }


#sandbox_workbook_path = os.environ.get('SANDBOX_PATH')

#assessors = {'Adam Fox' : 'Adam Badger' }

#update_reg_orgs(sandbox_workbook_path, 'reg_orgs', org_variables, other_variables, assessors)
