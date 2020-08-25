# -----------------------------------------------------------
# reg_orgs_sheet
#
# Draws on ms_excel functions to update sales tracker sheet in sandbox workbook.
# -----------------------------------------------------------

# External modules

import time
from datetime import date, timedelta


# Custom modules

from customModules.MSOffice.ms_excel import set_up_sheet_object, clear_sheet_row, update_cell, save_workbook


# Returns financial year quarter based on current month

def fy_quarter(current_month):

    a = 'Q2 2020/21'
    b = 'Q3 2020/21'
    c = 'Q4 2020/21'
    d = 'Q1 2021/22'
    e = 'Q2 2021/22'
    f = 'Q3 2021/22'
    g = 'Q4 2021/22'

    month = {
        
        '07/2020': a,
        '08/2020': a,
        '09/2020': a,
        '10/2020': b,
        '11/2020': b,
        '12/2020': b,
        '01/2021': c,
        '02/2021': c,
        '03/2021': c,
        '04/2021': d,
        '05/2021': d,
        '06/2021': d,
        '07/2021': e,
        '08/2021': e,
        '09/2021': e,
        '10/2021': f,
        '11/2021': f,
        '12/2021': f,
        '01/2022': g,
        '02/2022': g,
        '03/2022': g,
        
        }
    
    try:
        return month[current_month]
    except Exception as e:
        print('fy function hasn\'t worked')
        print(str(e))
        return


# -----------------------------------------------------------
# Updates sales tracker sheet
# -----------------------------------------------------------

def update_sales_tracker(filepath, sheet_name, org_variables, other_variables):
    
    [workbook, sandbox_sales_tracker] = set_up_sheet_object(filepath, sheet_name)

    clear_sheet_row(sandbox_sales_tracker, 1, 42)

    update_cell(sandbox_sales_tracker, 'A4', org_variables['org_name'])
    update_cell(sandbox_sales_tracker, 'B4', other_variables['sale_value'])
    update_cell(sandbox_sales_tracker, 'C4', other_variables['number_of_days'])
    update_cell(sandbox_sales_tracker, 'D4', float(other_variables['number_of_days']) * 250)

    if other_variables['sale_type'] == 'Renewal':
        update_cell(sandbox_sales_tracker, 'G4', other_variables['discount_value'])
    elif org_variables['org_members'] == 'Yes':
        update_cell(sandbox_sales_tracker, 'F4', other_variables['discount_value'])

    update_cell(sandbox_sales_tracker, 'J4', other_variables['current_month'])
    update_cell(sandbox_sales_tracker, 'K4', fy_quarter(other_variables['current_month']))

    if org_variables['org_sector'] == 'Voluntary':
        update_cell(sandbox_sales_tracker, 'M4', 'Vol')
    elif org_variables['org_sector'] == 'Corporate':
        update_cell(sandbox_sales_tracker, 'M4', 'Corporate')
    elif org_variables['org_sector'] == 'Public':
        update_cell(sandbox_sales_tracker, 'M4', 'Statutory')

    update_cell(sandbox_sales_tracker, 'N4', org_variables['org_turnover_band'])

    if int(org_variables['org_no_of_volunteers']) > 1000:
        update_cell(sandbox_sales_tracker, 'P4', 'Y')
    else:
        update_cell(sandbox_sales_tracker, 'P4', 'N')

    update_cell(sandbox_sales_tracker, 'Q4', other_variables['sale_type'])

    if org_variables['org_members'] == 'Yes':
        update_cell(sandbox_sales_tracker, 'R4', 'mem')
    else:
        update_cell(sandbox_sales_tracker, 'R4', 'non')
         
    update_cell(sandbox_sales_tracker, 'X4', (date.today() + timedelta(days = 30)).strftime('%m-%Y'))
    update_cell(sandbox_sales_tracker, 'AB4', (date.today() + timedelta(days = 150)).strftime('%m-%Y'))
    update_cell(sandbox_sales_tracker, 'AF4', (date.today() + timedelta(days = 330)).strftime('%m-%Y'))
    update_cell(sandbox_sales_tracker, 'AJ4', (date.today() + timedelta(days = 360)).strftime('%m-%Y'))
    update_cell(sandbox_sales_tracker, 'Y4', 'expected')
    update_cell(sandbox_sales_tracker, 'AC4', 'expected')
    update_cell(sandbox_sales_tracker, 'AG4', 'expected')
    update_cell(sandbox_sales_tracker, 'AK4', 'expected')

    save_workbook(workbook, filepath)

    print('sales_tracker sheet updated')


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


# -----------------------------------------------------------
# For testing
# -----------------------------------------------------------


#import os
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

#update_sales_tracker(sandbox_workbook_path, 'sales_tracker', org_variables, other_variables)




    
    
