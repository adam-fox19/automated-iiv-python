# -----------------------------------------------------------
# client_contract
#
# Draws on ms_word utility functions to update client contract.
# -----------------------------------------------------------


# External modules

import time
from datetime import date
from datetime import timedelta


# Custom modules

from customModules.MSOffice.ms_word import set_up_doc_object, find_replace, set_up_table, update_table_cell, correct_table_font, save_document


# -----------------------------------------------------------
# Updates client contract
# -----------------------------------------------------------

def update_contract(filepath, org_variables, other_variables):

    file = set_up_doc_object(filepath)

    # text replacements
    
    find_replace(file, 'XXX', org_variables['org_name'])
    find_replace(file, '? volunteers in ? roles, based at ? site', f"{org_variables['org_no_of_volunteers']} volunteers in {org_variables['org_no_of_vol_roles']} roles, based at {org_variables['org_no_of_sites']} sites")
    find_replace(file, 'sum of ? for the following ? days', f"sum of £{other_variables['sale_value']} + VAT for the following {other_variables['number_of_days']} days")
    find_replace(file, 'Date:', 'Date: ' + date.today().strftime('%d %B %Y'))
    
    # schedule table updates (expected timeframes for client)

    schedule_table = set_up_table(file, 1)

    update_table_cell(schedule_table, 0, 2, schedule_table.cell(0,2).text.replace('XXX', org_variables['org_name']))                 
    update_table_cell(schedule_table, 1, 2, date.today().strftime('%b - %y'))
    update_table_cell(schedule_table, 2, 2, (date.today() + timedelta(days = 60)).strftime('%b - %y'))
    update_table_cell(schedule_table, 3, 2, (date.today() + timedelta(days = 120)).strftime('%b - %y'))
    update_table_cell(schedule_table, 4, 2, (date.today() + timedelta(days = 150)).strftime('%b - %y'))
    update_table_cell(schedule_table, 5, 2, (date.today() + timedelta(days = 420)).strftime('%b - %y'))
    update_table_cell(schedule_table, 6, 2, (date.today() + timedelta(days = 450)).strftime('%b - %y'))
    update_table_cell(schedule_table, 7, 2, (date.today() + timedelta(days = 450)).strftime('%b %y') + ' - ' + (date.today() + timedelta(days = 540)).strftime('%b %y'))

    # corrects font in all cells & bolds 2nd column header only (org name) 
    
    correct_table_font(schedule_table, 0, 2, True)
    for i in range (1,8):
        correct_table_font(schedule_table, i, 2, False)

    save_document(file, filepath, org_variables['org_name'], 'Contract ready')
    
    
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

#new_sale_contract_template_path = os.environ.get('NEW_SALE_CONTRACT_TEMPLATE_PATH')

#renewal_contract_template_path = os.environ.get('RENEWAL_SALE_CONTRACT_TEMPLATE_PATH')

#update_contract(renewal_contract_template_path, org_variables, other_variables)

#update_contract(new_sale_template_path, org_variables, other_variables)
