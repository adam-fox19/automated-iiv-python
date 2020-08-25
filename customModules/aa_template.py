# -----------------------------------------------------------
# aa_template
#
# Draws on ms_word utility functions to update assessor agreement template.
# -----------------------------------------------------------


# External modules

import time
from datetime import date, timedelta


# Custom modules

from customModules.MSOffice.ms_word import set_up_doc_object, append_text, set_up_table, update_table_cell, bold_cell, save_document

# -----------------------------------------------------------
# Updates assessor agreement template
# -----------------------------------------------------------

def update_aa_template(filepath, org_variables, other_variables):
        
    file = set_up_doc_object(filepath)

    # text updates
            
    append_text(file, 'Assessor:', other_variables['assessor_name'])    
    append_text(file, 'Organisation to be assessed:', org_variables['org_name'])
    append_text(file, 'Number of volunteers:', org_variables['org_no_of_volunteers'])
    append_text(file, 'Number of sites:', org_variables['org_no_of_sites'])
    append_text(file, 'Postcode of main site:', org_variables['org_postcode'])

    if other_variables['sale_type'] == 'Renewal':
        append_text(file,'Renewal (number):', other_variables['renewal_number'])
    elif other_variables['sale_type'] == 'New':
        append_text(file, 'Renewal (number):', r'N/A - new sale')

            
    # fees table updates (number of days allocated for interviews & overall assessment, & corresponding fees)

    fees_table = set_up_table(file, 0)

    update_table_cell(fees_table, 3, 1, str(float(other_variables['number_of_days']) - 2))               #interview
    update_table_cell(fees_table, 5, 1, other_variables['number_of_days'])                               #overall assessment
    update_table_cell(fees_table, 3, 2, '£' + str(int(float(fees_table.cell(3,1).text) * 250)))          #interview fee
    update_table_cell(fees_table, 5, 2, '£' + str(int(float(other_variables['number_of_days']) * 250)))  #overall fee

    bold_cell(fees_table, 5, 1)
    bold_cell(fees_table, 5, 2)


    # schedule table updates (expected timeframes for client) 

    schedule_table = set_up_table(file, 1)

    update_table_cell(schedule_table, 1, 2, date.today().strftime('%b - %y'))
    update_table_cell(schedule_table, 2, 2, (date.today() + timedelta(days = 60)).strftime('%b - %y'))
    update_table_cell(schedule_table, 3, 2, (date.today() + timedelta(days = 120)).strftime('%b - %y'))
    update_table_cell(schedule_table, 4, 2, (date.today() + timedelta(days = 150)).strftime('%b - %y'))
    update_table_cell(schedule_table, 5, 2, (date.today() + timedelta(days = 420)).strftime('%b - %y'))
    update_table_cell(schedule_table, 6, 2, (date.today() + timedelta(days = 450)).strftime('%b - %y'))
    update_table_cell(schedule_table, 7, 2, (date.today() + timedelta(days = 450)).strftime('%b %y') + ' - ' + (date.today() + timedelta(days = 540)).strftime('%b %y'))

    save_document(file, filepath, org_variables['org_name'], 'Assessor agreement ready')
    

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


#assessor_agreement_path = os.environ.get('ASSESSOR_AGREEMENT_PATH')

#update_aa_template(assessor_agreement_path, org_variables, other_variables)



    
