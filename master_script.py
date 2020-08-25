
# -----------------------------------------------------------
# Automated Processing of Investing in Volunteers sales
#
# Through user input & client's registration form, automates completion of:
#
#       - sandbox IiV sales tracker & registered orgs excel spreadsheets
#       - assessor agreement template (word document)
#       - client contract (word document)
#
# Coded by Adam Fox (2019/20)
# adamfox03@gmail.com
# -----------------------------------------------------------


# External modules

import json, os, dotenv
from dotenv import load_dotenv


# Custom modules

from customModules import client_registration_form, user_input, reg_orgs_sheet, sales_tracker_sheet, aa_template, client_contract


# Filepaths 

load_dotenv()

sandbox_workbook_path = os.environ.get('SANDBOX_PATH')

registration_form_path = os.environ.get('REGISTRATION_FORM_PATH')

assessor_agreement_path = os.environ.get('ASSESSOR_AGREEMENT_PATH')

new_sale_contract_template_path = os.environ.get('NEW_SALE_CONTRACT_TEMPLATE_PATH')

renewal_sale_contract_template_path = os.environ.get('RENEWAL_SALE_CONTRACT_TEMPLATE_PATH')


# Sets up global assessor dictionary

assessors = json.loads(os.environ.get('ASSESSORS_AND_LEAD_ASSESSORS'))


# -----------------------------------------------------------
# Stage 1 - pulling variables from client registration form
# -----------------------------------------------------------

[org_variables, other_variables] = client_registration_form.pull_org_and_other_variables(registration_form_path)


# -----------------------------------------------------------
# Stage 2 - user input, setting remaining org & other variables
# -----------------------------------------------------------

[org_variables, other_variables] = user_input.request_user_input(org_variables, other_variables, assessors)


# -----------------------------------------------------------
# Stage 3 - updates reg orgs sheet in sandbox spreadsheet with appropriate variables
# -----------------------------------------------------------

reg_orgs_sheet.update_reg_orgs(sandbox_workbook_path, 'reg_orgs', org_variables, other_variables, assessors)


# -----------------------------------------------------------
# Stage 4 - updates sales tracker sheet in sandbox spreadsheet with appropriate variables
# -----------------------------------------------------------

sales_tracker_sheet.update_sales_tracker(sandbox_workbook_path, 'sales_tracker', org_variables, other_variables)


# -----------------------------------------------------------
# Stage 5 - writes to assessor agreement template
# -----------------------------------------------------------

aa_template.update_aa_template(assessor_agreement_path, org_variables, other_variables)


# -----------------------------------------------------------
# Stage 6 - writes to client contract template
# -----------------------------------------------------------

if other_variables['sale_type'] == 'Renewal':
    client_contract.update_contract(renewal_sale_contract_template_path, org_variables, other_variables)
elif other_variables['sale_type'] == 'New':
    client_contract.update_contract(new_sale_contract_template_path, org_variables, other_variables)


# -----------------------------------------------------------
# End
# -----------------------------------------------------------

print('Program finished')
















