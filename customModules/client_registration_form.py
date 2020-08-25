# -----------------------------------------------------------
# client_registration_form
#
# Pulls variables from client registration form word document using regexes
# Sets variables as key value pairs within org_variables & other_variables dictionaries
# Returns both dictionaries
#
# -----------------------------------------------------------


# External modules

import re, zipfile, xml.dom.minidom
from datetime import date


# Custom modules

from customModules.MSOffice.ms_word import set_up_doc_object
import customModules.twitter as twitter


# ---------------------------------------------------
# Client registration form functions
# -----------------------------------------------------------

# Pulls client response to form_field in registration form & assigns as value in key_name value pair in dictionary

def pullVariable(document, dictionary, form_field, key_name):

    # regex below looks for matches containing form_field: and any other characters - ie information client has entered in form field
    # Example would be - Organisation name: abc
    
    regex = re.compile(f'{form_field}: (.+)')
    
    for i in range(len(document.paragraphs)):            
        if form_field in document.paragraphs[i].text:

            # When querying regex with findall, only matches inside grouped parentheses are returned as client_response
            
            client_response = regex.findall(document.paragraphs[i].text)
            
            try:
                dictionary.setdefault(key_name, client_response[0])
            except Exception as e:
                dictionary.setdefault(key_name,'No value found')
                print('No value found for ' + key_name)
                print(str(e))


# Pulls client e-mail address from registration form & assigns to dictionary as key_value pair
# (docx doesn't pick up e-mail/hyperlinks. pullEmail uses XML instead)
                            
def pullEmail(dictionary, filepath):
    
    document = zipfile.ZipFile(filepath)
    uglyXml = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')
    text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
    prettyXml = text_re.sub('>\g<1></', uglyXml)
    reg = re.compile(r'>(.+@.+)<')
    
    try:
        dictionary.setdefault('person_email', reg.findall(prettyXml)[0])
    except Exception as e:
        dictionary.setdefault('person_email','No value found')
        print('No e-mail found from registration form')
        print(str(e))


    
# -----------------------------------------------------------
# Sets up variables in respective dictionaries
# -----------------------------------------------------------

def pull_org_and_other_variables(filepath):

    org_variables = {}
    other_variables = {}
    
    reg_doc = set_up_doc_object(filepath)

    # Pulls variables & creates key value pairs in org_variables

    pullVariable(reg_doc, org_variables,'Name of organisation', 'org_name')
    pullVariable(reg_doc, org_variables,'Postcode', 'org_postcode')
    pullVariable(reg_doc, org_variables,'Area of work', 'org_area_of_work')
    pullVariable(reg_doc, org_variables,'Head Office Region', 'org_head_office_region')
    pullVariable(reg_doc, org_variables,'Annual turnover', 'org_turnover_band')
    pullVariable(reg_doc, org_variables,'Number of volunteers - excluding trustees', 'org_no_of_volunteers')
    pullVariable(reg_doc, org_variables,'Number of volunteer roles - excluding trustee role', 'org_no_of_vol_roles')
    pullVariable(reg_doc, org_variables,'Number of sites/branches', 'org_no_of_sites')
    pullVariable(reg_doc, org_variables,'Telephone', 'person_telephone')

    # Contact name not used in pullVariable as two separate form_fields need to be concatenated before assigning to dictionary

    first_name_regex = re.compile('Contact person first name\(s\): (.+)')
    first_name = first_name_regex.findall(reg_doc.paragraphs[10].text)
    last_name_regex = re.compile('Contact person last name\(s\): (.+)')
    last_name = last_name_regex.findall(reg_doc.paragraphs[11].text)
    org_variables.setdefault('person_name', first_name[0] + ' ' + last_name[0])

    pullEmail(org_variables, filepath)

    org_variables.setdefault('org_twitter', twitter.pull_twitter_handle(org_variables['org_name']))
    
    print('Client registration form variables extracted')


    # Sets up date today & current month variables as key value pairs in other_variables

    other_variables.setdefault('date_today', date.today().strftime('%d/%m/%Y'))
    other_variables.setdefault('current_month', date.today().strftime('%m/%Y'))

    return [org_variables, other_variables]


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


# -----------------------------------------------------------
# For testing
# -----------------------------------------------------------

#import os
#from dotenv import load_dotenv
#load_dotenv()

#registration_form_path = os.environ.get('REGISTRATION_FORM_PATH')

#[org_variables, other_variables] = pull_org_and_other_variables(registration_form_path)

#print(org_variables)
#print(other_variables)




    
