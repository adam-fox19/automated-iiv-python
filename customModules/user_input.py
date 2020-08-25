# -----------------------------------------------------------
# user_input
#
# Requests user input & allocates responses to org_variables & other_variables dictionaries.
# -----------------------------------------------------------


# Calculates & returns sale discount following user input of sale value

def sale_discount(org_variables, other_variables):

    try:    
    
        if other_variables['sale_type'] == 'Renewal':
            discount = ((int(other_variables['sale_value']) * 100) / 85) * 0.15
        elif org_variables['org_members'] == 'Yes':
            discount = ((int(other_variables['sale_value']) * 100) / 90) * 0.1
        else:
            discount = 0

        return discount

    except Exception as e:
        print('Sale discount calculation hasn\'t worked')
        print(str(e))

      
# -----------------------------------------------------------
# Requests & allocates user input. Validates user input where appropriate.
# -----------------------------------------------------------

def request_user_input(org_variables, other_variables, assessors):
    
    # org_variables user input
    
    print('Are they NCVO members? Yes / No')
    member = input()    
    while member != 'Yes' and member != 'No':
        print('Please enter Yes or No')
        member = input()        
    org_variables.setdefault('org_members', member)


    print('Public, Voluntary or Corporate?')
    org_type = input()
    while org_type not in ['Public', 'Voluntary', 'Corporate']:
        print('Please enter Public, Voluntary or Corporate')
        org_type = input()
    org_variables.setdefault('org_sector', org_type)


    # other variables user input
    
    print('Is this a new sale or a renewal?')
    print('Enter: New or Renewal')
    sale_type = input()
    while sale_type != 'New' and sale_type != 'Renewal':
        print('Please enter New or Renewal')
        sale_type = input()    
    other_variables.setdefault('sale_type', sale_type)
    

    if other_variables['sale_type'] == 'Renewal':
        print('What number renewal is it? (1,2,3 etc)')
        other_variables.setdefault('renewal_number', input())
        
        
    print('How many days assessment? (e.g 3, 3.5)')
    other_variables.setdefault('number_of_days', input())


    print('What is the assessor\'s name?')
    assessor_name = input()
    while assessor_name not in list(assessors.keys()):
        print('Please re-enter assessor name - watch spelling')
        assessor_name = input()
    other_variables.setdefault('assessor_name', assessor_name)
    
    
    print('Sale value? (no VAT). Don\'t include £ sign.')
    sale = input()
    while sale[0] == '£':
        print('please don\'t include the currency of the sale value')
        sale = input()
    other_variables.setdefault('sale_value', sale)
    
    
    other_variables.setdefault('discount_value', sale_discount(org_variables, other_variables))

    
    print('UK wide? X if yes, blank if no')
    other_variables.setdefault('uk_wide', input())


    print('User input completed')
    return [org_variables, other_variables]


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


# -----------------------------------------------------------
# For testing
# -----------------------------------------------------------

#import os, json
#from dotenv import load_dotenv

#load_dotenv()

#assessors = json.loads(os.environ.get('ASSESSORS_AND_LEAD_ASSESSORS'))


#[org_variables, other_variables] = request_user_input({},{}, assessors)

#print(org_variables)
#print(other_variables)



 
