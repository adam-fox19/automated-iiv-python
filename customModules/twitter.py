
# -----------------------------------------------------------
# twitter
#
# Automated search & return of organisation's twitter handle 
# -----------------------------------------------------------


# External modules

from selenium import webdriver
from googlesearch import search


# -----------------------------------------------------------
# Runs a google search (on Firefox) of the org_name variable concatenated with 'twitter', and loads first result
# Once org twitter page has rendered, searches page for span elements (org twitter handle is contained in span)
# Finds first span element's inner text with the first character of '@' - this is the org twitter handle & is returned
# -----------------------------------------------------------


def pull_twitter_handle(org_name):

    browser = webdriver.Firefox()
    searchResults = []
    
    for url in search(org_name + ' twitter', stop = 5):
            searchResults.append(url)
    
    browser.get(searchResults[0])

    span_inner_texts = browser.find_elements_by_tag_name('span')

    try:
        for x in span_inner_texts:
            if len(x.text) > 0:       #omits any span elements with no inner text, speeds up loop
                if x.text[0] == '@':
                    twitter_handle = x.text
                    browser.quit()
                    return twitter_handle
    except Exception as e:
        print('No twitter handle found')
        twitter_handle = 'No value found'
        browser.quit()
        print(str(e))
        return twitter_handle


# -----------------------------------------------------------
# End
# -----------------------------------------------------------


# -----------------------------------------------------------
# For testing
# -----------------------------------------------------------

#org_twitter = pull_twitter_handle('Oxfam')
#print(org_twitter)
