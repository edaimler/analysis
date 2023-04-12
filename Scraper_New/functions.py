

#import urllib

from bs4 import BeautifulSoup
#from selenium import webdriver
from urllib.parse import urljoin
from openpyxl import load_workbook, Workbook
#from openpyxl.worksheet.table import Table, TableStyleInfo
from requests.auth import HTTPBasicAuth

import form_elements
import functions
import pandas
#import urllib.request
import re
import time
#import requests



gohsform = form_elements.gohsform
username = form_elements.username
password = form_elements.password
eventtgt = form_elements.eventtgt
timeoutredirect = form_elements.timeoutredirect
pagecontent = form_elements.timeoutredirect
searchdocumenttype = form_elements.searchdocumenttype
searchdocumentstatus = form_elements.searchdocumentstatus
email = form_elements.email
displaymessage = form_elements.displaymessage
eventtgt = form_elements.eventtgt
eventagmt = form_elements.eventagmt
lastfocus = form_elements.lastfocus
txtsubject = form_elements.txtsubject
txtbody = form_elements.txtbody
year = form_elements.year
intelli = form_elements.intelli
reminderredirect = form_elements.reminderredirect
alertmsg = form_elements.alertmsg
sortable = form_elements.sortable
sortorder = form_elements.sortorder
finalreportfield = form_elements.finalreportfield
person = form_elements.person
organization = form_elements.organization
doctitle = form_elements.doctitle
option = form_elements.option

url = 'insert url'

def fetch(session, url, data=None):
    if data is None:
        print("fetching...")
        time.sleep(15)
        return session.get(url).content
    else:
        print("else fetching...")
        time.sleep(15)
        return session.post(url, data=data).content

def login(session):
    response = functions.fetch(session, url)
    user = functions.enter_field(response, gohsform, username, "username")  # update formdata with username
    newuser = {
        username['name']: user[username['name']]
    }
    pw = functions.enter_field(response, gohsform, password, "password")  # update formdata with password
    pw['txtPassword'] = "password"
    newpw = {
        password['name']: pw[password['name']]
    }
    defaultfields = functions.default_fields(response, gohsform)

    newuser.update(newpw)
    newuser.update(defaultfields)
    submit = BeautifulSoup(response, 'html.parser').find(gohsform['tag'], id=gohsform['id'])
    postinfo = urljoin(url, submit['action'])
    time.sleep(15)
    r = session.post(postinfo, data=newuser)
    time.sleep(15)

    return r

def enter_field(response, webform, fielddict, entry):
    form = BeautifulSoup(response, 'html.parser').find(webform['tag'], id=webform['id'])
    fields = form.findAll(fielddict['tag'], id=fielddict['id'])
    formdata = dict( (field.get('name'), field.get('value')) for field in fields)
    formdata[fielddict['name']] = entry
    
    defaultfields = default_fields(response, webform)
    formdata.update(defaultfields)
    
    #print("all data is: ", formdata)

    return formdata

def default_fields(response, webform):
    time.sleep(5)
    form = BeautifulSoup(response, 'html.parser').find(webform['tag'], id=webform['id'])
    fields = form.findAll("input")
    formdata = dict( (field.get('name'), field.get('value')) for field in fields)
    # print("default fields are: ", formdata)

    return formdata


def prep_form(response, webform):
    form = BeautifulSoup(response, 'html.parser').find(webform['tag'], id=webform['id'])
    fields = form.findAll("input")
    formdata = dict( (field.get('name'), field.get('value')) for field in fields)
    #print("formdata2 is: ", formdata)
    formdata[pagecontent['name']] = "Search"
    defaultfields = functions.default_fields(response, webform)
    formdata.update(defaultfields)
    formdata[eventtgt['name']] = ""
    formdata[eventagmt['name']] = ""
    formdata[lastfocus['name']] = ""
    formdata[intelli['name']] = ""
    formdata[txtsubject['name']] = ""
    formdata[txtbody['name']] = ""
    formdata[reminderredirect['name']] = ""
    formdata[alertmsg['name']] = ""
    formdata[sortable['name']] = ""
    formdata[sortorder['name']] = ""
    formdata[finalreportfield['name']] = ""
    formdata[person['name']] = ""
    formdata[organization['name']] = ""
    formdata[year['name']] = "2020"
    formdata[searchdocumenttype['name']] = '0'
    formdata[searchdocumentstatus['name']] = '0'
    formdata[email['name']] = 'Email'
    formdata[displaymessage['name']] = '0'

    return formdata


def java_form(response, webform, link, doc):
    form = BeautifulSoup(response, 'html.parser').find(webform['tag'], id=webform['id'])
    fields = form.findAll("input")
    formdata = dict( (field.get('name'), field.get('value')) for field in fields)
    #print("formdata3 is: ", formdata)
    formdata[pagecontent['name']] = "Search"
    defaultfields = functions.default_fields(response, webform)
    formdata.update(defaultfields)
    formdata[eventtgt['name']] = link
    formdata[eventagmt['name']] = ""
    formdata[lastfocus['name']] = ""
    formdata[intelli['name']] = ""
    formdata[txtsubject['name']] = ""
    formdata[txtbody['name']] = ""
    formdata[reminderredirect['name']] = ""
    formdata['ctl00$hdnSystemPage'] = "Menu_Person2.aspx_3"
    formdata[alertmsg['name']] = ""
    formdata[sortable['name']] = ""
    formdata[sortorder['name']] = ""
    formdata[finalreportfield['name']] = doc
    formdata[person['name']] = ""
    formdata[organization['name']] = ""
    formdata[year['name']] = ""
    formdata[searchdocumenttype['name']] = '0'	
    formdata[searchdocumentstatus['name']] = '0'
    formdata[email['name']] = 'Email'
    formdata[displaymessage['name']] = '0'

    return formdata


def option_form(response, webform, link, doc, activity):
    form = BeautifulSoup(response, 'html.parser').find(webform['tag'], id=webform['id'])
    fields = form.findAll("input")
    formdata = dict( (field.get('name'), field.get('value')) for field in fields)
    #print("formdata3 is: ", formdata)
    formdata[pagecontent['name']] = "Search"
    defaultfields = functions.default_fields(response, webform)
    formdata.update(defaultfields)
    formdata[eventtgt['name']] = link
    formdata[eventagmt['name']] = ""
    formdata[lastfocus['name']] = ""
    formdata[intelli['name']] = ""
    formdata[txtsubject['name']] = ""
    formdata[txtbody['name']] = ""
    formdata[reminderredirect['name']] = ""
    formdata['ctl00$hdnSystemPage'] = "Menu_Person2.aspx_3"
    formdata[alertmsg['name']] = ""
    formdata[sortable['name']] = ""
    formdata[sortorder['name']] = ""
    formdata[finalreportfield['name']] = doc
    formdata[person['name']] = ""
    formdata[organization['name']] = ""
    formdata[year['name']] = ""
    formdata[searchdocumenttype['name']] = '0'	
    formdata[searchdocumentstatus['name']] = '0'
    formdata[email['name']] = 'Email'
    formdata[displaymessage['name']] = '0'
    formdata[option['name']] = activity

    return formdata

def open_final_reports(session, response):
    link = BeautifulSoup(response, 'html.parser').find('a', href="Menu_Person2.aspx?NavItem1=8&NavItemID1=76551")  # TODO: change href to id, because I think this href is unique to one grantee
    span = link.find('span', text='Final Reports')
    #print("link text is: ", span.text)
    postnext = urljoin(url, link['href'])
    time.sleep(4)
    r_next = session.post(postnext)

    return r_next


def doc_id(tag):
    if tag.has_attr('id'):
        return (tag['id'].__contains__("hplDocumentName") )
    else: return False

def org(tag):
    if tag.has_attr('id'):
        return (tag['id'].__contains__("hplOrganization") )
    else: return False

def doc_list(response):
    return BeautifulSoup(response).find_all(doc_id)


def tag_text(response, data):
    tag = BeautifulSoup(response, 'html.parser').find(data['tag'], id=data['id'])

    return tag.text
    
def open_link(session, tag, id:bool = False):
    return session.post(urljoin(url, tag['href']))

def open_id(session, response, tag):
    link = BeautifulSoup(response, 'html.parser').find('a', id = tag['id'])
    #print("link is: ", link)
    postnext = urljoin(url, link['href'])
    time.sleep(4)
    r_next = session.post(postnext)

    return r_next

def change_asterisk(classification, save):
    result = classification
    if save == True and classification.find('*') is not None:
        result = classification.replace('*', '#')
    if save == False and classification.find('#') is not None:
        result = classification.replace('#', '*')

    return result

def parse_doc_title(response):
    parsed = response.split("-")

    document = {
    "docname": response,
    "program": parsed[0],
    "classification": change_asterisk(parsed[2], True),
    "year": parsed[1],
    "docnumber": parsed[4],
    "id": parsed[3]
    }
    
    return document



def list_result_pages(response):
    page_list = BeautifulSoup(response, 'html.parser').find_all('a', href = re.compile("ctl00\$cphPageContent\$wclDocuments\$dgdMyDocuments\$ctl"))
    print("page list is: ", page_list)
    if len(page_list) == 0:
        print("page list failed")
        return page_list
    else:
        return page_list

