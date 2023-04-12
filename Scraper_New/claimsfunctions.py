import pyodbc
import pdb

import urllib
from bs4 import BeautifulSoup
from selenium import webdriver
from urllib.parse import urljoin
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from requests.auth import HTTPBasicAuth

import codecs
import datetime
import doc_elements
import form_elements
import functions
import numpy as np
import pandas
import urllib.request
import re
import time
import requests

gohsform = form_elements.gohsform
url = 'https://georgia.intelligrants.com/login2.aspx'

def open_claims(session, response):
    link = BeautifulSoup(response.text, 'html.parser').find('a', href="Menu_Person2.aspx?NavItem1=7&NavItemID1=76550") # TODO: change href to id, because I think this href is unique to one grantee
    span = link.find('span', text='Claims')
    postnext = urljoin(url, link['href'])
    print("url is: ", postnext)
    time.sleep(13)
    r_next = session.post(postnext)
    time.sleep(1)
    return r_next

def get_claims(session, response, grant_name):
    claimsvar = open_claims(session, response)
    response3 = functions.fetch(session, claimsvar.url)
    form3 = BeautifulSoup(response3).find(gohsform['tag'], id=gohsform['id'])
    formdata3 = functions.java_form(response3, gohsform, '', grant_name)

    for i in form_elements.pop_list:
            formdata3.pop(i)  # TODO REMOVE 'CLEAR' value from form

    time.sleep(3)
    claimurl = urljoin(claimsvar.url, form3['action']) # get the URL of progressreports page
    claimreport = session.post(claimurl, data = formdata3) 
    time.sleep(3)

    claimlist = functions.doc_list(claimreport.text)

    return claimlist


def get_claim_status(session, claim):
    link = functions.open_link(session, claim)
    mytable = BeautifulSoup(link.text).find(doc_elements.claim_status['tag'], id=doc_elements.claim_status['id'])
    time.sleep(10)
    statusframe = pandas.read_html(str(mytable))[0]
    status = statusframe.iloc[1,4] 

    if status.find('Process') >= 0: 
        status_id = 2 
    elif status.find('Submitted') >= 0:
        status_id = 3 
    elif status.find('Claim') >= 0 and status.find('Approved') >= 0: 
        status_id = 4 
    elif status.find('Payment') >= 0 and status.find('Approved') >= 0:
        status_id = 5
    elif status.find('Payment') >= 0 and status.find('Complete') >= 0:
        status_id = 6 
    elif status.find('Payment') >= 0 and status.find('Created') >= 0:
        status_id = 7 
    elif status.find('Fiscal') >= 0: 
        status_id = 8
    else:
        status_id = 1

    return status_id

def last_modification(session, claim):
    link = functions.open_link(session, claim)
    mytable = BeautifulSoup(link.text).find('table', id="ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu")
    time.sleep(10)
    statusframe = pandas.read_html(str(mytable))[0] 
    fourth_column = statusframe.iloc[2:,4].to_list()
    fifth_column = statusframe.iloc[2:,5].to_list() 
    entries = fourth_column + fifth_column
    datelist = []
    for entry in entries: 
            if type(entry) != float: 
                    fields = entry.split(' ') 
                    if len(fields) > 3:
                        dates = fields[-3].split('/')
                        datelist = datelist + [datetime.date( int(dates[2]), int(dates[0]), int(dates[1]) )]
    print("datelist is: ", datelist)

    return max(datelist)


def currency_parser(cur_str):
    # Remove any non-numerical characters
    # except for ',' '.' or '-' (e.g. EUR)
    cur_str = re.sub("[^-0-9.,]", '', cur_str)
    # Remove any 000s separators (either , or .)
    cur_str = re.sub("[.,]", '', cur_str[:-3]) + cur_str[-3:]

    if '.' in list(cur_str[-3:]):
        num = float(cur_str)
    elif ',' in list(cur_str[-3:]):
        num = float(cur_str.replace(',', '.'))
    else:
        num = float(cur_str)

    return np.round(num, 2)

def get_current_expense(session, claim):
    link = functions.open_link(session, claim)
    time.sleep(13)
    exp = BeautifulSoup(link.text, 'html.parser').find('a', text = 'Expense Summary')
    #print("link is: ", link)
    postnext = urljoin(url, exp['href'])
    time.sleep(4)
    expense = session.post(postnext)
    expensetable = BeautifulSoup(expense.text).find_all(doc_elements.expense_table['tag'], id=re.compile(doc_elements.expense_table['id']))
    expenseframe = pandas.read_html(str(expensetable[7]))[0]
    current_expense1 = expenseframe.iloc[13,5]
    current_expense2 = expenseframe.iloc[12,5]
    current_expense3 = expenseframe.iloc[11,5]
    current_expense = max(current_expense1, current_expense2, current_expense3)
    if type(current_expense) == float:
        current_expense = '$0.0'
        print("current expense is: ", current_expense)
    print("current expense is: ", current_expense)

    return currency_parser(current_expense)