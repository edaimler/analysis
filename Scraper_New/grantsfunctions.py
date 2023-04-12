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
import pandas
import urllib.request
import re
import time
import requests


gohsform = form_elements.gohsform
url = 'insert url'


def open_grants(session, response):
        link = BeautifulSoup(response.text, 'html.parser').find('a', href="Menu_Person2.aspx?NavItem1=3&NavItemID1=76546") # TODO: change href to id, because I think this href is unique to one grantee
        span = link.find('span', text='Applications')
        #print("link text is: ", span.text)
        postnext = urljoin(url, link['href'])
        time.sleep(3)
        r_next = session.post(postnext)

        return r_next


def get_grants(session, response, grant_name):
        grants = open_grants(session, response)

        response2 = functions.fetch(session, grants.url)
        form2 = BeautifulSoup(response2, 'html.parser').find(gohsform['tag'], id=gohsform['id'])
        formdata2 = functions.java_form(response2, gohsform, '', grant_name)

        for i in form_elements.pop_list:
                formdata2.pop(i)  # TODO REMOVE 'CLEAR' value from form

        grants_url = urljoin(url, form2['action']) # get the URL of progressreports page
        grant = session.post(grants_url, data = formdata2) 
        grant_list = functions.doc_list(grant.text)
        time.sleep(3)

        return grant_list


def get_grant_status(session, grant):
        link = functions.open_link(session, grant)
        statustable = BeautifulSoup(link.text, 'html.parser').find('table', id='ctl00_cphPageContent_wclDocumentInformation_dgdDocumentInformation')
        statusframe = pandas.read_html(str(statustable))[0]
        status = statusframe.iloc[1,4] 
        time.sleep(5)
        if status.find('Executed') > 0: 
                status_id = 2 
        elif status.find('Closeout in Process') > 0:
                status_id = 3 
        elif status.find('Grant Closed') > 0: 
                status_id = 4 
        else:
                status_id = 1

        return status_id


def get_project_title_and_summary(session, grant):
        link = functions.open_link(session, grant)
        time.sleep(3)
        link2 = BeautifulSoup(link.text, 'html.parser').find('a', text="Application Project Information")
        postnext = urljoin(url, link2['href'])
        info = session.post(postnext)
        time.sleep(3)
        title = BeautifulSoup(info.text, 'html.parser').find(doc_elements.project_title['tag'], id=doc_elements.project_title['id'])['value']
        summary = BeautifulSoup(info.text, 'html.parser').find(doc_elements.project_summary['tag'], id=doc_elements.project_summary['id']).text
        summary = summary.replace('\r\n', '')

        return (title, summary)


#def get_funding_source(grant_name): # for this function, create a list of all funding sources, then search that list using the grant name
                                        # Example: GA-2020-405b M1*OP OP HIGH(2020)-003 uses funding source 405B M1*OP



def get_budget(session, grant):
        link = functions.open_link(session, grant)
        time.sleep(3)
        #cost_category = functions.open_id(session, link.text, doc_elements.cost_category)
        link2 = BeautifulSoup(link.text, 'html.parser').find('a', text=doc_elements.cost_category['text'])

        postnext = urljoin(url, link2['href'])
        cost_category = session.post(postnext)
        time.sleep(3)
        total_budget = BeautifulSoup(cost_category.text).find(doc_elements.total_budget['tag'], id=doc_elements.total_budget['id']).text

        return total_budget