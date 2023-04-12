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



def open_progress_reports(session, response):
    link = BeautifulSoup(response.text, 'html.parser').find('a', href="Menu_Person2.aspx?NavItem1=4&NavItemID1=76547") # TODO: change href to id, because I think this href is unique to one grantee
    span = link.find('span', text='Progress Reports')
    #print("link text is: ", span.text)
    postnext = urljoin(url, link['href'])
    time.sleep(10)
    r_next = session.post(postnext)

    return r_next

def get_progress_reports(session, response, grant_name):
    progressreports = open_progress_reports(session, response)
    
    response2 = functions.fetch(session, progressreports.url)
    form2 = BeautifulSoup(response2, 'html.parser').find(gohsform['tag'], id=gohsform['id'])
    formdata2 = functions.java_form(response2, gohsform, '', grant_name)

    for i in form_elements.pop_list:
        formdata2.pop(i)  # TODO REMOVE 'CLEAR' value from form

    progressurl = urljoin(url, form2['action']) # get the URL of progressreports page
    progressreport = session.post(progressurl, data = formdata2) 
    proglist = functions.doc_list(progressreport.text)
    time.sleep(10)

    return proglist


def get_report_status(session, report):
    link = functions.open_link(session,report)
    statustable = BeautifulSoup(link.text, 'html.parser').find('table', id='ctl00_cphPageContent_wclDocumentInformation_dgdDocumentInformation')
    statusframe = pandas.read_html(str(statustable))[0]
    status = statusframe.iloc[1,4] 
    time.sleep(5)
    if status.find('Process') > 0: 
        status_id = 2 
    elif status.find('Submitted') > 0:
        status_id = 3 
    elif status.find('Approved') > 0: 
        status_id = 4 
    else:
        status_id = 1

    return status_id

def last_modification(session, report):
    link = functions.open_link(session, report)
    mytable = BeautifulSoup(link.text, 'html.parser').find('table', id="ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu")
    time.sleep(3)
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

def get_agency_personnel(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    monthly = functions.open_id(session, link.text, doc_elements.monthly)
    monthlytable = BeautifulSoup(monthly.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    time.sleep(3)
    monthlyframe = pandas.read_html(str(monthlytable))[0]

    return monthlyframe
    

def get_narratives(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    narrative = functions.open_id(session, link.text, doc_elements.narrative)
    narrativetable = BeautifulSoup(narrative.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    time.sleep(3)
    narrativeframe = pandas.read_html(str(narrativetable))[0]

    accomplishment_narrative = narrativeframe.iloc[1][0]
    challenges_narrative = narrativeframe.iloc[4][0]

    return (accomplishment_narrative, challenges_narrative)

def get_enforcement(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    enforcement = functions.open_id(session, link.text, doc_elements.enforcement)
    time.sleep(3)
    enforcementtable = BeautifulSoup(enforcement.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    time.sleep(3)
    enforcementframe = pandas.read_html(str(enforcementtable))[0]
    #THIS FUNCTION IS UNFINISHED.  COMPLETE THIS FUNCTION IF/WHEN ENFORCEMENT DATA ARE REQUIRED
    return enforcementframe

def get_milestone_inputs(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    milestone = functions.open_id(session, link.text, doc_elements.milestone)
    table = BeautifulSoup(milestone.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    time.sleep(3)
    if table == None:
        table = BeautifulSoup(milestone.text, 'html.parser').find('table', id="ctl00_cphPageContent_tblPageTable13113")
    if table == None:
        table = BeautifulSoup(milestone.text, 'html.parser').find('table', id="ctl00_cphPageContent_tblPageTable13823")
    
    valueoct = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intOCT_'))
    valuenov = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intNOV_'))
    valuedec = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intDEC_'))
    valuejan = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intJAN_'))
    valuefeb = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intFEB_'))
    valuemar = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intMAR_'))
    valueapr = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intAPR_'))
    valuemay = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intMAY_'))
    valuejun = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intJUN_'))
    valuejul = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intJUL_'))
    valueaug = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intAUG_'))
    valuesep = table.find_all('input', id = re.compile('^ctl00_cphPageContent_intSEP_'))

    months = [valueoct, valuenov, valuedec, valuejan, valuefeb, valuemar, valueapr, valuemay, valuejun, valuejul, valueaug, valuesep]
    all_milestones = []

    for month in months:
        milestones = list(value.get('value') for value in month)
        print('milestones are: ', milestones)
        all_milestones.append(milestones)

    df = pandas.DataFrame(all_milestones)
    inputs_frame = df.transpose()

    return inputs_frame

def get_milestone_targets(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    milestone = functions.open_id(session, link.text, doc_elements.milestone)
    time.sleep(3)
    milestonetable = BeautifulSoup(milestone.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    if milestonetable == None:
        milestonetable = BeautifulSoup(milestone.text, 'html.parser').find('table', id="ctl00_cphPageContent_tblPageTable13113")
    if milestonetable == None:
        milestonetable = BeautifulSoup(milestone.text, 'html.parser').find('table', id="ctl00_cphPageContent_tblPageTable13823")
    milestoneframe = pandas.read_html(str(milestonetable))[0]

    valueoct = milestoneframe.iloc[1:][1].to_list()
    valuenov = milestoneframe.iloc[1:][2].to_list()
    valuedec = milestoneframe.iloc[1:][3].to_list()
    valuejan = milestoneframe.iloc[1:][4].to_list()
    valuefeb = milestoneframe.iloc[1:][5].to_list()
    valuemar = milestoneframe.iloc[1:][6].to_list()
    valueapr = milestoneframe.iloc[1:][7].to_list()
    valuemay = milestoneframe.iloc[1:][8].to_list()
    valuejun = milestoneframe.iloc[1:][9].to_list()
    valuejul = milestoneframe.iloc[1:][10].to_list()
    valueaug = milestoneframe.iloc[1:][11].to_list()
    valuesep = milestoneframe.iloc[1:][12].to_list()

    months = [valueoct, valuenov, valuedec, valuejan, valuefeb, valuemar, valueapr, valuemay, valuejun, valuejul, valueaug, valuesep]
    monthsframe = pandas.DataFrame(months).transpose()

    return monthsframe

def get_statistics(session, report):
    link = functions.open_link(session, report)
    time.sleep(3)
    statistical = functions.open_id(session, link.text, doc_elements.statistical)
    statisticaltable = BeautifulSoup(statistical.text, 'html.parser').find('table', id=re.compile('^ctl00_cphPageContent_tblPageTable'))
    time.sleep(3)
    statisticalframe = pandas.read_html(str(statisticaltable))[0]
    
    return statisticalframe