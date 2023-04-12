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


def open_final_reports(session, response):
        link = BeautifulSoup(response.text, 'html.parser').find('a', href="Menu_Person2.aspx?NavItem1=8&NavItemID1=76551") # TODO: change href to id, because I think this href is unique to one grantee
        span = link.find('span', text='Final Reports')
        #print("link text is: ", span.text)
        postnext = urljoin(url, link['href'])
        time.sleep(3)
        r_next = session.post(postnext)

        return r_next


def get_final_reports(session, response, grant_name):
        finalreports = open_final_reports(session, response)
        
        response2 = functions.fetch(session, finalreports.url)
        form2 = BeautifulSoup(response2, 'html.parser').find(gohsform['tag'], id=gohsform['id'])
        formdata2 = functions.java_form(response2, gohsform, '', grant_name)

        for i in form_elements.pop_list:
                formdata2.pop(i)  # TODO REMOVE 'CLEAR' value from form

        finalurl = urljoin(url, form2['action']) # get the URL of progressreports page
        finalreport = session.post(finalurl, data = formdata2) 
        final_list = functions.doc_list(finalreport.text)
        time.sleep(10)

        return final_list


def get_final_status(session, final):
        link = functions.open_link(session, final)
        statustable = BeautifulSoup(link.text, 'html.parser').find('table', id='ctl00_cphPageContent_wclDocumentInformation_dgdDocumentInformation')
        statusframe = pandas.read_html(str(statustable))[0]
        status = statusframe.iloc[1,4] 
        time.sleep(5)
        if status.find('Submitted') > 0: 
                status_id = 2 
        elif status.find('Approved') > 0:
                status_id = 3 
        else:
                status_id = 1

        return status_id


def last_modification(session, final):
        link = functions.open_link(session, final)
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


def get_total_spent(session, final): 
        link1 = functions.open_link(session, final)
        time.sleep(3)
        link = BeautifulSoup(link1.text, 'html.parser').find('a', text = doc_elements.final_expenditures['text'])
        #print("link is: ", link)
        postnext = urljoin(url, link['href'])
        time.sleep(3)
        expenditures = session.post(postnext)



        #expenditures = functions.open_id(session, link.text, doc_elements.final_expenditures)
        time.sleep(3)  
        total_spent = BeautifulSoup(expenditures.text, 'html.parser').find(doc_elements.total_spent['tag'], id=doc_elements.total_spent['id']).text

        return total_spent



def get_activities_and_results(session, final, grant_name):
        link = functions.open_link(session, final)
        time.sleep(3)
        link2 = BeautifulSoup(link.text, 'html.parser').find('a', text=doc_elements.final_activities['text'])
        postnext = urljoin(url, link2['href'])
        activities = session.post(postnext)
        time.sleep(3)
        select = BeautifulSoup(activities.text).find(doc_elements.select_activity['tag'], title=doc_elements.select_activity['title'])
        if not select: 
                link2 = BeautifulSoup(link.text, 'html.parser').find('a', text=doc_elements.ya_final_activities['text'])
                postnext = urljoin(url, link2['href'])
                activities = session.post(postnext)
                time.sleep(3)
                select = BeautifulSoup(activities.text).find(doc_elements.select_activity['tag'], title=doc_elements.select_activity['title'])
        activity_list = select.find_all()
        form2 = BeautifulSoup(activities.text, 'html.parser').find(gohsform['tag'], id=gohsform['id'])

        all_activities = []

        for activity in activity_list: 
                formdata2 = functions.option_form(activities.text, gohsform, '', grant_name, activity['value'])

                for i in form_elements.pop_list2:
                        formdata2.pop(i)  # TODO REMOVE 'CLEAR' value from form

                next_activity = urljoin(url, form2['action']) # get the next activity
                page = session.post(next_activity, data = formdata2) 
                        
                activity_funded = BeautifulSoup(page.text).find_all(doc_elements.activity_funded['tag'], id=doc_elements.activity_funded['id'])
                sadd_activity_funded = BeautifulSoup(page.text).find_all(doc_elements.sadd_activity_funded['tag'], id=doc_elements.sadd_activity_funded['id'])
                try:
                        activity_funded = activity_funded[0].text
                except: 
                        activity_funded = sadd_activity_funded[0].text

                desc = BeautifulSoup(page.text).find_all(doc_elements.activity_desc['tag'], id=doc_elements.activity_desc['accomplishments'])[0].text

                desc = desc.replace('\r\n', '')
                if len(desc) < 1: 
                    desc = BeautifulSoup(page.text).find_all(doc_elements.activity_desc['tag'], id=doc_elements.activity_desc['challenges'])[0].text

                desc = desc.replace('\r\n', '')
                all_activities.append((activity_funded, desc))

        return all_activities