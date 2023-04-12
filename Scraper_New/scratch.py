from openpyxl import load_workbook, Workbook
from docx import Document
from pathlib import Path
import pdb
import pyodbc

import urllib
from bs4 import BeautifulSoup
from selenium import webdriver
from urllib.parse import urljoin
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from requests.auth import HTTPBasicAuth

import claimsfunctions
import grantsfunctions
import codecs
import doc_elements
import final_reports
import form_elements
import functions
import progress_reports
import pandas
import queries
import urllib.request
import re
import time
import requests



    path = Path('F:\TSREG')
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.DocTitle, Claims.DocNumber
        FROM Claims 
        INNER JOIN Grants ON Claims.GrantID = Grants.GrantID
        WHERE Status <> 7''').fetchall()  


with requests.Session() as session:

        response = functions.login(session)

        grant_list = grantsfunctions.get_grants(session, response, grant_name)
        final_list = final_reports.get_final_reports(session, response, grant_name)

        print(final_list)
        print(grant_list)

        if len(final_list) > 0:
                final = final_list[0] 

        if len(grant_list) > 0:
                grant = grant_list[0] 


        if final_reports.get_final_status(session, final) > 1: 
                grant_status = grantsfunctions.get_grant_status(session, grant)

                grant_desc = grantsfunctions.get_project_title_and_summary(session, grant)

                grant_title = grant_desc[0]
                grant_summary = grant_desc[1]

                grant_budget = grantsfunctions.get_budget(session, grant)
                budget_spent = final_reports.get_total_spent(session, final)

                pdb.set_trace()

                test = "test"
 
       


