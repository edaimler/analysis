from openpyxl import load_workbook, Workbook
from pathlib import Path
import pdb
import pyodbc
from datetime import datetime
import urllib
from bs4 import BeautifulSoup
from selenium import webdriver
from urllib.parse import urljoin
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from requests.auth import HTTPBasicAuth

import claimsfunctions
import codecs
import doc_elements
import form_elements
import functions
import progress_reports
import pandas
import queries
import urllib.request
import re
import time
import requests

orgs = ["GA-2022-402 OP-039-C3"
        ]

gohsform = form_elements.gohsform
url = 'https://georgia.intelligrants.com/login2.aspx'

with requests.Session() as s:

        r = functions.login(s)

        '''
        doclist = queries.get_prog_report_titles()
        
        doc_num_test = queries.get_report_number(doclist[0])
        print("doc_num_test is: ", doc_num_test)
        print("Doc size is: ", len(doclist))
        for doc in doclist:
                doc_year = functions.parse_doc_title(doc)['year']
                print("document title is: ", doc)
                report = progress_reports.get_progress_reports(s, r, doc)
                if len(report) > 0:
                        report = report[0] 
                elif int(doc_year) < 2022: 
                        eight_tuple = ('', '0', '', '[]', '[]', '0', '5', str(datetime(1,1,1)))
                        doc_num = queries.get_report_number(doc)
                        queries.update_progress_report(doc_num, eight_tuple)
                        print("No report: ", doc)
                        continue
                else: 
                        print("No report found: ", doc)
                        continue
                status = progress_reports.get_report_status(s, report)

                mod = progress_reports.last_modification(s, report) 

                persons = progress_reports.get_agency_personnel(s, report) 

                narr = progress_reports.get_narratives(s, report) 

                enf = progress_reports.get_enforcement(s, report)

                mile_inputs = progress_reports.get_milestone_inputs(s, report) 

                targets = progress_reports.get_milestone_targets(s, report) 

                stats = progress_reports.get_statistics(s, report) 


                report_num = int(doc.split('-')[-1].replace('R', ''))
                mile_column = mile_inputs.iloc[:,report_num-1].to_list()
                targets_column = targets.iloc[:,report_num-1].to_list()
                try: 
                        stats = stats.iloc[1:,report_num].to_list() 
                except: 
                        stats = []
                
                eight_tuple = (narr[0], narr[1], str(stats), str(targets_column), str(mile_column), str(stats), str(status), str(mod))

                doc_num = queries.get_report_number(doc)

                queries.update_progress_report(doc_num, eight_tuple)
        
        '''

        
        claim_titles = queries.get_claim_titles()
        #claim_titles = ['GA-2022-402 PT-077-C4']
        for title in claim_titles: 
                claim_year = functions.parse_doc_title(title)['year']
                print("document title is: ", title)
                claim = claimsfunctions.get_claims(s, r, title)
                if len(claim) > 0:
                        claim = claim[0] 
                elif int(claim_year) < 2022: 
                        three_tuple = ('0.0', '9', str(datetime(1,1,1)))
                        doc_num = queries.get_claim_number(title)
                        queries.update_claim(doc_num, three_tuple)
                        print("No report: ", claim)
                        continue
                else: 
                        print("No report found: ", claim)
                        continue

                status = claimsfunctions.get_claim_status(s, claim)

                expense = claimsfunctions.get_current_expense(s, claim)

                claim_mod = claimsfunctions.last_modification(s, claim)

                three_tuple = (str(expense), str(status), str(claim_mod))

                claim_num = queries.get_claim_number(title)

                queries.update_claim(claim_num, three_tuple)






                '''
                session = s
                link = functions.open_link(session, claim)
                time.sleep(13)
                exp = BeautifulSoup(link.text, 'html.parser').find('a', text = 'Expense Summary')
                #print("link is: ", link)
                postnext = urljoin(url, exp['href'])
                time.sleep(4)
                expense = session.post(postnext)
                expensetable = BeautifulSoup(expense.text).find_all(doc_elements.expense_table['tag'], id=re.compile(doc_elements.expense_table['id']))
                expenseframe = pandas.read_html(str(expensetable[7]))[0]
                current_expense = expenseframe.iloc[13,5]
                if type(current_expense) == float:
                        current_expense = '$0.0'
                        print("current expense is: ", current_expense)
                print("current expense is: ", current_expense)

                pdb.set_trace()
                '''

        print(doclist)
