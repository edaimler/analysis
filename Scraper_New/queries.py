from datetime import datetime
import pandas
from pathlib import Path
import pyodbc
import pdb


database = 'GOHS Grants.accdb'
#database = 'TEST.accdb'

def report_numbers():
    current_month = datetime.now().month
    late_report_number = current_month + 1 % 12 
    number_list = []
    for n in range(1, 1+ late_report_number):
        number_list.append(n%12)

    number_string = "("
    for i in number_list:
        number_string = number_string + ''' DocNumber = 'R'''+str(i)+ '''' OR'''
    number_string = number_string.rstrip('OR')
    number_string = number_string + '))'
        
    return number_string

def get_grants_and_report_numbers():
    path = Path('F:\TSREG')
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.DocTitle, ProgressReports.DocNumber
        FROM ProgressReports 
        INNER JOIN Grants ON ProgressReports.GrantID = Grants.GrantID
        WHERE (Status NOT BETWEEN 4 AND 5
        AND ''' + report_numbers()).fetchall()  

    return tuples


def get_prog_report_titles():
    reports = get_grants_and_report_numbers()

    docs = []

    for report in reports: 
        doc_title = str(report[0]) + "-" + str(report[1]) 
        docs = docs + [doc_title]
    
    return docs

def claim_numbers():
    current_month = datetime.now().month
    late_report_number = current_month + 1 % 12 
    number_list = []
    for n in range(1, 1+ late_report_number):
        number_list.append(n%12)

    number_string = "("
    for i in number_list:
        number_string = number_string + ''' DocNumber = 'C'''+str(i)+ '''' OR'''
    number_string = number_string.rstrip('OR')
    number_string = number_string + '))'
        
    return number_string

def get_grants_and_claim_numbers():
    path = Path('F:\TSREG')
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.DocTitle, Claims.DocNumber
        FROM Claims 
        INNER JOIN Grants ON Claims.GrantID = Grants.GrantID
        WHERE ((Status < 6)
        AND Grants.Year = 2022
        AND ''' + claim_numbers()).fetchall()  

    return tuples


def get_claim_titles():
    claims = get_grants_and_claim_numbers()

    docs = []

    for claim in claims: 
        doc_title = str(claim[0]) + "-" + str(claim[1]) 
        docs = docs + [doc_title]
    
    return docs



def get_report_number(report_title):
        split = report_title.replace("'", "").split('-')
        doc_number = "'" + split[-1] + "'"
        remove = '-'+split[-1]
        doc_title = "'" + report_title.replace(remove, '').replace("'", "") + "'"

        path = Path('F:\TSREG')
        conn_filename = path / database
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
        cursor = conn.cursor()

        report_id = cursor.execute(''' 
        SELECT ProgressReports.*, Grants.GrantID FROM ProgressReports, Grants
        WHERE ProgressReports.GrantID = Grants.GrantID 
        AND Grants.DocTitle = ''' + doc_title + ''' 
        AND ProgressReports.DocNumber = ''' + doc_number ).fetchall()
        print("report_id is: ", report_id)
        report_number = str(report_id[0][0])

        return report_number


def update_progress_report(doc_number, eight_tuple):
        path = Path('F:\TSREG')
        conn_filename = path / database
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
        cursor = conn.cursor()

        tuples = cursor.execute('''
        UPDATE ProgressReports
        SET Accomplishments=? 
        , Challenges=?
        , PIandE=?
        , TargetMilestones=?
        , AchievedMilestones=?
        , StatisticalSummary=?
        , Status=?
        , LastModified=?
        WHERE ReportNumber='''+doc_number
        , eight_tuple)

        conn.commit()


def get_claim_number(claim_title):
        split = claim_title.replace("'", "").split('-')
        doc_number = "'" + split[-1] + "'"
        remove = '-'+split[-1]
        doc_title = "'" + claim_title.replace(remove, '').replace("'", "") + "'"

        path = Path('F:\TSREG')
        conn_filename = path / database
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
        cursor = conn.cursor()

        report_id = cursor.execute(''' 
        SELECT Claims.*, Grants.GrantID FROM Claims, Grants
        WHERE Claims.GrantID = Grants.GrantID 
        AND Grants.DocTitle = ''' + doc_title + ''' 
        AND Claims.DocNumber = ''' + doc_number ).fetchall()

        report_number = str(report_id[0][0])

        return report_number


def update_claim(doc_number, three_tuple):
        path = Path('F:\TSREG')
        conn_filename = path / database
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
        cursor = conn.cursor()

        tuples = cursor.execute('''
        UPDATE Claims
        SET CurrentExpense=? 
        , Status=?
        , LastModified=?
        WHERE ClaimID='''+doc_number
        , three_tuple)

        conn.commit()
