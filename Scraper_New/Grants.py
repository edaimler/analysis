
from openpyxl import load_workbook, Workbook
from pathlib import Path
import pdb
import pyodbc


def parse_doc_title(doc_title):
    if doc_title == None:
        return None
    parsed = doc_title.split("-")
    if len(parsed) > 3:
        document = {
        "doctitle": doc_title,
        #"organization": 
        "program": parsed[0],
        "year": int(parsed[1]),
        }

        if len(parsed) > 4:
            document["classification"] = parsed[2] + '-' + parsed[3]
            document["id"] = parsed[4]
        else:
            document["classification"] = parsed[2]
            document["id"] = parsed[3]
    else:
        document = None
    return document


def populate_grants():
    path = Path('F:\TSREG')
    filename = 'All Grants.xlsx'
    database = 'GOHS Grants.accdb'
    fullname = path / filename
    workbook = load_workbook(filename = fullname)
    sheet = workbook['Sheet1']
    grants_count = len(sheet['B'])
    orgs_grants = sheet['B1:C'+str(grants_count)]
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    for row in orgs_grants:
        document = parse_doc_title(row[1].value)
        if document is not None:
            document["organization"] = str(row[0].value)
            organization = document["organization"]
            organization_apost = organization.replace("'", "''")
            existing_org = cursor.execute("SELECT Organization, GranteeID FROM Grantees WHERE Organization = " + "'" + organization_apost + "'").fetchall()
            existing_grant = cursor.execute("SELECT DocTitle FROM Grants WHERE DocTitle = " + "'" + document["doctitle"] + "'").fetchall() # TODO: ADD THE GRANTEE ORG TO THIS FETCH

            if len(existing_org) == 0:
                cursor.execute(''' 
                    INSERT INTO Grantees 
                    (Organization) 
                    VALUES (?) 
                    ''',
                    (organization) 
                )
                existing_org = cursor.execute("SELECT Organization, GranteeID FROM Grantees WHERE Organization = " + "'" + organization_apost + "'").fetchall()
            if len(existing_grant) == 0:
                cursor.execute(''' 
                    INSERT INTO Grants 
                    (DocTitle, GranteeID, Program, Classification, Year, GrantNumber) 
                    VALUES (?,?,?,?,?,?) 
                    ''',
                    (document["doctitle"], existing_org[0][1] , document["program"], document["classification"], document["year"], document['id']) #TODO: add the grantee ID
                )
    conn.commit()
    

def populate_progress_reports():
    path = Path('F:\TSREG')
    database = 'GOHS Grants.accdb'
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    grants = cursor.execute("SELECT GrantID FROM Grants WHERE NOT EXISTS (SELECT * FROM ProgressReports WHERE GrantID = Grants.GrantID)").fetchall()
    
    for grant in grants: 
        for i in range(12):
            cursor.execute('''INSERT INTO ProgressReports 
                ( GrantID, DocNumber) 
                VALUES (?, ?)''',
                (grant[0], "R"+str(i+1))
            )
    conn.commit()
    print("populated %s progress report records"%(len(grants)*12))

def populate_claims():
    path = Path('F:\TSREG')
    database = 'GOHS Grants.accdb'
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    grants = cursor.execute("SELECT GrantID FROM Grants WHERE NOT EXISTS (SELECT * FROM Claims WHERE GrantID = Grants.GrantID)").fetchall()
    
    for grant in grants: 
        for i in range(12):
            cursor.execute('''INSERT INTO Claims 
                ( GrantID, DocNumber) 
                VALUES (?, ?)''',
                (grant[0], "C"+str(i+1))
            )
    conn.commit()
    print("populated %s claim records"%(len(grants)*12))

def populate_final_reports():
    path = Path('F:\TSREG')
    database = 'GOHS Grants.accdb'
    conn_filename = path / database
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    grants = cursor.execute("SELECT GrantID FROM Grants WHERE NOT EXISTS (SELECT * FROM FinalReports WHERE GrantID = Grants.GrantID)").fetchall()
    
    for grant in grants: 
        cursor.execute('''INSERT INTO FinalReports 
            ( GrantID, DocNumber) 
            VALUES (?, ?)''',
            (grant[0], "FR"+str(1))
        )
            
    conn.commit()
    print("populated %s final report records"%len(grants))


