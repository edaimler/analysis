from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas
from pathlib import Path
import pyodbc
import pdb

path = Path('F:\TSREG')
database = 'GOHS Grants.accdb'
conn_filename = path / database

def get_grants(year):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.GrantID, Grants.GranteeID, Grants.Program, Grants.Classification, Grants.Budget
        FROM Grants
        Where Year = ?
        ''', year).fetchall()  

    return tuples

def get_grants2(grantee):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT ta1.GrantID, ta2.GrantID, ta1.GranteeID, ta1.Program, ta1.Classification, ta1.Budget, ta2.Budget
        FROM (Grants ta1
        INNER JOIN Grants ta2
        ON ta1.GrantID <> ta2.GrantID)
        WHERE ta1.Program = 'GA4'
        AND ta1.Program = ta2.Program
        AND ta1.Year = 2021 
        AND ta2.Year = 2022 
        AND ta1.GranteeID = ?
        AND ta2.GranteeID = ?
        ''', grantee, grantee).fetchall()  

    return tuples

def get_grant_reports_year_one(grant_id):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT *
        FROM ProgressReports
        WHERE (ProgressReports.GrantID = ?
        AND (ProgressReports.DocNumber = ? 
        OR ProgressReports.DocNumber = ? 
        OR ProgressReports.DocNumber = ? 
        ))
        ''', grant_id, 'R10', 'R11', 'R12').fetchall()  

    return tuples

def get_grant_reports_year_two(grant_id):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT *
        FROM ProgressReports
        WHERE (ProgressReports.GrantID = ?
        AND (ProgressReports.DocNumber = ? 
        OR ProgressReports.DocNumber = ? 
        OR ProgressReports.DocNumber = ? 
        ))
        ''', grant_id, 'R1', 'R2', 'R3').fetchall()  

    return tuples

def get_grant_claims_year_one(grant_id):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT *
        FROM Claims
        WHERE (Claims.GrantID = ?
        AND (Claims.DocNumber = ?
        OR Claims.DocNumber = ?\
        OR Claims.DocNumber = ?
        ))
        ''', grant_id, 'C10', 'C11', 'C12').fetchall()  

    return tuples

def get_grant_claims_year_two(grant_id):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT *
        FROM Claims
        WHERE (Claims.GrantID = ?
        AND (Claims.DocNumber = ?
        OR Claims.DocNumber = ?\
        OR Claims.DocNumber = ?
        ))
        ''', grant_id, 'C1', 'C2', 'C3').fetchall()  

    return tuples

def get_scores(year):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    scores = cursor.execute('''SELECT Grantees.Organization, Grants.DocTitle, Grants.Program, GrantScores.GoalsScore, GrantScores.MilestonesScore, GrantScores.SpendingScore, GrantScores.QualityScore, GrantScores.TimelinessScore, GrantScores.OverallWeighted, GrantScores.Grade    
        FROM (Grants 
        INNER JOIN Grantees ON Grants.GranteeID = Grantees.GranteeID)
        INNER JOIN GrantScores ON Grants.GrantID = GrantScores.GrantID
        Where (Year = ?
        AND ScoreType = ?)
        ''', year, '2').fetchall()  
    score_list = []
    for score in scores:
        score_list.append(list(score))

    score_fields = ['Organization', 'DocTitle', 'Program', 'GoalsScore', 'MilestoneScore' , 'SpendingScore', 'QualityScore', 'TimelinessScore', 'OverallWeighted', 'Grade']
    score_frame = pandas.DataFrame(score_list, columns = score_fields)
    
    return score_frame

grant_id = get_grants(2021)[5]

def report_frame(grant, year):
    #grant_id = grant[0]
    grant_id = grant
    if year == 1:
        reports = get_grant_reports_year_one(grant_id)
    else: 
        reports = get_grant_reports_year_two(grant_id)
    report_list = []
    for report in reports: 
        report_list.append(list(report))

    report_fields = ['ReportNumber','GrantID','DocNumber','Accomplishments','Challenges','PIandE','FutureEvents','TargetMilestones','AchievedMilestones','StatisticalSummary','Status','LastModified']
    report_frame = pandas.DataFrame(report_list, columns = report_fields)

    return report_frame

def claim_frame(grant_id, year):
    if year == 1:
        claims = get_grant_claims_year_one(grant_id)
    else: 
        claims = get_grant_claims_year_two(grant_id)
    claim_list = []
    for claim in claims: 
        claim_list.append(list(claim))

    claim_fields = ['ClaimID','GrantID','DocNumber','CurrentExpense','Status','LastModified']
    claim_frame = pandas.DataFrame(claim_list, columns = claim_fields)
    print("year is ", year)
    print("claim frame is ", claim_frame)

    return claim_frame

def str_to_int(str_list):
    result_list = []
    result = str_list[1:-1].replace(" ", "").split(',')
    for i in result:
        try:
            result_list.append(int(i[1:-1]))
        except:
            result_list.append(0)

    return result_list

def miles_percent(num_achieved, num_target):
    try:
        individual_score = min(float(num_achieved)/float(num_target), 1)
    except: 
        individual_score = 1

    return individual_score

def percent_achieved(grant_id, year):
    achieved = report_frame(grant_id, year)['AchievedMilestones']
    targets = report_frame(grant_id, year)['TargetMilestones']
    try:
        achieved_int = achieved.map(str_to_int)
        targets_int = targets.map(str_to_int)
        df_ach = pandas.DataFrame(achieved_int.to_list())
        df_tar = pandas.DataFrame(targets_int.to_list())
        ach_totals = []
        tar_totals = []
        for column in df_ach.iteritems():
            ach_totals.append(column[1].sum())

        for column in df_tar.iteritems():
            tar_totals.append(column[1].sum())

        percent = list(map(miles_percent, ach_totals, tar_totals))

        return percent

    except:
        return [0]

def milestone(grant_id, year):
    percent = percent_achieved(grant_id, year) 
    mile_list = []

    for i in percent: 
        mile_list.append(float(i))
    try:
        result = sum(mile_list)/len(mile_list)
    except:
        result = 0
    return result

def goal(grant_id, year):
    percent = percent_achieved(grant_id, year)
    goal_list = []
    
    for i in percent:
        if i < 1:
            goal_list.append(0)
        else:
            goal_list.append(1)

    return sum(goal_list)/len(goal_list)

def narrative(grant_id, year):
    accomplishments = report_frame(grant_id, year)['Accomplishments']
    score_list = []
    for accomplishment in accomplishments:
        try:
            score = min(1, len(accomplishment)/100)
        except:
            score = 0
        score_list.append(score)

    return sum(score_list)/3

def due_date_year_one(docnumber):
    year = 2021

    start_date = datetime(year-1, 10, 20)
    number = int(docnumber.replace(" ", "").replace("R", ""))
    due = start_date + relativedelta(months = number)

    return due

def due_date_year_two(docnumber):
    year = 2022

    start_date = datetime(year-1, 10, 21)
    number = int(docnumber.replace(" ", "").replace("R", ""))
    due = start_date + relativedelta(months = number)

    return due

def overdue_year_one(frame_row):
    try:
        overdue = frame_row['LastModified'] - due_date_year_one(frame_row['DocNumber'])  

        if overdue <= timedelta(days=0):
            time_score = 1
        elif timedelta(days=0) < overdue < timedelta(days = 90):
            time_score = 0.5 
        elif timedelta(days=90) < overdue:
            time_score = 0.25 
        else: 
            time_score = 0.0
    except:
        time_score = 0.0

    return time_score

def overdue_year_two(frame_row):
    try:
        overdue = frame_row['LastModified'] - due_date_year_two(frame_row['DocNumber'])  

        if overdue <= timedelta(days=0):
            time_score = 1
        elif timedelta(days=0) < overdue < timedelta(days = 90):
            time_score = 0.5 
        elif timedelta(days=90) < overdue:
            time_score = 0.25 
        else: 
            time_score = 0.0
    except:
        time_score = 0.0

    return time_score

def timeliness(grant_id, year):
    rep_frame = report_frame(grant_id, year)
    if year == 1:
        rep_frame['Overdue'] = rep_frame.apply(overdue_year_one, axis=1)
    else: 
        rep_frame['Overdue'] = rep_frame.apply(overdue_year_two, axis=1)
    return sum(rep_frame['Overdue'])/3

def spent(grant, year):
    if year == 1:
        grant_budget = grant[5]
        grant_id = grant[0]
    else: 
        grant_budget = grant[6]
        grant_id = grant[1]

    quarter_budget = float(grant_budget)*0.25

    expense_list = claim_frame(grant_id, year)['CurrentExpense']
    total_spend = float(expense_list.sum())
    if grant_budget > 0: 
        spent_score = min(1, total_spend/quarter_budget)
    else: 
        spent_score = 1.0

    return spent_score

def weight_scores(unique_grant):
    grant_id_year_one = unique_grant[0]
    grant_id_year_two = unique_grant[1]

    milestones_score = milestone(grant_id_year_two, 2)
    narrative_score = narrative(grant_id_year_two, 2)
    timeliness_score = timeliness(grant_id_year_two, 2)
    spent_score = spent(unique_grant, 2)

    weighted_score = (milestones_score)*0.5 + 0.25*spent_score + (0.6*narrative_score + 0.4*timeliness_score)*0.25

    return weighted_score

def grade(grant_id):
    weighted_score = weight_scores(grant_id)
    grade = ""
    if weighted_score > .954:
        grade = "A+"
    elif weighted_score >= .895:
        grade = "A"
    elif weighted_score >= .795:
        grade = "B"
    elif weighted_score >= .695:
        grade = "C"
    else:
        grade = "D"
    
    return grade

def record_scores(unique_grant):
    grant_id_year_one = unique_grant[0]
    grant_id_year_two = unique_grant[1] 

    milestones_score = milestone(grant_id_year_two, 2)
    narrative_score = narrative(grant_id_year_two, 2)
    timeliness_score = timeliness(grant_id_year_two, 2)
    spent_score = spent(unique_grant, 2)
    weighted = weight_scores(unique_grant)

    scores_tuple = (grant_id_year_two, 0, milestones_score, spent_score, narrative_score, timeliness_score, weighted, 0, 4)
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()

    tuples = cursor.execute('''INSERT INTO GrantScores (GrantID, GoalsScore, MilestonesScore, SpendingScore, QualityScore, TimelinessScore, OverallWeighted, Grade, ScoreType)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''', scores_tuple).commit()
    


def score_grants(year):
    for grant in get_grants(year):
       print("grant is ", grant)
       for unique_grant in get_grants2(grant[1]):
           print("unique grant is ", unique_grant)
           record_scores(unique_grant)
           try:
                print("unique grant is ", unique_grant)
                record_scores(unique_grant)
           except:
                print('error scoring grant')
                continue 

test = get_grants(2022)[10][1]
test2 = get_grants2(test)

def all_reports(year):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.DocTitle, Grants.Year, ProgressReports.*
        FROM (ProgressReports
        INNER JOIN Grants ON ProgressReports.GrantID = Grants.GrantID)
        Where Grants.Year = ?
        ''', year).fetchall()  

    return tuples

def search_challenges(search_string, year):
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()
    tuples = cursor.execute('''SELECT Grants.DocTitle, Grants.Year, ProgressReports.*
        FROM (ProgressReports
        INNER JOIN Grants ON ProgressReports.GrantID = Grants.GrantID)
        Where ProgressReports.Challenges LIKE ?
        AND Grants.Year = ?
        ''', search_string, year).fetchall()  
    return tuples

all = all_reports(2021)
covid = search_challenges('%COVID%', 2021)
staff = search_challenges(r'%staff%', 2021)
shortage = search_challenges(r'%shortage%', 2021)

test_grant = '1390' #2022
'''
unique_grant = get_grants2(424)
grant_id_year_one = get_grants2(424)[0][0]
grant_id_year_two = get_grants2(424)[0][1]

claimtest = claim_frame(grant_id_year_one, 1)
claimtest2 = claim_frame(grant_id_year_two, 2)


milestones_score = milestone(grant_id_year_one, 1)
narrative_score = narrative(grant_id_year_one, 1)
timeliness_score = timeliness(grant_id_year_one, 1)
spent_score = spent(get_grants2(424)[0], 1)

get_grants2(543)
'''
#score_grants(2022)
#pdb.set_trace()
