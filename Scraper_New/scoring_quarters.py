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
    tuples = cursor.execute('''SELECT Grants.GrantID, Grants.Budget
        FROM Grants
        Where Year = ?
        ''', year).fetchall()  

    return tuples

def get_grant_reports(grant_id):
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

def get_grant_claims(grant_id):
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
        ''', year, '5').fetchall()  
    score_list = []
    for score in scores:
        score_list.append(list(score))

    score_fields = ['Organization', 'DocTitle', 'Program', 'GoalsScore', 'MilestoneScore' , 'SpendingScore', 'QualityScore', 'TimelinessScore', 'OverallWeighted', 'Grade']
    score_frame = pandas.DataFrame(score_list, columns = score_fields)
    
    return score_frame

grant_id = get_grants(2021)[5]

def report_frame(grant):
    grant_id = grant[0]
    reports = get_grant_reports(grant_id)
    report_list = []
    for report in reports: 
        report_list.append(list(report))

    report_fields = ['ReportNumber','GrantID','DocNumber','Accomplishments','Challenges','PIandE','FutureEvents','TargetMilestones','AchievedMilestones','StatisticalSummary','Status','LastModified']
    report_frame = pandas.DataFrame(report_list, columns = report_fields)

    return report_frame

def claim_frame(grant_id):
    claims = get_grant_claims(grant_id[0])
    claim_list = []
    for claim in claims: 
        claim_list.append(list(claim))

    claim_fields = ['ClaimID','GrantID','DocNumber','CurrentExpense','Status','LastModified']
    claim_frame = pandas.DataFrame(claim_list, columns = claim_fields)

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

def percent_achieved(grant_id):
    achieved = report_frame(grant_id)['AchievedMilestones']
    targets = report_frame(grant_id)['TargetMilestones']
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

def milestone(grant_id):
    percent = percent_achieved(grant_id) 
    mile_list = []

    for i in percent: 
        mile_list.append(float(i))
    try:
        result = sum(mile_list)/len(mile_list)
    except:
        result = 0
    return result

def goal(grant_id):
    percent = percent_achieved(grant_id)
    goal_list = []
    
    for i in percent:
        if i < 1:
            goal_list.append(0)
        else:
            goal_list.append(1)

    return sum(goal_list)/len(goal_list)

def narrative(grant_id):
    accomplishments = report_frame(grant_id)['Accomplishments']
    score_list = []
    for accomplishment in accomplishments:
        try:
            score = min(1, len(accomplishment)/100)
        except:
            score = 0
        score_list.append(score)

    return sum(score_list)/3

def due_date(docnumber):
    year = 2022

    start_date = datetime(year-1, 10, 20)
    number = int(docnumber.replace(" ", "").replace("R", ""))
    due = start_date + relativedelta(months = number)

    return due

def overdue(frame_row):
    try:
        overdue = frame_row['LastModified'] - due_date(frame_row['DocNumber'])  

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

def timeliness(grant_id):
    rep_frame = report_frame(grant_id)
    rep_frame['Overdue'] = rep_frame.apply(overdue, axis=1)

    return sum(rep_frame['Overdue'])/3

def spent(grant_id):
    grant_budget = grant_id[1]
    quarter_budget = float(grant_budget)*0.25
    expense_list = claim_frame(grant_id)['CurrentExpense']
    total_spend = float(expense_list.sum())
    if grant_budget > 0: 
        spent_score = min(1, total_spend/quarter_budget)
    else: 
        spent_score = 1.0

    return spent_score

def weight_scores(grant_id):
    goals_score = goal(grant_id)
    milestones_score = milestone(grant_id)
    narrative_score = narrative(grant_id)
    timeliness_score = timeliness(grant_id)
    spent_score = spent(grant_id)

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

def record_scores(grant_id):
    scores_tuple = (grant_id[0], 0, milestone(grant_id), spent(grant_id), narrative(grant_id), timeliness(grant_id), weight_scores(grant_id), 0, 6)
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=' + str(conn_filename) + ';')
    cursor = conn.cursor()

    tuples = cursor.execute('''INSERT INTO GrantScores (GrantID, GoalsScore, MilestonesScore, SpendingScore, QualityScore, TimelinessScore, OverallWeighted, Grade, ScoreType)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''', scores_tuple).commit()
    
goals_score = goal(grant_id)
milestones_score = milestone(grant_id)
narrative_score = narrative(grant_id)
timeliness_score = timeliness(grant_id)
spent_score = spent(grant_id)

def score_grants(year):
    for grant in get_grants(year):
       record_scores(grant)


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

#pdb.set_trace()