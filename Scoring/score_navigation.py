from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from pathlib import Path
import score_functions
import pandas



def open_file(org_double, docnumber, frame_name):
    doc_dict = score_functions.parse_doc_title(org_double)

    path = Path('Grantee Data')
    year = path / doc_dict['year'] 
    prog = year / 'Progress Reports'
    claim = year / 'Claims'
    final = year / 'Final Reports'
    
    if docnumber[:1] == "R":
        org = prog / doc_dict['organization']
    if docnumber[:1] == "C":
        org = claim / doc_dict['organization']
    if docnumber[:1] == "F":
        org = final / doc_dict['organization'] 
    
    if not org.is_dir():
        print("error: {} not found", doc_dict['organization'])
        return

    filename = doc_dict['grantname']+'-'+docnumber+'.xlsx'
    fullname = org / filename.strip()
    if not fullname.exists():
        print("error 2: {} not found", fullname)
        df = pandas.DataFrame([[0]*15]*15)
    else: 
        df = pandas.read_excel(fullname, sheet_name = frame_name)
    
    return df


def goals(org_double, month):
    report_number = "R"+str(month)
    achieved = open_file(org_double, report_number, "achieved_milestones")
    objectives = open_file(org_double, report_number, "milestone")

    #achieved.pop(0)
    objectives.pop(0)
    #objectives.pop(1)
    objectives.pop(13)
    obj2 = objectives.drop([0], axis=0)
    obj2

    goals = []
    for i in range(obj2.shape[0]):
        obj3 = obj2.iloc[i].astype(float)
        obj3 = obj3.iloc[1:]
        print("obj3 is: ", obj3[:9])
        goals.append(obj3[:9].sum())

    attempts = []
    for i in range(achieved.shape[0]):
        remove_na = achieved.iloc[i].fillna(0)
        remove_na = remove_na.iloc[1:]
        print("remove_na is: ", remove_na[:9])
        ach = remove_na[:9].astype(float)
        attempts.append(ach.sum())
        print("ach sum is: ", ach.sum())
    
    goals_score = []
    for i in range(len(goals)):
        attempts = attempts + [0,0,0,0,0,0,0,0,0,0]
        if attempts[i] >= goals[i]:
            goals_score.append(1)
        else:
            goals_score.append(0)
    print("goals_score is: ", goals_score)
    print("goals is: ", goals)
    print("attempts is: ", attempts)
    goals_score = sum(goals_score)/len(goals_score)
    
    return goals_score


def milestones(org_double, month):
    report_number = "R"+str(month)
    achieved = open_file(org_double, report_number, "achieved_milestones")
    objectives = open_file(org_double, report_number, "milestone")

    #achieved.pop(0)
    objectives.pop(0)
    #objectives.pop(1)
    objectives.pop(11)
    obj2 = objectives.drop([0], axis=0)
    obj2

    goals = []
    for i in range(obj2.shape[0]):
        obj3 = obj2.iloc[i].astype(float)
        obj3 = obj3.iloc[1:]
        goals.append(obj3[:9].sum())

    attempts = []
    for i in range(achieved.shape[0]):
        remove_na = achieved.iloc[i].fillna(0)
        remove_na = remove_na.iloc[1:]
        ach = remove_na.astype(float)
        attempts.append(ach[:9].sum())
    
    milestones_score = []
    for i in range(len(goals)):
        attempts = attempts + [0,0,0,0,0,0,0,0,0,0]
        if attempts[i] >= goals[i]:
            milestones_score.append(1)
        else:
            milestones_score.append(attempts[i]/goals[i])
    milestones_score = sum(milestones_score)/len(milestones_score)
    
    return milestones_score

def spending(org_double, month):
    report_number = "C"+str(month)
    spend = open_file(org_double, report_number, "expense")
    if spend.iloc[12,2] != 0:
        budget_total = float(spend.iloc[12,2][1:].replace(",",""))
        budget_pending = float(spend.iloc[12,8][1:].replace(",",""))
    else: 
        budget_total = 1
        budget_pending = 0
    budget_spent = budget_total - budget_pending
    spend_score = budget_spent/budget_total
    print(spend_score)

    return spend_score

def quality(org_double, month):
    narrative_list = []
    for i in range(month):
        narrative = open_file(org_double, "R"+str(i+1), "narrative")
        if narrative.iloc[1,1] != 0:
            character_count = len(narrative.iloc[1,1])
        else:
            character_count = 0
        if character_count < 120:
            narrative_list.append(0.50)
        elif 120 <= character_count < 150:
            narrative_list.append(0.75)
        else:
            narrative_list.append(1.0)
    score = sum(narrative_list)/9
    
    return score

def timeliness(org_double, month):
    duedate = datetime(2020, 10, 20)
    scorelist = []
    for i in range(month):
        name = open_file(org_double, "R"+str(i+1), "name")
        datestring = name.iloc[5,6]
        if type(datestring) is str:
            date_list = datestring.split(' ')
            date = None
            for i in date_list:
                if i[0].isdigit():
                    date = datetime.strptime(i, '%m/%d/%Y')
                    break
            duedate = duedate + relativedelta(months=1)
            print("duedate is: ", duedate)
            print("date is: ", date)
            difference = duedate - date
            print("difference is: ", difference)
            if difference.days >= 0:
                scorelist.append(1.0)
            elif -90 <= difference.days < 0:
                scorelist.append(0.5)
            elif -90 > difference.days:
                scorelist.append(0.25)
            else: scorelist.append(0.0)
            print("scorelist is: ", scorelist)
        else: 
            scorelist.append(0.0)
            print("scorelist is: ", scorelist)
    score = sum(scorelist)/9
    print("timeliness is: ", score)
    return score

def weight_score(org_double, month):
    goalsw = goals(org_double, month)
    milestonesw = milestones(org_double, month)
    spendingw = spending(org_double, month)
    qualityw = quality(org_double, month)
    timelinessw = timeliness(org_double, month)
    print("goals, milestones, spending, quality, timeliness: ", [goalsw, milestonesw, spendingw, qualityw, timelinessw])

    weighted = (goalsw*.6+milestonesw*.4)*.5 + (spendingw*.25) + (qualityw*.6 + timelinessw*.4)*.25
    print("weighted is: ", weighted)
    return weighted


def grade_score(weighted_score):
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

def save_file(doc_dict, data_frame, frame_name):
    path = Path('F:\TSREG\Grantee Data')
    year = path / doc_dict['year'] 
    prog = year / 'Progress Reports'
    claim = year / 'Claims'
    final = year / 'Final Reports'
    
    if doc_dict['docnumber'][:1] == "R":
        org = prog / doc_dict['organization']
    if doc_dict['docnumber'][:1] == "C":
        org = claim / doc_dict['organization']
    if doc_dict['docnumber'][:1] == "F":
        org = final / doc_dict['organization'] 
    
    if not org.is_dir():
        org.mkdir(parents=True)

    filename = doc_dict['docname']+'.xlsx'
    fullname = org / filename

    if not fullname.exists():
        data_frame.to_excel(fullname, sheet_name = frame_name)
    else: 
        with pandas.ExcelWriter(fullname, mode = 'a', if_sheet_exists = 'replace') as writer:
            data_frame.to_excel(writer, sheet_name = frame_name)

doclist = [("Insert Grantee Here")]
