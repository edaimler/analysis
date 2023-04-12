

from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from doctest import DONT_ACCEPT_TRUE_FOR_1
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas
from pathlib import Path
import pyodbc
import pdb
import scoring_quarters


scores = scoring_quarters.get_scores(2021)
path = Path('F:\TSREG')
excel = 'latex_data.xlsx'
filename = path / excel

def empty():
    return None 

#with open('F:/TSREG/2022 Grant/Lifesavers Project/Data/charts.txt', 'w') as output_file:  

def program_data(scores):
    ga = scores[scores['Program']=='GA']
    ya = scores[scores['Program']=='YA']
    ten = scores[scores['Program']=='TEN']
    sadd = scores[scores['Program']=='SADD']
    overall = scores 

    program = ['GA', 'YA', 'TEN', 'SADD', 'Overall'] 
    count = [len(ga), len(ya), len(ten), len(sadd), len(overall)]

    average_goals = [ga['GoalsScore'].sum()/len(ga), ya['GoalsScore'].sum()/len(ya), ten['GoalsScore'].sum()/len(ten), sadd['GoalsScore'].sum()/len(sadd), overall['GoalsScore'].sum()/len(overall)]
    average_milestones = [ga['MilestoneScore'].sum()/len(ga), ya['MilestoneScore'].sum()/len(ya), ten['MilestoneScore'].sum()/len(ten), sadd['MilestoneScore'].sum()/len(sadd), overall['MilestoneScore'].sum()/len(overall)]
    average_spending = [ga['SpendingScore'].sum()/len(ga), ya['SpendingScore'].sum()/len(ya), ten['SpendingScore'].sum()/len(ten), sadd['SpendingScore'].sum()/len(sadd), overall['SpendingScore'].sum()/len(overall)] 
    average_quality = [ga['QualityScore'].sum()/len(ga), ya['QualityScore'].sum()/len(ya), ten['QualityScore'].sum()/len(ten), sadd['QualityScore'].sum()/len(sadd), overall['QualityScore'].sum()/len(overall)]
    average_timeliness = [ga['TimelinessScore'].sum()/len(ga), ya['TimelinessScore'].sum()/len(ya), ten['TimelinessScore'].sum()/len(ten), sadd['TimelinessScore'].sum()/len(sadd), overall['TimelinessScore'].sum()/len(overall)]
    average_weighted = [ga['OverallWeighted'].sum()/len(ga), ya['OverallWeighted'].sum()/len(ya), ten['OverallWeighted'].sum()/len(ten), sadd['OverallWeighted'].sum()/len(sadd), overall['OverallWeighted'].sum()/len(overall)]

    programs = pandas.DataFrame([program, count, average_goals, average_milestones, average_spending, average_quality, average_timeliness, average_weighted])
    programs = programs.transpose() 
    programs.columns = ['Program', 'ProgramCount', 'AverageGoals', 'AverageMilestones', 'AverageSpending', 'AverageQuality', 'AverageTimeliness', 'AverageWeighted']

    return programs

def count_grades(scores, program = 'Overall'):
    if program == 'Overall':
        a_plus = len(scores[scores['Grade'] == 'A+'])
        a_grade = len(scores[scores['Grade'] == 'A'])
        b_grade = len(scores[scores['Grade'] == 'B'])
        c_grade = len(scores[scores['Grade'] == 'C'])
        d_grade = len(scores[scores['Grade'] == 'D'])
        all_grades = len(scores)
    else:
        a_plus = len(scores[(scores['Grade'] == 'A+') & (scores['Program']==program)])
        a_grade = len(scores[(scores['Grade'] == 'A') & (scores['Program']==program)])
        b_grade = len(scores[(scores['Grade'] == 'B') & (scores['Program']==program)])
        c_grade = len(scores[(scores['Grade'] == 'C') & (scores['Program']==program)])
        d_grade = len(scores[(scores['Grade'] == 'D') & (scores['Program']==program)])
        all_grades = len(scores[scores['Program']==program])

    grades = ['A+', 'A', 'B', 'C', 'D']
    count = [a_plus, a_grade, b_grade, c_grade, d_grade]
    percent = [0, 0, 0, 0, 0] 

    grade_data = pandas.DataFrame([grades, count, percent])
    grade_data = grade_data.transpose()
    grade_data.columns = ['Grade', 'Count', 'Percent']

    return grade_data


def grade_chart(scores, program = 'Overall'):
    grades = count_grades(scores, program)

    chart_string = r'''\begin{figure}[h]
	\begin{center}
	\begin{tikzpicture}
	\begin{axis}[ybar,
		%ymin=0,
		enlarge x limits=.25,
		x=-1.4cm,
		bar width=1cm,
		ymin=0,
		ymax=50,
		%legend style={at={(1.6, 0.6)}}, %Positions the Chart's Legend
		ylabel={Number of Grantees},
		symbolic x coords={D, C, B, A, A+},
		xtick=data,
		x tick label style={rotate=0, anchor=north},
		%nodes near coords, 
		nodes near coords align={horizontal},]
	
		\addplot+ [ybar, draw=herrick, fill=herrick] coordinates{
            (A+,''' + grades[grades['Grade']=='A+']['Count']+''') 
            (A,''' + grades[grades['Grade']=='A']['Count']+''') 
            (B,''' + grades[grades['Grade']=='B']['Count']+''') 
            (C,''' + grades[grades['Grade']=='C']['Count']+''') 
            (D,''' + grades[grades['Grade']=='D']['Count']+''')};
	%\legend{Distraction-Prone, Distraction-Averse};
	
	\end{axis}
	\end{tikzpicture}
	\caption{'''+program+'''Grade Distribution}
	\label{table:allgrades}
	\end{center}
	\end{figure}
    '''

    return chart_string
     
def grade_narrative(scores, program = 'Overall'):
    grades = count_grades(scores, program)
    program_total = str(grades['Count'].sum())

    narr_string = '''The '''+program_total+''' grantees had an overall score of  '''

overall = count_grades(scores)
ga = count_grades(scores, 'GA')
ya = count_grades(scores, 'YA')
ten = count_grades(scores, 'TEN')
sadd = count_grades(scores, 'SADD')



pdb.set_trace()



