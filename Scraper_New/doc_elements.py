import re

document = {
    "docname": "",
    "organization": "",
    "program": "",
    "year": "",
    "docnumber": ""
}

monthly = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl08_hplMenuItem",
    "href": "ctl00$hdnSortOrder"
}

narrative = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl09_hplMenuItem",
    "name": "ctl00$hdnSortOrder"
}

enforcement = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl10_hplMenuItem",
}

milestone = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl11_hplMenuItem",
}

statistical = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl12_hplMenuItem",
}

expense = {
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl09_hplMenuItem",
}

expense_table = {
    "tag": "table", 
    "id": "ctl00_cphPageContent_tblPageTable",
}

claim_status = { 
    "tag": "table", 
    "id": "ctl00_cphPageContent_wclDocumentInformation_dgdDocumentInformation",
}

project_info = { 
    "tag": "a",
    "id": "ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl13_hplMenuItem",
    "text": "Application Project Information"
}

project_title = { 
    "tag": "input",
    "id": re.compile('^ctl00_cphPageContent_txtProjecttitle')
}

project_summary = { 
    "tag": "textarea", 
    "id": re.compile('^ctl00_cphPageContent_txtProjectSummary')
}

cost_category = { 
    "tag": "a", 
    "id": re.compile('^ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl45_hplMenuItem'),
    "text": "Cost Category Summary"
}

budget_table = { 
    "tag": "table", 
    "id": re.compile('^ctl00_cphPageContent_tblPageTable12')
}

total_budget = { 
    "tag": "span", 
    "id": re.compile('^ctl00_cphPageContent_lblTotalFedAmt')
}

final_expenditures = { 
    "tag": "a",
    "id": re.compile('^ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl12_hplMenuItem'),
    "text": "Expenditures"
}

total_spent = { 
    "tag": "span",
    "id": re.compile('^ctl00_cphPageContent_lblTotal_')
}

final_activities = { 
    "tag": "a",
    "id": re.compile('^ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl09'),
    "text": re.compile('^Goals/Objectives/Activities')
}

ya_final_activities = { 
    "tag": "a",
    "id": re.compile('^ctl00_cphPageContent_wclDocumentForms_dgdDocumentMenu_ctl09'),
    "text": re.compile('^Goals/Objectives/Activities ')
}

select_activity = { 
    "tag": "select", 
    "title": "Object Pages"
}

activity_options = { 
    "tag": "option", 
}

activity_funded = { 
    "tag": "span", 
    "id": re.compile('^ctl00_cphPageContent_lblActivity')
}

sadd_activity_funded = { 
    "tag": "span", 
    "id": re.compile('^ctl00_cphPageContent_lblObjective')
}

activity_desc = { 
    "tag": "textarea", 
    "accomplishments": re.compile('^ctl00_cphPageContent_txtAccomplishments'),
    "challenges": re.compile('^ctl00_cphPageContent_txtChallenges')
}