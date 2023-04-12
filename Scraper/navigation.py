from pathlib import Path
import pandas

def save_file(doc_dict, data_frame, frame_name):
    path = Path('Grantee Data')
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
