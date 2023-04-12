


def change_asterisk(classification, save):
    result = classification
    if save == True and classification.find('*') is not None:
        result = classification.replace('*', '#')
    if save == False and classification.find('#') is not None:
        result = classification.replace('#', '*')

    return result


def parse_doc_title(org_double):

    parsed = org_double[1].split("-")
    docnum = ""
    if len(parsed) > 4:
        docnum = parsed[4]

    document = {
    "organization": org_double[0],
    "grantname": change_asterisk(org_double[1], True),
    "program": parsed[0],
    "year": parsed[1],
    "classification": change_asterisk(parsed[2], True),
    "id": parsed[3],
    "docnumber": docnum
    }
    
    return document