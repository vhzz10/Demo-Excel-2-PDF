import logging, dominate, openpyxl, sys, os
from dominate.tags import *


PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__) + '/../..')

template_path = f'{PROJECT_ROOT}/tests/fixtures/mer/MER_Quotation_Tool_Template.test.xlsx'
path = f'{PROJECT_ROOT}/tests/fixtures/mer/'

def excel_sheet_processor(workbookfilepath: str):
    '''
        Converts Excel Sheet data into an array with dictionaries
        arguments:
            workbookfilepath: filepath to workbook
        returns:
            list
    '''
    #open workbook
    wb = openpyxl.load_workbook(workbookfilepath)
    #select active sheet
    ws = wb.active
    #define list
    workbook_list = []
    #define keys as a list
    my_keys = []
    for col in range(0, ws.max_column):
        my_keys.append(   ws.cell(row=1, column=col + 1).value  )
    #define loop to convert rows to dictionaries
    for row in range( 2,  ws.max_row ):
            #create dictionary
            dictionary = {}
            for pos in range(0, len(my_keys)):
                dictionary[my_keys[pos]] = ws.cell(row=row, column=pos+1).value
            workbook_list.append(dictionary)
    return workbook_list

def list_diction_to_html(workbook_list: list):
    '''
        Creates HTML Pages with table containing "list" data
        arguments:
            workbook_list; a list of dictionaries
        returns:
            html file path; str
    '''

    doc = dominate.document(title="Excel Spread Sheet") #sets html title tag
    with doc.head:
        link(rel="stylesheet", href="style.css")
    with doc:
        with div(id="excel_table").add(table()):

            with thead():
                #add header
                dictionary = workbook_list[0]
                for key in dictionary.keys():
                    table_header = td()
                    if key:
                        print(key)
                        table_header.add(p(key))

            for dictionary in workbook_list:
                #loop through row; create table row
                table_row = tr(cls="excel_table_row")
                #loop through each key in dictionary
                for key in dictionary:
                    with table_row.add(td()):
                        if key:
                            p(dictionary[key])
    return str(doc) #turns the document into a string

def save_dom_to_html(dom):
    '''
        Saves DOM string into newly generated HTML file
        arguments:
            dom- str
        returns:
            filepath; str
    '''
    filepath = os.path.abspath(f"{path}excel.html")
    htmfile = open(filepath, "w")
    htmfile.write(dom)
    htmfile.close()
    return filepath
if __name__ == "__main__":
    #filepath = os.path.abspath(template_path)
    list_work = excel_sheet_processor(template_path)
    if list_work:
        dom = list_diction_to_html(list_work)
        save_dom_to_html(dom)