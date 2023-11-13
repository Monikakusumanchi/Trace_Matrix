from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from fastapi.responses import JSONResponse
import pandas as pd
import numpy as np
import time
import re
from typing import Annotated
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import GoogleCredentials
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from google.auth import default
from collections import defaultdict
from gspread_formatting import *
from fastapi import Request, FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates


templates = Jinja2Templates(directory="templates")
app = FastAPI()

# Create a sample dataframe
df = pd.DataFrame(columns=['Enter the url', 'Category', 'Grant Access'])
def formatting(worksheet):
    all_values = worksheet.get_all_values()

    # Calculate the number of filled rows and columns
    num_filled_rows = len(all_values)
    num_filled_columns = len(all_values[0]) if num_filled_rows > 0 else 0
    cell_range = f"A2:{chr(ord('A') + num_filled_columns - 1)}{num_filled_rows + 1}"
    header_fmt = cellFormat(
        backgroundColor=color(1, 0.8, 1),
        textFormat=textFormat(bold=True, foregroundColor=color(0, 0, 0)),
        horizontalAlignment='CENTER',
        borders=borders(
            top=border('SOLID', color=color(0, 0, 0)),
            bottom=border('SOLID', color=color(0, 0, 0)),
            left=border('SOLID', color=color(0, 0, 0)),
            right=border('SOLID', color=color(0, 0, 0))
        )
    )

    data_fmt = cellFormat(
        horizontalAlignment='LEFT',
        borders=borders(
            top=border('SOLID', color=color(0, 0, 0)),
            bottom=border('SOLID', color=color(0, 0, 0)),
            left=border('SOLID', color=color(0, 0, 0)),
            right=border('SOLID', color=color(0, 0, 0))
        )
    )

    headers = worksheet.row_values(1)  # Get the headers from the first row

    num_columns = len(headers)
    max_lengths = []

    for i in range(num_columns):
        column_values = worksheet.col_values(i + 1)  # Google Sheets columns are 1-indexed
        length = max(len(value) for value in column_values)
        max_length = 250 if length>=40 else 140
        max_lengths.append((chr(65 + i), max_length))
    set_column_widths(worksheet,max_lengths)
    worksheet.format("A:ZZ", {"wrapStrategy": "WRAP"})
    cell_format = {
        "verticalAlignment": "TOP"
    }
    worksheet.format('A1:ZZ1000', cell_format)
    format_cell_range(worksheet, '1', header_fmt)

    # Format the data rows
    format_cell_range(worksheet, cell_range, data_fmt)

    # Freeze the header row
    set_frozen(worksheet, rows=1)

    # Format the header row
    format_cell_range(worksheet, '1', header_fmt)

    # Format the data rows
    format_cell_range(worksheet, cell_range, data_fmt)

    # Freeze the header row
    set_frozen(worksheet, rows=1)

    print("Successfully made changes in the Google sheet.")
    
def fn(ijno_val, val):
    if ijno_val.startswith(val):
        return "X"
    else:
        return "" 

def one_master_sheet(gc,sht1,Master_columns):
    worksheet_list = sht1.worksheets()
    if len(worksheet_list) == 1 and worksheet_list[0].title == 'Master':
        print("The spreadsheet have only one 'Master' sheet.")
        worksheet = sht1.worksheet('Master')
        all_records = worksheet.get_all_values()
        data = pd.DataFrame(all_records)
        cols = data[:1][0]
        for i in range(len(cols)):
            cols[i] = cols[i].replace("\n"," ").strip()
        for col in Master_columns:
            if col not in cols:
                print (col)        
                return True
            else:
                print("The spreadsheet does not have only one 'Master' sheet.")
                return False


def execute_RiskAnalysis(gc,sht1,FILE_ID,user_input,credentials):
    #sht1 = gc.open_by_key(FILE_ID)
    worksheet = sht1.worksheet('Master')
    data = worksheet.get_all_values()
    df=pd.DataFrame(data)
    print("data_of_df is:",df)
    cols = data[:1][0]
    print("columns are:",cols)
    for i in range(len(cols)):
        cols[i] = cols[i].replace("\n"," ").strip()
    df = pd.DataFrame(data[1:], columns = cols)
    df = df.iloc[:, :20]

    df.info()
    df_new = pd.DataFrame(np.repeat(df.values, 4, axis=0),columns = df.columns)
    df_new.info()
    df_new["Controls"] = ""
    for i in range(len(df)):
        df_new.at[4*i,"Controls"] = df.iloc[i][8]
        df_new.at[4*i+1,"Controls"] = df.iloc[i][9]
        df_new.at[4*i+2,"Controls"] = df.iloc[i][13]
        df_new.at[4*i+3,"Controls"] = df.iloc[i][14]
    df_new['IQ'] = df_new["Controls"].apply(lambda X: fn(X,"IQ"))
    df_new['OQ'] = df_new["Controls"].apply(lambda X: fn(X,"OQ"))
    df_new['PQ'] = df_new["Controls"].apply(lambda X: fn(X,"PQ"))
    df_new['SOP'] = df_new["Controls"].apply(lambda X: fn(X,"SOP"))
    df_new = df_new[df_new['Controls'].str.lower().str.contains("none") == False]
    df2 = pd.DataFrame()
    df2['Controls'] = df_new['Controls']
    df2['Function of Field Unit'] = df_new['Function of field unit']

    df2['Requirement from URS or RA'] = df_new['Controls'] + " " + df_new['Function of field unit']
    df2['URS Num'] = " "
    df2['RA Num'] = df_new['Row ID#']
    df2['IQ'] = df_new['IQ']
    df2['OQ'] = df_new['OQ']
    df2['PQ'] = df_new['PQ']
    df2['SOP'] = df_new['SOP']
    df3 = df2.groupby('Requirement from URS or RA')['RA Num'].apply(list)
    df2 = df2.drop_duplicates('Requirement from URS or RA', keep='first')
    df2["RA Num"] = df2["Requirement from URS or RA"].apply(lambda x:",".join(df3[x]))
    df2["Name of Document"] = df2["Controls"]
    sh = gc.open_by_url(user_input)
    worksheet_name = 'TM 1Step RA'
    worksheet = None
    try:
        worksheet = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet is not found, create it
        worksheet = sh.add_worksheet(title=worksheet_name, rows=1, cols=len(df2.columns))
    else:
        # If the worksheet exists, clear its content
        worksheet.clear()
    worksheet.update('A1', [df2.columns.values.tolist()])  # Update header

    
    worksheet.append_rows(df2.values.tolist())
    formatting(worksheet)
    worksheet = sht1.worksheet('TM 1Step RA')
    #<========================TM 1Step RA===============================>
    new_df_step4_rano_dict_2 = defaultdict(set)

    for i in range(len(df2['Requirement from URS or RA'])):
        ra_nums = set(map(int, str(df2.iloc[i]['RA Num']).split(',')))
        req_type = df2.iloc[i]['Requirement from URS or RA']

        if 'OQ alarm Test' in req_type:
            new_df_step4_rano_dict_2['OQ alarm Test: all sensors'].update(ra_nums)
        elif 'OQ calibration Sensor' in req_type:
            new_df_step4_rano_dict_2['OQ calibration-all sensors'].update(ra_nums)
        else:
            new_df_step4_rano_dict_2[str(req_type)].update(ra_nums)
    name_of_OQ_test = {
    'OQ calibration-all sensors': 1,
    'OQ alarm Test: all sensors':2,
    'OQ Test reaction of system in the case of power loss':3,
    'OQ Test Access Control':4,
    'OQ Test Verification of User Accounts and responsibilities of all users incl. Emergency accounts':5,
    'OQ Test Verification of syncronization with time server':6,
    'OQ Test Verification of batch functions':7,
    'OQ Test Verification of recipies':8,
    'OQ Test Verification of backup after program changes & desaster recovery':9,
    'OQ Test Verification that all ports of the PC-System are deactivated':10,
    'OQ Test Verification of Gateway to Data Historian':11,
    'OQ Test Verification of Audit trail for GMP relevant functions':12,
    'OQ Test Verification of GMP relevant counters':13,
    'OQ Test of shift register':14,
    'OQ Test Software review for GAMP5 Class 5 software':15,
    'OQ Test Verification of redundant Servers':16,
    'OQ functional test':17.1,
    }
    df_name_of_OQ_test = pd.DataFrame(list(name_of_OQ_test.items()), columns=['key', 'value'])
    df_dicts = pd.DataFrame(list(new_df_step4_rano_dict_2.items()), columns=['key', 'value'])
    new_df_step4_rano_df = pd.DataFrame(list(new_df_step4_rano_dict_2.items()), columns=['Requirement from URS or RA', 'RA Num'])
    new_df_step4_rano_df['RA Num'] = new_df_step4_rano_df['RA Num'].apply(lambda x: str(sorted(x))[1:-1])
    count = 17
    list_of_OQ = []
    for index, row in df_name_of_OQ_test.iterrows():
        for key, val in new_df_step4_rano_df.iterrows():
            if row['key'] in val['Requirement from URS or RA']:
                if row['key'] == 'OQ functional test':
                    count += 0.1
                    row_val = count
                else:
                    row_val = row['value']
                print(row['key'],round(row_val,2))
                list_of_OQ.append(round(row_val,2))
    print(list_of_OQ)
    cols = "Requirement from URS or RA,URS Num,RA Num,Name of document,IQ,OQ,PQ,SOP".split(',')
    rows = []

    for index, row in new_df_step4_rano_df.iterrows():
        formatted_float_value = "{:.1f}".format(list_of_OQ[index])
        new_row = {
            'Requirement from URS or RA': row['Requirement from URS or RA'],
            'URS Num': ' ',
            'RA Num': row['RA Num'],
            'Name of document': row['Requirement from URS or RA'],
            'IQ': " ",
            'OQ': formatted_float_value,
            'PQ': " ",
            'SOP': " "
        }
        rows.append(new_row)
    new_df_step4 = pd.DataFrame(rows, columns=cols)
    worksheet_name = 'TM 4Step RA'
    worksheet = None
    try:
        worksheet = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet is not found, create it
        worksheet = sh.add_worksheet(title=worksheet_name, rows=1, cols=len(df2.columns))
    else:
        # If the worksheet exists, clear its content
        worksheet.clear()
    worksheet.update('A1', [new_df_step4.columns.values.tolist()])  # Update header

    
    worksheet.append_rows(new_df_step4.values.tolist())
    formatting(worksheet)
    worksheet = sht1.worksheet('TM 4Step RA')
    #<========================TM 1Step RA===============================>
    print(f"Trace Matrix successfully generated. [Click here to view]({user_input})")
    output_message = f"Trace Matrix successfully generated . <a href='{user_input}'>Click Here to view</a>"
    return output_message


def execute_URS(gc,sht1,FILE_ID,user_input,credentials):
    #gc = gspread.authorize(credentials)
    #sht1 = gc.open_by_key(FILE_ID)
    worksheet = sht1.worksheet('Master')
    all_records = worksheet.get_all_records()
    df = pd.DataFrame(worksheet.get_all_records())
    mask = df['QP, BEA or ES'] == 'QP'
    filtered_df = df.loc[mask, ['Requirement-ID \nClient', 'DI Control', 'QP, BEA or ES', 'Requirement Description','Tag (QualificationDocuments)']]
    filtered_df.columns = ['Requirement Num', 'DI Control', 'GxP Critical', 'Requirement Description','Tag (QualificationDocuments)']
    new_df_step1 = pd.DataFrame(filtered_df)
    new_df_step1.reset_index(drop=True, inplace=True)
    new_df_step1.fillna('', inplace=True)
    cols = "Requirement from URS or RA,URS Num,RA Num,Name of Document,IQ,OQ,PQ,SOP".split(',')
    new_df_step2 = pd.DataFrame(columns=cols)

    for i in range(len(new_df_step1)):
        row = new_df_step1.iloc[i]
        document_tags = row['Tag (QualificationDocuments)']
        new_row = [
            row['Requirement Description'],
            row['Requirement Num'],
            " ",
            row['Tag (QualificationDocuments)'],
            fn(document_tags, "IQ"),
            fn(document_tags, "OQ"),
            fn(document_tags, "PQ"),
            fn(document_tags, "SOP")
        ]
        new_df_step2.loc[len(new_df_step2)] = new_row
    #<========================20_URS_1===============================>
    try:
        worksheet = sht1.worksheet('30_1 st step TM')
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sht1.add_worksheet(title='30_1 st step TM', rows="115", cols="20")

    # Clear the worksheet
    worksheet.clear()
    
    # Update the Google Sheets worksheet
    worksheet = sht1.worksheet('30_1 st step TM')
    worksheet.update([new_df_step2.columns.values.tolist()] + new_df_step2.values.tolist())
    formatting(worksheet)    
    output_message = f"Trace Matrix successfully generated . <a href='{user_input}'>Click Here to view</a>"

    return output_message

@app.get("/")
async def dynamic_file(request: Request):
    return templates.TemplateResponse("intro.html", {"request": request})

# FastAPI route to retrieve data
@app.get("/data",response_class=HTMLResponse)
async def read_root(request: Request):
    # Render the HTML template with the data from the dataframe
    context = {"request": request}

    return templates.TemplateResponse("index.html", context)
@app.post("/post_data/{data_path:path}")
async def get_data(data_path: str):
    # Process data_path if needed
    # For now, redirect to the original /postdata
    return RedirectResponse(url="/post_data")

@app.post("/post_data")
async def post_data(request: Request, 
                    user_input: Annotated[str,Form(...)],
                    Category: Annotated[str,Form(...)]
                    ):
    if Category not in ("URS", "RA"):
        return {"error": "Invalid Category selection"}
    RA_Master_columns = ['Row ID#', 'Function of field unit', 'Potential \nfailure \nmode',
                                'Potential \nEffects of \nfailure Mode', 'Machine \nreaction',
                                'Potential\ncosequences\nfor the patient', 'Serverity\nRanking\n\nS',
                                'potential\nCauses', 'Current\nPrevention\nControl(s)',
                                'Current \nDetection Control(s)', 'Occurence\nRanking\n\nO',
                                'Detection \nRanking\n\nD', 'Risk Priority\nNumber\nS*O*D\n=RPN',
                                'Mitigation\nPrevention\nControl(s)', 'Mitigation \nDetection \nControl(s)',
                                'Person\naccountable', 'Post\nMitigation\nOccurency\nOp',
                                'Post\nMitigation\nDetection\nDp', 'P-Mitigation\nRisk Priority Number\nSp*Op*Dp\n=RPNp',
                                'Comment',]
    URS_Master_columns = ['Requirement-ID \nLSE', '', 'Requirement-ID \nClient', 'DI Control',
        'QP, BEA or ES', 'Requirement \nGroup', 'IQ-Plan', 'OQ-Test', 'SOP ',
        'Tag (QualificationDocuments)', 'Requirement Description', 'Remark']

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", user_input)
    print(match)
    FILE_ID = match.group(1)

    if match:
        FILE_ID = match.group(1)
        print(FILE_ID)
    else:
        print("Invalid Google Sheets URL")

    credentials = ServiceAccountCredentials.from_json_keyfile_name('red-studio-400805-60aea2585639.json', ['https://www.googleapis.com/auth/spreadsheets'])
    gc = gspread.authorize(credentials)
    sp1 = gc.open_by_key(FILE_ID)
    worksheet = sp1.worksheet('Master')
    data = worksheet.get_all_values()   

    #gc = gspread.authorize(credentials)
    
    #sht1 = gc.open_by_key(FILE_ID)
    output_message = ""

    if Category == "RA":
        if one_master_sheet(gc,sp1,RA_Master_columns):

            output_message = execute_RiskAnalysis(gc,sp1,FILE_ID,user_input,credentials)
            print(output_message)
            return output_message
        else:
            output_message = f"Check the sheet for correct format and check if the spreadsheet does not have only one 'Master' sheet. <a href='{user_input}'>Here</a>"
            print(output_message)

    else :
        if one_master_sheet(gc,sp1,URS_Master_columns):

            output_message = execute_URS(gc,sp1,FILE_ID,user_input,credentials)
            print(output_message)
            return output_message
        else:
            output_message = f"Check the sheet for correct format and check if the spreadsheet does not have only one 'Master' sheet. <a href='{user_input}'>Here</a>"
            print(output_message)
            
    print(output_message)
    
    return JSONResponse(content=output_message)
   # return templates.TemplateResponse("intro.html",{"request": request,  "output":output_message })