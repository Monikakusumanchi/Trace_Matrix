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


    #<========================LoD (QualificationDocuments)===============================>
# def lod_create(gc,sht1):   
#         cols = "QualificationStep,QualificationDocument,Company GroningerID".split(',')

#         FILE_ID_2 = '1lTuFfxbYnXgZMWXoJMFr2UcIr0cB16XOIcrMoXZ6jFw'
#         sht2 = gc.open_by_key(FILE_ID_2)

#         worksheet = sht2.worksheet('LoD (QualificationDocuments)')
       
#         df_step4 = pd.DataFrame(worksheet.get_all_values()[0])

#         print(df_step4.head())
#         data = worksheet.get_all_values()
#         data = [row[:3] for row in data]
#         df_step4 = pd.DataFrame(data, columns=cols)

#         try:
#             NEW_SHEET_NAME = 'LoD (QualificationDocuments)'
#             worksheet = sht1.add_worksheet(title=NEW_SHEET_NAME, rows=200, cols=26)
#         except gspread.exceptions.APIError as e:
#             print(e)

#         worksheet = sht1.worksheet('LoD (QualificationDocuments)')
#         worksheet.update(df_step4.values.tolist())
#         formatting(worksheet)



def one_master_sheet(gc,sht1):
    worksheet_list = sht1.worksheets()
    if len(worksheet_list) == 1 and worksheet_list[0].title == 'Master':
        print("The spreadsheet have only one 'Master' sheet.")
        worksheet = sht1.worksheet('Master')
        all_records = worksheet.get_all_values()
        df = pd.DataFrame(all_records)

        expected_columns = ['Row ID#', 'Function of field unit', 'Potential \nfailure \nmode',
                                'Potential \nEffects of \nfailure Mode', 'Machine \nreaction',
                                'Potential\ncosequences\nfor the patient', 'Serverity\nRanking\n\nS',
                                'potential\nCauses', 'Current\nPrevention\nControl(s)',
                                'Current \nDetection Control(s)', 'Occurence\nRanking\n\nO',
                                'Detection \nRanking\n\nD', 'Risk Priority\nNumber\nS*O*D\n=RPN',
                                'Mitigation\nPrevention\nControl(s)', 'Mitigation \nDetection \nControl(s)',
                                'Person\naccountable', 'Post\nMitigation\nOccurency\nOp',
                                'Post\nMitigation\nDetection\nDp', 'P-Mitigation\nRisk Priority Number\nSp*Op*Dp\n=RPNp',
                                'Comment', '', '&']
        return True
    else:
        print("The spreadsheet does not have only one 'Master' sheet.")
        return False


def execute_RiskAnalysis(gc,sht1,FILE_ID,user_input,credentials):
    if one_master_sheet(gc,sht1):
        sht1 = gc.open_by_key(FILE_ID)
        worksheet = sht1.worksheet('Master')
        data = worksheet.get_all_values()
        #<========================TM 1Step RA===============================>
        # Extract headers and ensure uniqueness
        headers = data[0]
        seen_headers = set()
        for i in range(len(headers)):
            header = headers[i]
            if header in seen_headers:
                headers[i] = f"{header}_{i}"
            seen_headers.add(header)
        # Extract relevant columns
        selected_columns = ['I', 'J', 'N', 'O']
        ijno_names = [headers[ord(i) - 65] for i in selected_columns]

        # Prepare lists to create DataFrame
        controls_list = []
        function_list = []
        urs_ra_list = []
        urs_num_list = []
        ra_num_list = []
        name_list = []
        iq_list = []
        oq_list = []
        pq_list = []
        sop_list = []

        # Iterate through the data
        for i in range(1, len(data)):
            row = data[i]
            for column, value in zip(headers, row):
                if column in ijno_names:
                    oq_value, iq_value, pq_value, sop_value = ' ', ' ', ' ', ' '
                    if value.startswith('OQ'):
                        oq_value = 'x'
                    elif value.startswith('IQ'):
                        iq_value = 'x'
                    elif value.startswith('PQ'):
                        pq_value = 'x'
                    elif value.startswith('SOP'):
                        sop_value = 'x'

                    # Append values to respective lists
                    controls_list.append(value)
                    function_list.append(row[headers.index('Function of field unit')])
                    urs_ra_list.append(value + " " + row[headers.index('Function of field unit')])
                    urs_num_list.append(' ')
                    ra_num_list.append(row[headers.index('Row ID#')])
                    name_list.append(' ')
                    iq_list.append(iq_value)
                    oq_list.append(oq_value)
                    pq_list.append(pq_value)
                    sop_list.append(sop_value)

        # Create the new DataFrame
        new_df = pd.DataFrame({
            'Controls': controls_list,
            'Function of Field Unit': function_list,
            'Requirement from URS or RA': urs_ra_list,
            'URS Num': urs_num_list,
            'RA Num': ra_num_list,
            'Name of document': name_list,
            'IQ': iq_list,
            'OQ': oq_list,
            'PQ': pq_list,
            'SOP': sop_list
        })

        #credentials = ServiceAccountCredentials.from_json_keyfile_name('/workspace/Demo/red-studio-400805-60aea2585639.json', ['https://www.googleapis.com/auth/spreadsheets'])
        gc = gspread.authorize(credentials)

        # Access the Google Sheets
        #sheet_url = "https://docs.google.com/spreadsheets/d/19ZW_Eq3ySx925glrnokXDLBvx69_A7sTP02f8-NuB4Q"
        sh = gc.open_by_url(user_input)
        worksheet_name = 'TM 1Step RA'
        worksheet = None
        try:
            worksheet = sh.worksheet(worksheet_name)
        except gspread.exceptions.WorksheetNotFound:
            # If the worksheet is not found, create it
            worksheet = sh.add_worksheet(title=worksheet_name, rows=1, cols=len(new_df.columns))
        else:
            # If the worksheet exists, clear its content
            worksheet.clear()
        worksheet.update('A1', [new_df.columns.values.tolist()])  # Update header

        
        worksheet.append_rows(new_df.values.tolist())
        formatting(worksheet)
        worksheet = sht1.worksheet('TM 1Step RA')
        df_step1 = worksheet.get_all_values()


        #<========================TM 2Step RA===============================>
        cols_step2 = ["Requirement from URS or RA", "URS Num", "RA Num", "Name of document", "IQ", "OQ", "PQ", "SOP"]
        # Extract headers and ensure uniqueness
        headers = df_step1[0]
        seen_headers = set()
        for i in range(len(headers)):
            header = headers[i]
            if header in seen_headers:
                headers[i] = f"{header}_{i}"
            seen_headers.add(header)

        # Create a DataFrame using the remaining rows as data and with the extracted headers
        df_step1 = pd.DataFrame(df_step1[1:], columns=headers)

        # Keep only the required columns
        new_df_step2 = df_step1[cols_step2]
        
        # Filter out rows where "Requirement from URS or RA" contains 'none'
        new_df_step2 = df_step1[~df_step1['Requirement from URS or RA'].str.contains('none')]

        # Keep only the required columns
        new_df_step2 = new_df_step2[cols_step2]

        # Update the Google Sheets for the second sheet
        worksheet_step2_name = 'TM 2Step RA'
        worksheet_step2 = None

        try:
            # Try to access the worksheet, if it exists
            worksheet_step2 = sh.worksheet(worksheet_step2_name)
        except gspread.exceptions.WorksheetNotFound:
            # If the worksheet is not found, create it
            worksheet_step2 = sh.add_worksheet(title=worksheet_step2_name, rows=1, cols=len(new_df_step2.columns))
        else:
        # If the worksheet exists, clear its content
            worksheet_step2.clear()
        # Update header
        worksheet_step2.update('A1', [new_df_step2.columns.values.tolist()])

        # Append data
        worksheet_step2.append_rows(new_df_step2.values.tolist())
        formatting(worksheet_step2)


        #<========================TM 3Step RA===============================>
        cols_step3 = "Requirement from URS or RA,URS Num,RA Num,Name of document,IQ,OQ,PQ,SOP".split(',')

        # Fetch data from TM 2Step RA worksheet
        worksheet_step3 = sht1.worksheet('TM 2Step RA')
        df_step3 = pd.DataFrame(worksheet_step3.get_all_records())
        # Create a new worksheet 'TM 3Step RA' and update header
        worksheet_step3_name = 'TM 3Step RA'
        worksheet_step3 = None

        try:
            # Try to access the worksheet, if it exists
            worksheet_step3 = sht1.worksheet(worksheet_step3_name)
        except gspread.exceptions.WorksheetNotFound:
            # If the worksheet is not found, create it
            worksheet_step3 = sht1.add_worksheet(title=worksheet_step3_name, rows=1, cols=len(cols_step3))
        else:
            worksheet_step3.clear()

        # Group RA Num based on Requirement from URS or RA
        new_df_step3_rano = df_step3.groupby('Requirement from URS or RA')['RA Num'].agg(list).reset_index()['RA Num']

        # Get unique Requirement from URS or RA values
        new_df_step3_req = list(set(df_step3['Requirement from URS or RA']))

        # Prepare a list of dictionaries for the new DataFrame
        new_data_step3 = []

        # Iterate through the data
        for i in range(len(new_df_step3_req)):
            new_row = {
                'Requirement from URS or RA': new_df_step3_req[i],
                'URS Num': ' ',
                'RA Num': str(new_df_step3_rano[i])[1:-1],
                'Name of document': ' ',
                'IQ': df_step3.iloc[i]['IQ'],
                'OQ': df_step3.iloc[i]['OQ'],
                'PQ': df_step3.iloc[i]['PQ'],
                'SOP': df_step3.iloc[i]['SOP']
            }
            new_data_step3.append(new_row)

        # Create the new DataFrame
        new_df_step3 = pd.DataFrame(new_data_step3, columns=cols_step3)
    
        # Update header
        worksheet_step3.update('A1', [cols_step3])

        # Append data
        worksheet_step3.append_rows(new_df_step3.values.tolist())
        formatting(worksheet_step3)


        #<========================TM 4Step RA===============================>
        cols_step4 = "Requirement from URS or RA,URS Num,RA Num,Name of document,IQ,OQ,PQ,SOP".split(',')
        worksheet_step4_name = 'TM 4Step RA'
        
        # Fetch data from TM 3Step RA worksheet
        worksheet_step3 = sht1.worksheet('TM 3Step RA')
        df_step3 = pd.DataFrame(worksheet_step3.get_all_records())

        # Initialize a defaultdict to store RA Num for each Requirement from URS or RA
        new_df_step4_rano_dict = defaultdict(set)

        # Map Requirement from URS or RA to corresponding RA Num
        keywords_mapping = {
            'alarm Test': 'OQ alarm Test: all sensors',
            'calibration Sensor': 'OQ calibration-all sensors',
            'test Sensor of the centuring frame': 'OQ functional test-Sensor of the centuring frame',
            'test Sensor CONVEYOR TUB PRESENCE': 'OQ functional test-Sensor CONVEYOR TUB PRESENCE'
        }

        # Iterate through the data and map keywords to corresponding RA Num
        for i in range(len(df_step3['Requirement from URS or RA'])):
            for keyword, mapped_keyword in keywords_mapping.items():
                if keyword in df_step3.iloc[i]['Requirement from URS or RA']:
                    ra_nums = set(map(int, str(df_step3.iloc[i]['RA Num']).split(', ')))
                    new_df_step4_rano_dict[mapped_keyword].update(ra_nums)

        # Create a new DataFrame for TM 4Step RA
        new_df_step4_rows = []
        for key, value in new_df_step4_rano_dict.items():
            row = [key, ' ', str(value)[1:-1], key, ' ', ' ', ' ', ' ']
            new_df_step4_rows.append(row)

        # Create the new DataFrame
        new_df_step4 = pd.DataFrame(new_df_step4_rows, columns=cols_step4)

        # Update or create the worksheet 'TM 4Step RA'
        worksheet_step4 = None

        try:
            # Try to access the worksheet, if it exists
            worksheet_step4 = sh.worksheet(worksheet_step4_name)
        except gspread.exceptions.WorksheetNotFound:
            # If the worksheet is not found, create it
            worksheet_step4 = sh.add_worksheet(title=worksheet_step4_name, rows=1, cols=len(cols_step4))
        else:
            # If the worksheet exists, clear its content
            worksheet_step4.clear()

        # Update header and append data
        worksheet_step4.update([new_df_step4.columns.values.tolist()] + new_df_step4.values.tolist())
        formatting(worksheet_step4)
        # lod_create(gc,sht1)

        print(f"Trace Matrix successfully generated. [Click here to view]({user_input})")
        output_message = f"Trace Matrix successfully generated . <a href='{user_input}'>Click Here to view</a>"
        print(output_message)

    else :
        output_message = f"Check the sheet for correct format and check if the spreadsheet does not have only one 'Master' sheet. <a href='{user_input}'>Here</a>"
    return output_message

def one_master_sheet_URS(gc,sht1):
    worksheet_list = sht1.worksheets()
    if len(worksheet_list) == 1 and worksheet_list[0].title == 'Master':
        print("The spreadsheet have only one 'Master' sheet.")
        worksheet = sht1.worksheet('Master')
        all_records = worksheet.get_all_values()
        df = pd.DataFrame(all_records)
        expected_columns = ['Requirement-ID \nLSE', '', 'Requirement-ID \nClient', 'DI Control',
        'QP, BEA or ES', 'Requirement \nGroup', 'IQ-Plan', 'OQ-Test', 'SOP ',
        'Tag (QualificationDocuments)', 'Requirement Description', 'Remark']
        return True
    else:
        print("The spreadsheet does not have only one 'Master' sheet.")
        return False


def execute_URS(gc,sht1,FILE_ID,user_input,credentials):
    if one_master_sheet_URS(gc,sht1):
        gc = gspread.authorize(credentials)
        sht1 = gc.open_by_key(FILE_ID)
        worksheet = sht1.worksheet('Master')
        all_records = worksheet.get_all_records()
        #<========================20_URS_1===============================>

        cols = "Requirement Num,DI Control,GxP Critical,Requirement Description,Tag (QualificationDocuments)".split(',')
        headers = all_records[0].keys()  # Extract headers
        try:
            worksheet = sht1.worksheet('20_URS_1')
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sht1.add_worksheet(title='20_URS_1', rows="115", cols="20")

        # Clear the worksheet
        worksheet.clear()
        df_step1 = pd.DataFrame(all_records)

        mask = df_step1['QP, BEA or ES'] == 'QP'
        filtered_df = df_step1.loc[mask, ['Requirement-ID \nClient', 'DI Control', 'QP, BEA or ES', 'Requirement Description','Tag (QualificationDocuments)']]
        filtered_df.columns = ['Requirement Num', 'DI Control', 'GxP Critical', 'Requirement Description','Tag (QualificationDocuments)']

        # Create a new DataFrame and reset the index
        new_df_step1 = pd.DataFrame(filtered_df)
        new_df_step1.reset_index(drop=True, inplace=True)

        # Fill NaN values with empty strings
        new_df_step1.fillna('', inplace=True)

        # Assuming cols is a list of columns to keep
        if cols:
            new_df_step1 = new_df_step1[cols]

        # Update the Google Sheets worksheet
        worksheet = sht1.worksheet('20_URS_1')
        worksheet.update([new_df_step1.columns.values.tolist()] + new_df_step1.values.tolist())
        formatting(worksheet)


        #<========================30_URS Step1===============================>
        cols = "Requirement from URS or RA,URS Num,RA Num,Name of Document,IQ,OQ,PQ,SOP,Tag (QualificationDocuments)".split(',')

        worksheet = sht1.worksheet('20_URS_1')
        df_step2 = pd.DataFrame(worksheet.get_all_records())
        try:
            worksheetURS_step1 = sht1.worksheet('30_URS Step1')
        except gspread.exceptions.WorksheetNotFound:
            worksheetURS_step1 = sht1.add_worksheet(title='30_URS Step1', rows="100", cols="20")

        # Clear the worksheet
        worksheetURS_step1.clear()
        print("data is cleared")
        new_df_step2 = pd.DataFrame(columns=cols)

        for i in range(len(df_step2)):
            new_row = [df_step2.iloc[i]['Requirement Description'],df_step2.iloc[i]['Requirement Num']," ", " ","X"," "," "," ",df_step2.iloc[i]['Tag (QualificationDocuments)']]
            new_df_step2.loc[len(new_df_step2)] = new_row
            
        worksheetURS_step1 = sht1.worksheet('30_URS Step1')
        worksheetURS_step1.update([new_df_step2.columns.values.tolist()] + new_df_step2.values.tolist())
        print("data is appended")
        formatting(worksheetURS_step1)



        #<========================Step2 TM===============================>
        df_step3 = pd.DataFrame(worksheetURS_step1.get_all_records())
        try:
            worksheet_step2= sht1.worksheet('Step2 TM')
        except gspread.exceptions.WorksheetNotFound:
            worksheet_step2 = sht1.add_worksheet(title='Step2 TM', rows="115", cols="20")

        # Clear the worksheet
        worksheet_step2.clear()

        cols = "Tag (QualificationDocuments),Requirement from URS or RA,URS Num,RA Num,Name of Document,IQ,OQ,PQ,SOP".split(',')
        new_df_step3 = pd.DataFrame(columns=cols)

        for i in range(len(df_step3)):
                    new_row = {
                        'Tag (QualificationDocuments)':df_step3.iloc[i]['Tag (QualificationDocuments)'],
                        'Requirement from URS or RA': df_step3.iloc[i]['Requirement from URS or RA'],
                        'URS Num': df_step3.iloc[i]['URS Num'],
                        'RA Num': df_step3.iloc[i]['RA Num'],
                        'Name of Document': df_step3.iloc[i]['Name of Document'],
                        'IQ': df_step3.iloc[i]['IQ'],
                        'OQ': df_step3.iloc[i]['OQ'],
                        'PQ': df_step3.iloc[i]['PQ'],
                        'SOP': df_step3.iloc[i]['SOP'],
                    }
                    new_df_step3 = pd.concat([new_df_step3, pd.DataFrame([new_row])], ignore_index=True)
        worksheet = sht1.worksheet('Step2 TM')
        worksheet.update([new_df_step3.columns.values.tolist()] + new_df_step3.values.tolist())
        formatting(worksheet)
        output_message = f"Trace Matrix successfully generated . <a href='{user_input}'>Click Here to view</a>"
        # lod_create(gc,sht1)
    else:
        output_message = f"Check the sheet for correct format and check if the spreadsheet does not have only one 'Master' sheet. <a href='{user_input}'>Here</a>"
        print(output_message)
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

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", user_input)
    print(match)
    FILE_ID = match.group(1)

    if match:
        FILE_ID = match.group(1)
        print(FILE_ID)
    else:
        print("Invalid Google Sheets URL")
        

    #FILE_ID="19ZW_Eq3ySx925glrnokXDLBvx69_A7sTP02f8-NuB4Q"
    credentials = ServiceAccountCredentials.from_json_keyfile_name('red-studio-400805-60aea2585639.json', ['https://www.googleapis.com/auth/spreadsheets'])

    gc = gspread.authorize(credentials)
    
    sht1 = gc.open_by_key(FILE_ID)
    output_message = ""

    if Category == "RA":
        output_message = execute_RiskAnalysis(gc,sht1,FILE_ID,user_input,credentials)
        print(output_message)
  

    else :
        output_message = execute_URS(gc,sht1,FILE_ID,user_input,credentials)
        print(output_message)

    print(output_message)
    
    return JSONResponse(content=output_message)
   # return templates.TemplateResponse("intro.html",{"request": request,  "output":output_message })