#!/usr/bin/env python
# coding: utf-8

# In[1]:


#authenticates with GSheets API
#GSheet Link: https://docs.google.com/spreadsheets/d/1UBQG_gf087GhatnGxNeGZmDSJti5WlUgdUhOe0QMl-Q/edit#gid=0

import numpy as np
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account

privKey = 'itmgt-final-project-a9d84967644a.json'
gsheetId = '1UBQG_gf087GhatnGxNeGZmDSJti5WlUgdUhOe0QMl-Q'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = None
creds = service_account.Credentials.from_service_account_file(privKey, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
openSheet = 'EmployeeDatabase'

#temporarily assigns values
idNumber = 0
firstName = ""
lastName = ""
vacStatus = ""
saveScan = False


# In[2]:


#selects the database sheet

def pull_sheet_data(SCOPES,gsheetId,openSheet):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=gsheetId,range=openSheet).execute()
    values = result.get('values', [])
    
    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=gsheetId, range=openSheet).execute()
        data = rows.get('values')
        print('Data copied')
        return data
        

# get all the records of the data
recordsData = pull_sheet_data(SCOPES,gsheetId,openSheet)
df = pd.DataFrame.from_dict(recordsData)
df = df.drop(index=[0])
df.columns = ['Employee Number', 'First Name', 'Last Name', 'Vaccination Status']
df


# In[3]:


#QR Code Scanner
import time
def codeScan():
    import cv2, webbrowser
    img = cv2.imread('code.png')
    cap = cv2.VideoCapture(0)
    detector = cv2.QRCodeDetector()
    while True:
        _, img = cap.read()
        data, bbox, _ = detector.detectAndDecode(img)
        if(data):
            a = str(data)
            time.sleep(1)
            break
            return(a)
    cap.release()
    
    #decodes ID number
    import base64
    base64_message = a
    base64_bytes = base64_message.encode('ascii')
    message_bytes = base64.b64decode(base64_bytes)
    idNumber = str(message_bytes.decode('ascii'))
    
    if (idNumber in df['Employee Number'].values and idNumber != ''):
        firstName = df.loc[df['Employee Number'] == idNumber, 'First Name'].item()
        lastName = df.loc[df['Employee Number'] == idNumber, 'Last Name'].item()
        vacStatus = df.loc[df['Employee Number'] == idNumber, 'Vaccination Status'].item()
        saveScan = True
    else:
        firstName = ''
        lastName = ''
        vacStatus = ''
        saveScan = False

    if saveScan == True:   
        from datetime import datetime
        getDate = datetime.today().strftime('%Y/%m/%d')
        getTime = datetime.now().strftime('%H:%M:%S')

        #exports data to gsheets
        scanInfo = [[idNumber,firstName,lastName,getTime,getDate]]
        newScan = service.spreadsheets().values().append(spreadsheetId=gsheetId, range='ScanRecord!A1:E1', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body={'values':scanInfo}).execute()         
    
    return(idNumber, firstName, lastName, vacStatus);


# In[5]:


#selects the database sheet 'ScanRecord'
openSheet2 = 'ScanRecord'

def pull_sheet_data2(SCOPES,gsheetId,openSheet2):
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=gsheetId,
        range=openSheet2).execute()
    values = result.get('values', [])
    
    if not values:
        print('No data found.')
    else:
        rows = sheet.values().get(spreadsheetId=gsheetId,
                                  range=openSheet2).execute()
        data = rows.get('values')
        print("COMPLETE: Data copied")
        return data

# get all the records of the data
recordsData2 = pull_sheet_data2(SCOPES,gsheetId,openSheet2)
df2 = pd.DataFrame.from_dict(recordsData2)
df2 = df2.drop(index=[0])
df2.columns = ["Employee Number", "First Name", "Last Name", "Time of Scan", "Date of Scan"]
display(df2)

df2["Date of Scan"] = df2["Date of Scan"].str[:-3] #Removes the day part of the date

breakdown_id = pd.pivot_table(df2, values = "Time of Scan", index = "Employee Number", columns = "Date of Scan", fill_value = 0, aggfunc = "nunique")
#Creates pivot table with the rows being the employee number and the columns being the date of scan
breakdown_id_df = breakdown_id.reset_index() #Makes pivot table into dataframe
breakdown_id_df["total"] = breakdown_id_df.sum(axis = 1, numeric_only=True)

highest_scan_count = None
highest_scan_count = breakdown_id_df.sort_values("total", ascending = False)
highest_scan_count.reset_index(drop = True, inplace = True)
display(highest_scan_count)


# In[ ]:


#website loader
from flask import Flask, flash, redirect, render_template, render_template_string, request
app = Flask(__name__)

@app.route('/')
def home():
    return render_template('homePageNew.html')

@app.route('/scanner')
def scanner():
    return render_template('scannerPageNew.html')

@app.route('/scanScript')
def scanScript():
    return (codeScan())
    return render_template('scannedPageNew.html')

@app.route('/scanned')
def scanned():
    if (saveScan):
        saveRec()
    return render_template('scannedPageNew.html')

if __name__ == '__main__':
     app.run(port=8080)

