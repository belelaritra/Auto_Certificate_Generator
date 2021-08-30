from __future__ import print_function
from googleapiclient.discovery import build
from google.oauth2 import service_account

import os
import pandas as pd
# === IMPORT SAME FOLDER FILES
import sendemail as email
import certificate as cer
import tkinter_window

# === TKINTER WINDOW
tkw = tkinter_window.Window()

# === SPREADSHEET TYPE 1 : URL (GOOGLE SHEET)  &  TYPE 2 : PATH (LOCAL)
spreadsheet_type = tkw.spreadsheet_type
spreadsheet_url = tkw.spreadsheet_url
spreadsheet_path = tkw.spreadsheet_path

template_path = tkw.template_path
font_style = tkw.font_style
font_colour = tkw.font_colour
fontsize = tkw.fontsize
coord = tkw.coord

# === CUTOFF TYPE 1 : YES ( CUTOFF != NULL )  & TYPE 2 : NO (CUTOFF == NULL )
cutoff_type = tkw.cutoff_type
cutoff = tkw.cutoff

# === EMAIL TYPE 1 : YES (ENABLED)  &  TYPE 2 : NO (DISABLED)
email_type = tkw.email_type
emailid = tkw.emailid
apppass = tkw.apppass
email_subject = tkw.email_subject
email_body = tkw.email_body

# ================================ IF SPREADSHEET_TYPE = 1 -->(URL):GOOGLE=============================================
if spreadsheet_type == 1:
    # === COLLECT SPREADSHEET ID FROM SPREADSHEET URL
    split_url = spreadsheet_url.split("/")
    SPREADSHEET_ID=split_url[5]
    # SPREADSHEET_ID = '1rrg6olzc9tzFxRGcXfD7XDMBks3hVqRdmWRNxFhKC5Q'

    # === GOOGLE API KEY
    SERVICE_ACCOUNT_FILE = 'keys.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # === CREDENTIALS
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # === CALL GOOGLE SHEET API
    sheet = service.spreadsheets()
    # === COLLECTS VALUES FROM RANGE A1 : Z
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range="A1:Z").execute()
    values = result.get('values', [])

    # === NO OF COLUMN
    column = len(values[0])
    # === NO OF ROWS
    row = len(values)
    # INITIALIZE EMAIL, NAME, SCORE INDEX = NULL
    email_index = name_index = score_index = ""

    # === COLLECTS INDEX OF EMAIL, NAME, SCORE
    for i in range(column):
        # === COLLECT INDEX OF EMAIL
        if ('email' in (values[0][i]).lower()):
            email_index = i
        # === COLLECT INDEX OF NAME
        if ('name' in (values[0][i]).lower()):
            name_index = i
        # === COLLECT INDEX OF SCORE
        if ('score' in (values[0][i]).lower()):
            score_index = i
        # === WHEN ALL VALUES ARE COLLECTED --> BREAK THE LOOP
        if (email_index != "" and name_index != "" and score_index != ""):
            break

    # === INITIALIZE BLANK LIST
    name_list = []
    email_list = []
    score_list = []
    percentage_list = []
    image_list = []

    # === COLLECTS NAME, EMAIL, CUTOFF, PERCENTAGE FROM 2D ARRAY VALUE
    for j in range(1, row):
        # === COLLECTS ALL NAMES FROM SPREADSHEET --> TO NAME_LIST
        Name = values[j][name_index]
        name_list.append(Name)
        # === COLLECTS ALL EMAILS FROM SPREADSHEET --> TO EMAIL_LIST
        Email = values[j][email_index]
        email_list.append(Email)
        # === COLLECTS ALL SCORE FROM SPREADSHEET --> TO SCORE_LIST : ( MARKS / FULL_MARKS)
        Score = values[j][score_index]
        score_list.append(Score)
        # === CALCULATES PERCENTAGE FROM SPREADSHEET --> TO PERCENTAGE_LIST
        score_Percentage = eval(values[j][score_index]) * 100
        percentage_list.append(score_Percentage)

    # ======================== CREATE CERTIFICATE FOR ALL --> IF (CUTOFF_TYPE = 2) =====================================
    if cutoff_type == 2:
        for k in range(1, row):
            Name = name_list[k-1]
            image_name = cer.createcertificate(Name, template_path, font_style, font_colour, fontsize, coord)
            image_list.append(image_name)
    # ======================== CREATE CERTIFICATE FOR SELECTED PEOPLE (SCORE>CUTOFF) ===================================
    # =====================--> IF (CUTOFF_TYPE == 1) && (PERCENTAGE_LIST[K-1] >= CUTOFF)
    else:
        for k in range(1, row):
            if percentage_list[k-1] >= float(cutoff):
                Name = name_list[k-1]
                image_name = cer.createcertificate(Name, template_path, font_style, font_colour, fontsize, coord)
                image_list.append(image_name)

    # ================================== SEND EMAIL --> IF (EMAIL_TYPE == 1) ===========================================
    if email_type == 1:
    # FOR ALL --> IF (CUTOFF_TYPE = 2)
        if cutoff_type ==2:
            for m in range(1, row):
                image_name = image_list[m-1]
            # SEND EMAIL : FUNCTION = SENDMAIL & FILE = SENDEMAIL.PY
                Receiver_Email = email_list[m-1]
                email.sendmail(Receiver_Email, image_name, emailid,apppass,email_subject,email_body)
            # === PRINT NAME, EMAIL, SCORE , PERCENTAGE OF ALL SUCCESSFUL SEND
                print("Name : " + name_list[m-1] + "; Email: " + email_list[m-1] + "; score: " + score_list[
                    m-1] + "; Percentage : " + str(percentage_list[m-1]) + "%")
            # === REMOVE THAT IMAGE FILE AFTER EMAIL
                os.remove(image_name)
    # FOR SELECTED PEOPLE (SCORE>CUTOFF) --> IF (CUTOFF_TYPE == 1) && (PERCENTAGE_LIST[M] >= CUTOFF)
        if cutoff_type == 1:
            for m in range(1, row):
                if percentage_list[m-1] >= float(cutoff):
                    image_name = image_list[m-1]
                # SEND EMAIL : FUNCTION = SENDMAIL & FILE = SENDEMAIL.PY
                    Receiver_Email = email_list[m-1]
                    email.sendmail(Receiver_Email, image_name, emailid,apppass,email_subject,email_body)
                # === PRINT NAME, EMAIL, SCORE , PERCENTAGE OF ALL SUCCESSFUL SEND
                    print("Name : " + name_list[m-1] + "; Email: " + email_list[m-1] + "; score: " + score_list[
                        m-1] + "; Percentage : " + str(percentage_list[m-1]) + "%")
                # === REMOVE THAT IMAGE FILE AFTER EMAIL
                    os.remove(image_name)

# ================================ IF SPREADSHEET_TYPE = 2 -->(PATH):LOCAL=============================================
else:
    data = pd.read_excel(spreadsheet_path)
    name_list = data["Name"].tolist()
    email_list = data["Email Address"].tolist()
    score_list = data["Score"].tolist()
    percentage_list = data["Score"].tolist()
    image_list = []
    row = len(name_list)

    # ======================== CREATE CERTIFICATE FOR ALL --> IF (CUTOFF_TYPE = 2) =====================================
    if cutoff_type == 2:
        for k in range(0, row):
            Name = name_list[k]
            image_name = cer.createcertificate(Name, template_path, font_style, font_colour, fontsize, coord)
            image_list.append(image_name)
    # ======================== CREATE CERTIFICATE FOR SELECTED PEOPLE (SCORE>CUTOFF) ===================================
    # =====================--> IF (CUTOFF_TYPE == 1) && (PERCENTAGE_LIST[K] >= CUTOFF)
    else:
        for k in range(0, row):
            if percentage_list[k] >= float(cutoff):
                Name = name_list[k-1]
                image_name = cer.createcertificate(Name, template_path, font_style, font_colour, fontsize, coord)
                image_list.append(image_name)

    # ================================== SEND EMAIL --> IF (EMAIL_TYPE == 1) ===========================================
    if email_type == 1:
    # FOR ALL --> IF (CUTOFF_TYPE = 2)
        if cutoff_type ==2:
            for m in range(0, row):
                image_name = image_list[m]
            # SEND EMAIL : FUNCTION = SENDMAIL & FILE = SENDEMAIL.PY
                Receiver_Email = email_list[m]
                email.sendmail(Receiver_Email, image_name, emailid,apppass,email_subject,email_body)
            # === PRINT NAME, EMAIL, SCORE , PERCENTAGE OF ALL SUCCESSFUL SEND
                print("Name : " + name_list[m] + "; Email: " + email_list[m] + "; score: " + score_list[
                    m] + "; Percentage : " + str(percentage_list[m]) + "%")
            # === REMOVE THAT IMAGE FILE AFTER EMAIL
                os.remove(image_name)
    # FOR SELECTED PEOPLE (SCORE>CUTOFF) --> IF (CUTOFF_TYPE == 1) && (PERCENTAGE_LIST[M] >= CUTOFF)
        if cutoff_type == 1:
            for m in range(1, row):
                if percentage_list[m] >= float(cutoff):
                    image_name = image_list[m]
                # SEND EMAIL : FUNCTION = SENDMAIL & FILE = SENDEMAIL.PY
                    Receiver_Email = email_list[m]
                    email.sendmail(Receiver_Email, image_name, emailid,apppass,email_subject,email_body)
                # === PRINT NAME, EMAIL, SCORE , PERCENTAGE OF ALL SUCCESSFUL SEND
                    print("Name : " + name_list[m] + "; Email: " + email_list[m] + "; score: " + score_list[
                        m] + "; Percentage : " + str(percentage_list[m]) + "%")
                # === REMOVE THAT IMAGE FILE AFTER EMAIL
                    os.remove(image_name)