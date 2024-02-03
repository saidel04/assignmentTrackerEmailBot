import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage
import datetime
import ssl
import smtplib

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "link your spreadsheet"

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials  =Credentials.from_authorized_user_file("token.json",SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json",SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json","w") as token:
            token.write(credentials.to_json())

    while True:
        today = datetime.date.today
        last_sent_message = None

        if today != last_sent_message:
            try:
                service = build("sheets","v4",credentials=credentials)
                sheets = service.spreadsheets()

                comp2406_dates = sheets.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!J2:J23").execute()
                erth2415_dates = sheets.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!B8:B13").execute()
                biol1902_dates = sheets.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!B2:B4").execute()
                comp2404_dates = sheets.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!F2:F16").execute()
                comp2804_dates = sheets.values().get(spreadsheetId = SPREADSHEET_ID, range="Sheet1!B20:B24").execute()

                biol1902 = biol1902_dates.get("values", [])
                comp2406 = comp2406_dates.get("values",[])
                comp2404 = comp2404_dates.get("values",[])
                comp2804 = comp2804_dates.get("values",[])
                erth2415 = erth2415_dates.get("values",[])

                allClasses = [biol1902,comp2406,comp2404,comp2804,erth2415]
                for i in range(len(allClasses)):
                    for j in range(len(allClasses[i])):
                        if dateCompare(allClasses[i][j][0]) <= 7 and dateCompare(allClasses[i][j][0]) > 0:
                            if i == 0:
                                theClass = "BIOL1902"
                            elif i == 1:
                                theClass = "COMP2406"
                            elif i == 2:
                                theClass = "COMP2404"
                            elif i == 3:
                                theClass = "COMP2804"
                            elif i == 4:
                                theClass = "ERTH2415"
                            
                            message = "You have an assignment due for " + theClass
                            sendEmail(message) 
                
                last_sent_message = datetime.date.today
    
            
                    

            except HttpError as error:
                print(error)

def dateCompare(dt):
    due_date = datetime.datetime.strptime(dt,"%m/%d/%Y")
    current_date = datetime.datetime.today()

    difference = due_date - current_date
    return difference.days

def sendEmail(message):

    print(" got here")
    EMAIL_SENDER = "email setup required"
    APP_PASSWORD = "googlesheets API setup required"
    SUBJECT = "ASSIGNMENT DUE SOON! "
    em = EmailMessage()
    em['From'] = EMAIL_SENDER
    em['To'] = EMAIL_SENDER
    em['Subject'] = SUBJECT
    em.set_content(message)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com',465, context = context) as smtp:
            smtp.login(EMAIL_SENDER, APP_PASSWORD)
            smtp.sendmail(EMAIL_SENDER,EMAIL_SENDER,em.as_string())

if __name__ == "__main__":
    main()
