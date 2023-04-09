# MIT License

# Copyright (c) 2023 Pranav Mehendale

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

"""
Acunetix Attendance Manager
"""

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os.path
import cv2
from time import sleep
from datetime import datetime
import pandas as pd
import json
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger


def accessDataSheet():
    """
    This function is used to connect to the google sheet file 
    where the data is stored.
    """

    global spreadsheet_id

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                params['credentials_path'], SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    spreadsheet_id = params['spreadsheet_id']
    range_name = f"{params['sheet_name']}!A:F"

    # Build the Sheets API client
    global service
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API to get data
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    local_df = pd.DataFrame(values, columns=[values[0][i] for i in range(
        len(values[0]))], index=[values[i][0] for i in range(len(values))])

    return local_df


def markAttendance(rId):
    global df
    try:
        row = df.loc[df['Registration ID'] == rId].iloc[0]
    except IndexError:
        print("Registration ID not in the database.")
        row = None

    # In attendance column mark ‘P’ if it is empty.
    if row is not None:
        if pd.isnull(row['Attendance']):
            df.loc[df['Registration ID'] == rId, 'Attendance'] = 'P'
            now = datetime.now().strftime("%H:%M:%S")
            df.loc[df['Registration ID'] == rId, 'Time Stamp'] = now
        else:
            print("Attendance already marked.")

    # writes the data in local .txt file as an backup
    file = open("Backup.txt", "a")
    file.write(
        f"Attendance for {df.loc[df['Registration ID'] == rId, 'Name'].values[0]}\t{df.loc[df['Registration ID']==rId,'Phone'].values[0]} marked at {now}!\n")

    return f"Attendance for {df.loc[df['Registration ID'] == rId, 'Name'].values[0]}\t{df.loc[df['Registration ID']==rId,'Phone'].values[0]} marked at {now}!"


def scanForQR():
    """
    This function is used for scanning QR code.
    """
    global service

    # Initialize QR code detector
    qr_detector = cv2.QRCodeDetector()

    # Capture frame from camera
    cap = cv2.VideoCapture(int(params['cam_no']))

    # Create window for video feed
    cv2.namedWindow("Acunetix Attendace System", cv2.WINDOW_AUTOSIZE)
    cap.set(3, 1)
    cap.set(4, 1)

    # List of Ids already marked present from local device
    present = []

    while True:
        # Exit if 'q' key is pressed
        key = cv2.waitKeyEx(1)
        if key == ord('q'):
            break

        # Read Frame
        ret, frame = cap.read()
        try:
            # Decode QR code
            data, bbox, _ = qr_detector.detectAndDecode(frame)
        except:
            continue

        # Show video feed in window
        cv2.imshow("Acunetix Attendace Manager", frame)

        # Print results
        if data:
            RegistrationID = data.lstrip("ID: ")
            if RegistrationID in present:
                print(
                    f"Attendance already marked at {df.loc[df['Registration ID'] == RegistrationID, 'Time Stamp'].values[0]}!\nIf it's a problem contact Data Manager team!")
                sleep(2)
                continue
            else:
                present.append(RegistrationID)
                # print(present)
                try:
                    confirmation = markAttendance(RegistrationID)
                except:
                    confirmation = 0
            print(confirmation)
            sleep(1.5)

        # Wait for a short time to limit CPU usage
        sleep(0.2)

    # Release resources
    cap.release()
    cv2.destroyAllWindows()


def updateSheet():
    global df
    global spreadsheet_id

    old_df = accessDataSheet()
    if old_df.equals(df):
        return
    else:
        range_name = f"{params['sheet_name']}!E:F"
        values = df.iloc[:, -2:].values.tolist()
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))


if __name__ == "__main__":
    global df
    global spreadsheet_id

    with open('config.json', 'r') as c:
        params = json.load(c)['params']
        print(params['sheet_name'])
        print(params['spreadsheet_id'])

    df = accessDataSheet()

    scheduler = BackgroundScheduler()
    scheduler.add_job(updateSheet, trigger=IntervalTrigger(
        seconds=params['sec_interval']))
    scheduler.start()

    scanForQR()
    updateSheet()
