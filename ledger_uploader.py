from __future__ import print_function

import configparser
import os.path
import pickle

from appJar import gui
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# This variable specifies the name of a file that contains the
# OAuth 2.0 information for this application, including its client_id
# and client_secret.
CLIENT_SECRETS_FILE = "credentials.json"

# This OAuth 2.0 access scope allows for full read/write access to the
# authenticated user's account and requires requests to use an SSL
# connection.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive.file"]

# This file stores the user's access and refresh tokens and is created
# automatically when the authorization flow completes for the first
# time. This specifies where the file is stored
TOKEN_PICKLE_FILE = "token.pickle"

# Create the GUI
APP_GUI = gui(showIcon=False)
APP_GUI.setOnTop()


def upload_ledger(dir_name: str):
    """Uploads the ledger to the Google Sheet with the macro.

    :param dir_name: the default directory to open the xlsx file from
    :type dir_name: str
    """

    # Open the spreadsheet
    try:
        xlsx_filepath = APP_GUI.openBox(title="Open spreadsheet",
                                        dirName=dir_name,
                                        fileTypes=[("Office Open XML Workbook", "*.xlsx")],
                                        asFile=True).name
        APP_GUI.removeAllWidgets()

    except Exception as err:
        print("There was an error opening the spreadsheet.")
        raise SystemExit(err)

    # Authenticate and retrieve the required services
    services = authorize()

    file_metadata = {"name": "Latest C30 Ledger",
                     "mimeType": "application/vnd.google-apps.spreadsheet"}

    media = MediaFileUpload(xlsx_filepath,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            resumable=True)
    file = services["drive"].files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()


def authorize() -> dict:
    """Authorizes access to the user's Drive and Sheets.

    :return: the authenticated services.
    :rtype: dict
    """

    credentials = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            credentials = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and \
                credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(credentials, token)

    # Build both services and return them in a dict
    drive_service = build('drive', 'v3', credentials=credentials,
                          cache_discovery=False)
    sheets_service = build('sheets', 'v4', credentials=credentials,
                           cache_discovery=False)
    return {"drive": drive_service, "sheets": sheets_service}


def main():
    """Runs the program."""

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    dir_name = parser.get("expense365", "dir_name")

    upload_ledger(dir_name)


if __name__ == "__main__":
    main()
