import base64
import configparser
import os
import pickle
import smtplib
import ssl

import googleapiclient
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from ledger_fetcher import download_pdf, convert_to_xlsx, upload_ledger

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# This variable specifies the name of a file that contains the
# OAuth 2.0 information for this application, including its client_id
# and client_secret.
CLIENT_SECRETS_FILE = "credentials.json"

# This OAuth 2.0 access scope allows for full read/write access to the
# authenticated user's account and requires requests to use an SSL
# connection.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# This file stores the user's access and refresh tokens and is created
# automatically when the authorization flow completes for the first
# time. This specifies where the file is stored
TOKEN_PICKLE_FILE = "ledger_checker_token.pickle"


def authorize() -> googleapiclient.discovery.Resource:
    """Authorizes access to the user's scripts.

    :return: the authenticated service
    :rtype: googleapiclient.discovery.Resource
    """

    credentials = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, "rb") as token:
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
        with open(TOKEN_PICKLE_FILE, "wb") as token:
            pickle.dump(credentials, token)

    # Build and return the service
    apps_script_service = build("script", "v1", credentials=credentials,
                                cache_discovery=False)
    return apps_script_service


def send_email(config: configparser.SectionProxy, changes: list):
    # create email object
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(config["host"], config["port"], context=context)
    server.login(config["username"], config["password"])
    server.sendmail(config["from"], config["to"], "this is the message. This is some <b>bold text</b>")


def main():
    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    config = parser["ledger_checker"]
    expense365 = parser["eXpense365"]

    # Prepare the authentication
    data = expense365["email"] + ":" + expense365["password"]
    auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

    # Download the ledger, convert it, and upload it to Google Sheets
    pdf_filepath = download_pdf(auth=auth,
                                group_id=expense365["group_id"],
                                subgroup_id=expense365["subgroup_id"],
                                filename_prefix=config["filename_prefix"],
                                dir_name=config["dir_name"],
                                app_gui=None)
    xlsx_filepath = convert_to_xlsx(pdf_filepath=pdf_filepath,
                                    app_gui=None,
                                    dir_name=config["dir_name"])
    new_url, sheet_name = upload_ledger(dir_name=config["dir_name"],
                                        destination_sheet_id=config["destination_sheet_id"],
                                        xlsx_filepath=xlsx_filepath)

    apps_script_service = authorize()
    print("Connected to the Apps Script service.")

    print("Executing the Apps Script function...")
    response = apps_script_service.scripts().run(body={"function": config["function"],
                                                       "parameters": sheet_name},
                                                 scriptId=config["script_id"]).execute()
    print("Apps Script function complete.")
    changes = response["response"].get("result")

    send_email(config=parser["email"], changes=changes)


if __name__ == "__main__":
    main()
