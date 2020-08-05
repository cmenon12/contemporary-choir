from __future__ import print_function

import base64
import configparser
import os
import pickle
import time
import webbrowser
from datetime import datetime
from json import decoder

import requests
from appJar import gui
from dateutil import tz
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# 30 is used for the ledger, 31 for the balance
REPORT_ID = "30"

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
TOKEN_PICKLE_FILE = "ledger_fetcher_token.pickle"


def download_pdf(auth: str, group_id: str, subgroup_id: str,
                 filename_prefix: str, dir_name: str,
                 app_gui: gui) -> str:
    """Downloads the ledger from expense365.com.

    :param auth: the authentication header with the email and password
    :type auth: str
    :param group_id: the ID of the group to download the ledger
    :type group_id: str
    :param subgroup_id: the ID of the subgroup to download the ledger
    :type subgroup_id: str
    :param filename_prefix: what to prefix the default filename with
    :type filename_prefix: str
    :param dir_name: the default directory to save the PDF
    :type dir_name: str
    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    :returns: the filepath of the saved PDF
    :rtype: str
    """

    # Prepare the request
    url = "https://service.expense365.com/ws/rest/eXpense365/RequestDocument"
    headers = {
        "Host": "service.expense365.com:443",
        "User-Agent": "eXpense365|1.5.2|Google Pixel XL|Android|10|en_GB",
        "Authorization": auth,
        "Accept": "application/json",
        "If-Modified-Since": "Mon, 1 Oct 1990 05:00:00 GMT",
        "Content-Type": "text/plain;charset=UTF-8",
    }

    # Prepare the body
    data = ('{"ReportID":' + REPORT_ID +
            ',"UserGroupID":' + group_id +
            ',"SubGroupID":' + subgroup_id + '}')

    # Make the request and check it was successful
    print("Making the HTTP request to service.expense365.com...")
    response = requests.post(url=url, headers=headers, data=data)
    try:
        response.raise_for_status()
    except requests.exceptions.HTTPError as err:
        raise SystemExit(err)
    print("The request was successful with no HTTP errors.")

    # Parse the date and convert it to the local timezone
    date_string = datetime.strptime(response.headers["Date"],
                                    "%a, %d %b %Y %H:%M:%S %Z") \
        .replace(tzinfo=tz.tzutc()) \
        .astimezone(tz.tzlocal()) \
        .strftime("%d-%m-%Y at %H.%M.%S")

    # Prepare to save the file
    filename = filename_prefix + " " + date_string

    # Attempt to save the PDF (fails gracefully if the user cancels)
    try:
        print("Saving the PDF...")
        if app_gui is not None:
            pdf_filepath = app_gui.saveBox("Save ledger",
                                           dirName=dir_name,
                                           fileName=filename,
                                           fileExt=".pdf",
                                           fileTypes=[("PDF file", "*.pdf")],
                                           asFile=True).name
            app_gui.removeAllWidgets()
        else:
            pdf_filepath = dir_name + filename + ".pdf"
        with open(pdf_filepath, "wb") as pdf_file:
            pdf_file.write(response.content)

    except Exception as err:
        print("There was an error saving the PDF ledger!")
        raise SystemExit(err)

    # If successful then return the file path
    else:
        print("PDF ledger saved successfully at " + pdf_filepath)
        return pdf_filepath


def convert_to_xlsx(pdf_filepath: str, dir_name: str,
                    app_gui: gui) -> str:
    """Converts the PDF to an Excel file and saves it.

    :param pdf_filepath: the path to the PDF file to convert
    :type pdf_filepath: str
    :param dir_name: the default directory to save the PDF
    :type dir_name: str
    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    :return: the path to the downloaded XLSX spreadsheet
    :rtype: str
    """

    # Prepare for the request
    url = "https://www.pdftoexcel.com/upload.instant.php"
    user_agent = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
                  'AppleWebKit/537.36 (KHTML, like Gecko) ' +
                  'Chrome/84.0.4147.105 Safari/537.36')
    headers = {
        'Connection': 'keep-alive',
        'Accept': '*/*',
        'X-Requested-With': 'XMLHttpRequest',
        'DNT': '1',
        'User-Agent': user_agent,
        'Origin': 'https://www.pdftoexcel.com',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://www.pdftoexcel.com/',
        'Accept-Language': 'en-GB,en;q=0.9',
    }
    files = {'Filedata': open(pdf_filepath, 'rb')}

    # Create a session that we can add headers to
    session = requests.Session()
    session.headers.update(headers)

    # Make the request and check that it was successful
    print("Sending the PDF for conversion to pdftoexcel.com...")
    response = session.post(url=url, files=files)
    try:
        response.raise_for_status()
        job_id = response.json()["jobId"]
    except requests.exceptions.HTTPError as err:
        raise SystemExit(err)
    except decoder.JSONDecodeError as err:
        raise SystemExit(err)
    print("The request was successful with no HTTP errors.")
    print("The jobId is %s" % job_id)

    # Prepare to keep checking the status of the conversion
    download_url = ""
    url = "https://www.pdftoexcel.com/status"

    # Whilst it is still being converted
    while download_url == "":

        # Prepare and make the request
        print("Checking if the conversion is complete...")
        response = session.post(url=url, data={"jobId": job_id, "rand": "0"})

        # Check that the request was successful
        try:
            response.raise_for_status()
            download_url = response.json()["download_url"]
        except requests.exceptions.HTTPError as err:
            raise SystemExit(err)
        except decoder.JSONDecodeError as err:
            raise SystemExit(err)

        # Wait before checking again
        if download_url == "":
            time.sleep(1)

    # Prepare and make the request to download the file
    url = "https://www.pdftoexcel.com" + download_url
    print("Downloading the converted file...")
    response = session.get(url=url, params={'id': job_id})

    # Check that the request was successful
    try:
        response.raise_for_status()
    except requests.exceptions.HTTPError as err:
        raise SystemExit(err)

    # Attempt to save the spreadsheet (fails gracefully if the user cancels)
    try:
        print("Saving the spreadsheet...")
        filename = pdf_filepath.split("\\")[-1]
        filename = filename.replace(".pdf", ".xlsx")
        if app_gui is not None:
            xlsx_filepath = app_gui.saveBox("Save spreadsheet",
                                            dirName=dir_name,
                                            fileName=filename,
                                            fileExt=".xlsx",
                                            fileTypes=[("Office Open XML " +
                                                        "Workbook", "*.xlsx")],
                                            asFile=True).name
            app_gui.removeAllWidgets()
        else:
            xlsx_filepath = pdf_filepath.replace(".pdf", ".xlsx")
        with open(xlsx_filepath, "wb") as xlsx_file:
            xlsx_file.write(response.content)

    except Exception as err:
        print("There was an error saving the spreadsheet!")
        raise SystemExit(err)

    # If successful then return the file path
    else:
        print("Spreadsheet ledger saved successfully at " + xlsx_filepath)
        return xlsx_filepath


def upload_ledger(dir_name: str, destination_sheet_id: str,
                  app_gui: gui = None, xlsx_filepath: str = "") -> tuple:
    """Uploads the ledger to the specified Google Sheet.

    :param dir_name: the default directory to open the XLSX file from
    :type dir_name: str
    :param destination_sheet_id: the id of the sheet to upload it to
    :type destination_sheet_id: str
    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    :param xlsx_filepath: the filepath of the XLSX spreadsheet
    :type xlsx_filepath: str, optional
    :return: the URL of the destination spreadsheet and the sheet name
    :rtype: tuple
    """

    # Open the spreadsheet
    if xlsx_filepath == "":
        try:
            xlsx_filepath = app_gui.openBox(title="Open spreadsheet",
                                            dirName=dir_name,
                                            fileTypes=[("Office Open XML " +
                                                        "Workbook", "*.xlsx")],
                                            asFile=True).name
            app_gui.removeAllWidgets()

        except Exception as err:
            print("There was an error opening the spreadsheet.")
            raise SystemExit(err)

    # Authenticate and retrieve the required services
    drive, sheets = authorize()

    # Upload the ledger
    print("Uploading the ledger...")
    file_metadata = {"name": "The Latest Ledger",
                     "mimeType": "application/vnd.google-apps.spreadsheet"}
    xlsx_mimetype = ("application/vnd.openxmlformats-officedocument" +
                     ".spreadsheetml.sheet")
    media = MediaFileUpload(xlsx_filepath,
                            mimetype=xlsx_mimetype,
                            resumable=True)
    file = drive.files().create(body=file_metadata,
                                media_body=media,
                                fields="id").execute()
    latest_ledger_id = file.get("id")

    # Get the ID of the first sheet in the newly-uploaded spreadsheet
    print("Fetching the sheet ID...")
    response = sheets.spreadsheets().get(spreadsheetId=latest_ledger_id,
                                         ranges="A1:D4",
                                         includeGridData=False).execute()
    sheet_id = response["sheets"][0]["properties"]["sheetId"]

    # Copy the uploader ledger to the sheet with the macro
    print("Copying the sheet...")
    body = {"destinationSpreadsheetId": destination_sheet_id}
    response = sheets.spreadsheets().sheets() \
        .copyTo(spreadsheetId=latest_ledger_id,
                sheetId=sheet_id,
                body=body).execute()
    new_sheet_id = response["sheetId"]
    new_sheet_title = response["title"]

    # Delete the uploaded Google Sheet
    print("Deleting the uploaded ledger...")
    drive.files().delete(fileId=latest_ledger_id).execute()

    print("New ledger uploaded to the destination sheet successfully!")

    return ("https://docs.google.com/spreadsheets/d/" + destination_sheet_id +
            "/edit#gid=" + str(new_sheet_id)), new_sheet_title


def authorize() -> tuple:
    """Authorizes access to the user's Drive and Sheets.

    :return: the authenticated services
    :rtype: tuple
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

    # Build both services and return them as a tuple
    drive_service = build("drive", "v3", credentials=credentials,
                          cache_discovery=False)
    sheets_service = build("sheets", "v4", credentials=credentials,
                           cache_discovery=False)
    return drive_service, sheets_service


def main(app_gui: gui):
    """Runs the program.

    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    """

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    config = parser["ledger_fetcher"]
    expense365 = parser["eXpense365"]

    # Prepare the authentication
    data = expense365["email"] + ":" + expense365["password"]
    auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

    # Download the PDF, returning the file path
    pdf_filepath = download_pdf(auth=auth,
                                group_id=expense365["group_id"],
                                subgroup_id=expense365["subgroup_id"],
                                filename_prefix=config["filename_prefix"],
                                dir_name=config["dir_name"],
                                app_gui=app_gui)

    # Ask the user if they want to open it
    if app_gui.yesNoBox("Open PDF?",
                        "Do you want to open the ledger?") is True:
        # If so then open it in the prescribed browser
        open_path = "file://///" + pdf_filepath
        webbrowser.register("my-browser",
                            None,
                            webbrowser.BackgroundBrowser(config["browser_path"]))
        webbrowser.get(using="my-browser").open(open_path)

    # Ask the user if they want to convert it to an XLSX spreadsheet
    if app_gui.yesNoBox("Convert to XLSX?",
                        ("Do you want to convert the PDF ledger to an XLSX " +
                         "spreadsheet using pdftoexcel.com, and then upload " +
                         "it to %s?" % config["destination_sheet_name"])) is True:

        # If so then convert it and upload it
        xlsx_filepath = convert_to_xlsx(pdf_filepath=pdf_filepath,
                                        app_gui=app_gui,
                                        dir_name=config["dir_name"])
        new_url, sheet_name = upload_ledger(dir_name=config["dir_name"],
                                            destination_sheet_id=config["destination_sheet_id"],
                                            app_gui=app_gui, xlsx_filepath=xlsx_filepath)

        # Ask the user if they want to open the new ledger in Google Sheets
        if app_gui.yesNoBox("Open %s?" % config["destination_sheet_name"],
                            ("Do you want to open the uploaded ledger in " +
                             "Google Sheets?")) is True:
            # If so then open it in the prescribed browser
            open_path = new_url
            webbrowser.register("my-browser",
                                None,
                                webbrowser.BackgroundBrowser(config["browser_path"]))
            webbrowser.get(using="my-browser").open(open_path)


if __name__ == "__main__":
    # Create the GUI
    appjar_gui = gui(showIcon=False)
    appjar_gui.setOnTop()

    main(appjar_gui)