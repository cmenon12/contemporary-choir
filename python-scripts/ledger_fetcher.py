"""Facilitates fetching and processing the ledger from eXpense365.

This module is designed to download a society ledger from eXpense365
instead of having to use their app. You can then also convert it to an
XLSX spreadsheet, upload that to Google Sheets, and upload the PDF to
Google Drive.

The Ledger class represents the ledger. The ledger is automatically
downloaded and saved upon instantiation. It contains methods to convert
it to an XLSX spreadsheet and upload it to Google Sheets & Drive, as
well as to delete the locally saved files or Google Sheet, and refresh
itself to a newer version (by re-downloading the original ledger). It
also has a large range of getter methods which can invoke these
behaviours.

The PDFToXLSXConverter class represents a PDF to XLSX converter. Upon
instantiation a converter is chosen (either by choice or randomly). It
contains the methods to convert the ledger to an XLSX.

The CustomEncoder class is a custom JSON encoder. It overrides the default()
method of json.JSONEncoder. It will skip bytes and instead just return
a string that states "bytes object of length X not shown", and use the
default encoder for all other types (falling back on converting it to a
string if this fails). This is used by the other classes to save
themselves to the log.

The authorize() function is used to authorize access to Google Drive,
Google Sheets, and Google Apps Script.

When main() is run then the ledger is downloaded, and the user is
asked if they want to open it and/or upload it to Google Sheets. This
is done using dialog boxes so the user doesn't need to interact with it
at the command line.

Before running this script, make sure you've got a valid config file
and a valid set of credentials (saved to config.ini and
credentials.json respectively).

This script relies on the classes in custom_exceptions.py.
"""

from __future__ import print_function

import base64
import configparser
import json
import logging
import os
import pathlib
import pickle
import shutil
import time
import traceback
import webbrowser
from datetime import datetime
from typing import Optional, Union
from pyexcel.cookbook import merge_csv_to_a_book

import camelot
import pyperclip
import requests
from appJar import gui
from dateutil import tz
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from inputimeout import inputimeout, TimeoutOccurred
from pushbullet import InvalidKeyError, Pushbullet, PushError

from custom_exceptions import *

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# The name of the config file
CONFIG_FILENAME = "config.ini"

# This variable specifies the name of a file that contains the
# OAuth 2.0 information for this application, including its client_id
# and client_secret.
CLIENT_SECRETS_FILE = "credentials.json"

# This OAuth 2.0 access scope allows for full read/write access to the
# authenticated user's account and requires requests to use an SSL
# connection.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

# How long to wait for authorization (in seconds)
AUTHORIZATION_TIMEOUT = 300

# This file stores the user's access and refresh tokens and is created
# automatically when the authorization flow completes for the first
# time. This specifies where the file is stored
TOKEN_PICKLE_FILE = "token.pickle"


class CustomEncoder(json.JSONEncoder):
    """Represents a custom JSON encoder."""

    def default(self, obj):
        """Overrides the default encoder"""

        # If it's bytes then skip it
        if isinstance(obj, bytes):
            return "bytes object of length %d not shown" % len(obj)

        # Try to use the default encoder
        try:
            return json.JSONEncoder.default(self, obj)

        # Otherwise just convert it to a string
        except TypeError:
            return str(obj)


class Ledger:
    """Represents the ledger."""

    def __init__(self, config: configparser.SectionProxy,
                 expense365: configparser.SectionProxy,
                 app_gui: gui = None):
        """Constructs the ledger, including downloading the PDF.

        :param config: the general config
        :type: configparser.SectionProxy
        :param expense365: the expense365-specific config
        :type: configparser.SectionProxy
        :param app_gui: the appJar GUI to use
        :type app_gui: gui
        """

        self.app_gui = app_gui

        # Prepare the authentication
        data = expense365["email"] + ":" + expense365["password"]
        self.auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

        # Download the ledger
        self.expense365_data = {"ReportID": int(expense365["report_id"]),
                                "UserGroupID": int(expense365["group_id"]),
                                "SubGroupID": int(expense365["subgroup_id"])}
        self.filename_prefix = config["filename_prefix"]
        self.dir_name = config["dir_name"]
        self.pushbullet = {"access_token": config["pushbullet_access_token"],
                           "device": config["pushbullet_device"]}
        self.pdf_filepath = None
        self.pdf_filename = None
        self.pdf_file = None
        self.timestamp = None
        self.download_pdf()

        # Prepare for future use
        self.pdf_ledger_id = config["pdf_ledger_id"]
        self.pdf_ledger_name = config["pdf_ledger_name"]
        self.destination_sheet_name = config["destination_sheet_name"]
        self.destination_sheet_id = config["destination_sheet_id"]
        if str(config["browser_path"]).lower() == "false":
            self.browser_path = False
        else:
            self.browser_path = config["browser_path"]
        self.drive_pdf_url = None
        self.xlsx_filepath = None
        self.xlsx_filename = None
        self.xlsx_file = None
        self.sheets_data = None

        self.log()

    def download_pdf(self, save: bool = True) -> None:
        """Downloads the ledger from expense365.com.

        :param save: whether to save the PDF ledger
        :type save: bool, optional
        :raises HTTPError: if an unsuccessful HTTP status code is returned
        """

        # Prepare the request
        url = "https://service.expense365.com/ws/rest/eXpense365/RequestDocument"
        headers = {
            "Host": "service.expense365.com:443",
            "User-Agent": "eXpense365|1.6.1|Google Pixel XL|Android|10|en_GB",
            "Authorization": self.auth,
            "Accept": "application/json",
            "If-Modified-Since": "Mon, 1 Oct 1990 05:00:00 GMT",
            "Content-Type": "text/plain;charset=UTF-8",
        }

        # Make the request and check it was successful
        LOGGER.info("Making the HTTP request to service.expense365.com...")
        response = requests.post(url=url, headers=headers,
                                 data=json.dumps(self.expense365_data))
        response.raise_for_status()
        LOGGER.info("The request was successful with no HTTP errors.")

        # Save the date and convert it to the local timezone
        self.timestamp = datetime.strptime(response.headers["Date"],
                                           "%a, %d %b %Y %H:%M:%S %Z") \
            .replace(tzinfo=tz.tzutc()) \
            .astimezone(tz.tzlocal())

        # Parse the date as a string
        date_string = self.get_timestamp().strftime("%d-%m-%Y at %H.%M.%S")

        # Save the file
        self.pdf_filename = self.filename_prefix + " " + date_string + ".pdf"
        self.pdf_file = response.content
        if save is True:
            self.save_pdf()

    def convert_to_xlsx(self) -> None:
        """Converts the PDF to an Excel file and saves it."""

        # Convert the PDF
        tables = camelot.read_pdf(self.get_pdf_filepath(),
                                  columns=["88.2,430.4,504.6"],
                                  table_areas=["27.4,810,568.6,30.2"],
                                  split_text=True,
                                  row_tol=15,
                                  flavor="stream",
                                  pages="1-end")

        # Make a temporary folder
        path = pathlib.Path(self.dir_name, "temp-camelot")
        if os.path.exists(path):
            shutil.rmtree(path)
        path.mkdir(parents=True)

        # Save the PDF as a set of CSV files
        filepath = os.path.join(path, "result.csv")
        LOGGER.info("Saving the CSV files to %s", filepath)
        tables.export(filepath, f="csv")

        # Combine the CSV files into one
        with open(os.path.join(path, "combined.csv"), "w") as combined_file:
            table = 1
            while True:
                filename = "result-page-%d-table-1.csv" % table
                if filename in os.listdir(path):
                    with open(os.path.join(path, filename)) as table_file:
                        combined_file.write(table_file.read())
                        LOGGER.info("Wrote %s to %s.", os.path.join(path, filename),
                                    os.path.join(path, "combined.csv"))
                else:
                    break
                table += 1

        # Convert the CSV to an XLSX and save it
        self.xlsx_filename = self.get_pdf_filename().replace(".pdf", ".xlsx")
        self.xlsx_filepath = self.dir_name + self.xlsx_filename
        merge_csv_to_a_book([os.path.join(path, "combined.csv")],
                            self.xlsx_filepath)
        with open(self.xlsx_filepath, "rb") as xlsx_file:
            self.xlsx_file = xlsx_file.read()
        LOGGER.info("Saved the XLSX to %s", self.xlsx_filepath)

        # Delete the temporary folder and all files
        shutil.rmtree(path)

    def update_drive_pdf(self, save: bool = True) -> None:
        """Updates the PDF in Drive with the new version of the ledger.

        :param save: whether to save the PDF ledger
        :type save: bool, optional
        """

        LOGGER.info("PDF at %s has been opened.",
                    self.get_pdf_filepath(save=save))

        # Authenticate and retrieve the required services
        drive, sheets, apps_script = authorize(pushbullet=self.pushbullet,
                                               open_browser=self.browser_path)

        # Update the PDF copy of the ledger with a new version
        LOGGER.info("Uploading the new PDF ledger to Drive...")
        file_metadata = {"name": self.pdf_ledger_name,
                         "mimeType": "application/pdf",
                         "originalFilename": self.get_pdf_filename()}
        media = MediaFileUpload(self.get_pdf_filepath(save=save),
                                mimetype="application/pdf",
                                resumable=True)
        file = drive.files().update(body=file_metadata,
                                    media_body=media,
                                    fields="webViewLink",
                                    fileId=self.pdf_ledger_id,
                                    keepRevisionForever=True).execute()
        pdf_url = file.get("webViewLink")
        LOGGER.info("PDF Ledger uploaded to Google Drive at %s.", pdf_url)

        self.drive_pdf_url = pdf_url

    def upload_to_sheets(self, convert: bool = True) -> None:
        """Uploads the ledger to the specified Google Sheet.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        """

        LOGGER.info("Spreadsheet at %s has been opened.",
                    self.get_xlsx_filepath(convert=convert))

        # Authenticate and retrieve the required services
        drive, sheets, apps_script = authorize(pushbullet=self.pushbullet,
                                               open_browser=self.browser_path)

        # Upload the ledger
        LOGGER.info("Uploading the ledger to Google Sheets")
        file_metadata = {"name": "The Latest Ledger (temporary)",
                         "mimeType": "application/vnd.google-apps.spreadsheet"}
        xlsx_mimetype = ("application/vnd.openxmlformats-officedocument" +
                         ".spreadsheetml.sheet")
        media = MediaFileUpload(self.get_xlsx_filepath(convert=convert),
                                mimetype=xlsx_mimetype,
                                resumable=True)
        file = drive.files().create(body=file_metadata,
                                    media_body=media,
                                    fields="id").execute()
        latest_ledger_id = file.get("id")
        LOGGER.info("Ledger uploaded to Google Sheets with file ID %s.",
                    latest_ledger_id)

        # Get the ID of the first sheet in the newly-uploaded spreadsheet
        LOGGER.info("Fetching the sheet ID...")
        response = sheets.spreadsheets().get(spreadsheetId=latest_ledger_id,
                                             ranges="A1:D4",
                                             includeGridData=False).execute()
        sheet_id = response["sheets"][0]["properties"]["sheetId"]
        LOGGER.info("The ledger is in the sheet with ID %s.", sheet_id)

        # Copy the uploader ledger to the sheet with the macro
        LOGGER.info("Copying the sheet to the spreadsheet with ID %s...",
                    self.destination_sheet_id)
        body = {"destinationSpreadsheetId": self.destination_sheet_id}
        response = sheets.spreadsheets().sheets() \
            .copyTo(spreadsheetId=latest_ledger_id,
                    sheetId=sheet_id,
                    body=body).execute()
        new_sheet_id = response["sheetId"]
        LOGGER.info("Sheet copied successfully. Sheet has ID %s.",
                    new_sheet_id)

        # Rename the copied sheet
        new_sheet_title = self.get_timestamp().strftime("%d/%m/%Y %H:%M:%S")
        body = {"requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": int(new_sheet_id),
                               "title": new_sheet_title},
                "fields": "title"
            }
        }], "includeSpreadsheetInResponse": False}
        sheets.spreadsheets().batchUpdate(spreadsheetId=self.destination_sheet_id,
                                          body=body).execute()
        LOGGER.info("Sheet renamed successfully. Sheet has ID %s and title %s.",
                    new_sheet_id, new_sheet_title)

        # Delete the uploaded Google Sheet
        LOGGER.info("Deleting the uploaded ledger...")
        drive.files().delete(fileId=latest_ledger_id).execute()
        LOGGER.info("Original uploaded ledger deleted successfully.")

        self.sheets_data = {"name": new_sheet_title,
                            "spreadsheet_id": self.destination_sheet_id,
                            "sheet_id": str(new_sheet_id),
                            "url": ("https://docs.google.com/spreadsheets/d/" +
                                    self.destination_sheet_id +
                                    "/edit#gid=" + str(new_sheet_id))}

    def get_timestamp(self) -> datetime:
        """Get the timestamp of the ledger.

        :return: the timestamp of the ledger.
        :rtype: datetime
        """

        return self.timestamp

    def get_pdf_filepath(self, save: bool = True) -> str:
        """Returns the filepath of the PDF ledger.

        :param save: whether to save the PDF ledger
        :type save: bool, optional
        :return: the filepath of the PDF ledger.
        :rtype: str
        :raises PDFIsNotSavedError: if the PDF file isn't saved
        """

        if self.pdf_filepath is None and save is True:
            self.save_pdf()
        elif self.pdf_filepath is None and save is False:
            raise PDFIsNotSavedError("The PDF ledger isn't saved.")
        return self.pdf_filepath

    def get_pdf_filename(self) -> str:
        """Returns the filename of the PDF.

        :return: the filename of the PDF ledger.
        :rtype: str
        """

        return self.pdf_filename

    def get_pdf_file(self) -> bytes:
        """Returns the PDF ledger as a string of bytes.

        :return: the PDF file itself.
        :rtype: bytes
        """

        return self.pdf_file

    def get_xlsx_filepath(self, convert: bool = True) -> str:
        """Returns the filepath of the XLSX spreadsheet.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :return: the filepath of the XLSX ledger
        :rtype: str
        :raises XLSXDoesNotExistError: when convert is False
        """

        if self.xlsx_filepath is None:
            self.save_xlsx(convert=convert)

        return self.xlsx_filepath

    def get_xlsx_filename(self, convert: bool = True) -> str:
        """Returns the filename of the XLSX spreadsheet.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :return: the filename of the XLSX ledger
        :rtype: str
        :raises XLSXDoesNotExistError: when convert is False
        """

        if self.xlsx_filename is None and convert is True:
            self.convert_to_xlsx()
        elif self.xlsx_filename is None and convert is False:
            raise XLSXDoesNotExistError("The XLSX ledger doesn't exist.")
        return self.xlsx_filename

    def get_xlsx_file(self, convert: bool = True) -> bytes:
        """Returns the XLSX spreadsheet as a string of bytes.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :return: the XLSX file itself
        :rtype: bytes
        :raises XLSXDoesNotExistError: when convert is False
        """

        if self.xlsx_filepath is None and convert is True:
            self.convert_to_xlsx()
        elif self.xlsx_filepath is None and convert is False:
            raise XLSXDoesNotExistError("The XLSX ledger doesn't exist.")
        return self.xlsx_file

    def get_drive_pdf_url(self, save: bool = True, upload: bool = True) -> str:
        """Returns the URL of the PDF ledger in Drive.

        :param save: whether to save the XLSX ledger
        :type save: bool, optional
        :param upload: whether to upload the PDF if needed
        :type upload: bool, optional
        :return: the URL of the PDF in Drive
        :rtype: str
        :raises URLDoesNotExistError: when upload is False
        """

        if self.drive_pdf_url is None and upload is True:
            self.update_drive_pdf(save=save)
        elif self.drive_pdf_url is None and upload is False:
            raise URLDoesNotExistError("The PDF ledger isn't in Drive.")
        return self.drive_pdf_url

    def get_sheets_data(self, convert: bool = True,
                        upload: bool = True) -> dict:
        """Returns the name and URL of the ledger in Google Sheets.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :param upload: whether to upload the XLSX if needed
        :type upload: bool, optional
        :returns: the sheet name and url
        :rtype: dict
        :raises XLSXDoesNotExistError: when convert is False
        :raises URLDoesNotExistError: when upload is False
        """

        if self.sheets_data is None and upload is True:
            self.upload_to_sheets(convert=convert)
        elif self.sheets_data is None and upload is False:
            raise URLDoesNotExistError("The XLSX ledger isn't in Sheets.")
        return self.sheets_data

    def refresh_ledger(self):
        """Fetches a fresh copy of the ledger and invalidates everything."""

        self.download_pdf()
        self.drive_pdf_url = None
        self.xlsx_filepath = None
        self.xlsx_filename = None
        self.xlsx_file = None
        self.sheets_data = None
        self.log()

    def open_pdf_in_browser(self, save: bool = True) -> None:
        """Opens the PDF ledger in the designated browser.

        :param save: whether to save the PDF ledger
        :type save: bool, optional
        """

        open_path = "file://///" + self.get_pdf_filepath(save=save)
        if self.browser_path is False:
            print("View the PDF ledger here: %s" % open_path)
        else:
            webbrowser.register("my-browser",
                                None,
                                webbrowser.BackgroundBrowser(self.browser_path))
            webbrowser.get(using="my-browser").open(open_path)

    def open_sheet_in_browser(self, convert: bool = True,
                              upload: bool = True) -> None:
        """Opens the ledger in Google Sheets in the designated browser.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :param upload: whether to upload the XLSX if needed
        :type upload: bool, optional
        """

        open_path = self.get_sheets_data(convert=convert, upload=upload)["url"]
        if self.browser_path is False:
            print("Visit the Google Sheet here: %s" % open_path)
        else:
            webbrowser.register("my-browser",
                                None,
                                webbrowser.BackgroundBrowser(self.browser_path))
            webbrowser.get(using="my-browser").open(open_path)

    def save_pdf(self, use_gui: bool = True) -> None:
        """Save the PDF ledger to the file.

        :param use_gui: whether to use the GUI if app_gui is True
        :type use_gui: bool, optional
        """

        LOGGER.info("Saving the PDF...")

        # Find out where the user wants to save the PDF
        filename = self.get_pdf_filename().replace(".pdf", "")
        if self.app_gui is not None and use_gui is True:
            pdf_file_box = self.app_gui.saveBox("Save ledger",
                                                dirName=self.dir_name,
                                                fileName=filename,
                                                fileExt=".pdf",
                                                fileTypes=[("PDF file", "*.pdf")],
                                                asFile=True)
            if pdf_file_box is None:
                LOGGER.warning("The user cancelled saving the PDF.")
                raise SystemExit("User cancelled saving the PDF.")

            # Update the attributes to the new location
            self.pdf_filepath = pdf_file_box.name
            head, filename = os.path.split(self.pdf_filepath)
            self.pdf_filename = filename
            self.app_gui.removeAllWidgets()

        # Otherwise just use the default location
        else:
            self.pdf_filepath = self.dir_name + self.get_pdf_filename()

        # Save it
        with open(self.pdf_filepath, "wb") as pdf_file:
            pdf_file.write(self.pdf_file)
        LOGGER.info("PDF saved to %s successfully", self.pdf_filepath)

    def save_xlsx(self, convert: bool = True, use_gui: bool = True) -> None:
        """Save the XLSX ledger to the file.

        :param convert: whether to convert the PDF if needed
        :type convert: bool, optional
        :param use_gui: whether to use the GUI if app_gui is True
        :type use_gui: bool, optional
        """

        LOGGER.info("Saving the XLSX...")

        # Find out where the user wants to save the XLSX
        filename = self.get_xlsx_filename(convert=convert).replace(".xlsx", "")
        if self.app_gui is not None and use_gui is True:
            xlsx_file_box = self.app_gui.saveBox("Save ledger",
                                                 dirName=self.dir_name,
                                                 fileName=filename,
                                                 fileExt=".xlsx",
                                                 fileTypes=[("XLSX Spreadsheet", "*.xlsx")],
                                                 asFile=True)
            if xlsx_file_box is None:
                LOGGER.warning("The user cancelled saving the XLSX.")
                raise SystemExit("User cancelled saving the XLSX.")

            # Update the attributes to the new location
            self.xlsx_filepath = xlsx_file_box.name
            head, filename = os.path.split(self.xlsx_filepath)
            self.xlsx_filename = filename
            self.app_gui.removeAllWidgets()

        # Otherwise just use the default location
        else:
            self.xlsx_filepath = self.dir_name + self.get_xlsx_filename(convert=convert)

        # Save it
        with open(self.xlsx_filepath, "wb") as xlsx_file:
            xlsx_file.write(self.get_xlsx_file(convert=convert))
        LOGGER.info("XLSX saved to %s successfully", self.xlsx_filepath)

    def delete_pdf(self) -> None:
        """Delete the PDF file and remove the filepath."""

        if self.pdf_filepath is not None:
            os.remove(self.pdf_filepath)
            LOGGER.info("Deleted %s", self.pdf_filepath)
            self.pdf_filepath = None

    def delete_xlsx(self):
        """Delete the XLSX file and remove the filepath."""

        if self.xlsx_filepath is not None:
            os.remove(self.get_xlsx_filepath())
            LOGGER.info("Deleted %s", self.xlsx_filepath)
            self.xlsx_filepath = None

    def delete_sheet(self) -> None:
        """Deletes the uploaded Google Sheet, if it exists."""

        if self.sheets_data is not None:

            # Authenticate and retrieve the required services
            drive, sheets, apps_script = authorize(pushbullet=self.pushbullet,
                                                   open_browser=self.browser_path)

            # Delete the sheet
            body = {"requests": {"deleteSheet": {"sheetId": self.sheets_data["sheet_id"]}}}
            sheets.spreadsheets().batchUpdate(spreadsheetId=self.sheets_data["spreadsheet_id"],
                                              body=body).execute()
            LOGGER.info("Sheet %s has been deleted successfully.", self.sheets_data["sheet_id"])
            self.sheets_data = None
        else:
            LOGGER.info("There is no Google Sheet to delete.")

    def hide_sheet(self, hide: bool = True) -> None:
        """Hides (or un-hides) the uploaded Google Sheet, if it exists.

        :param hide: True to hide, False to un-hide
        :type hide: bool, optional
        """

        if self.sheets_data is not None:

            # Authenticate and retrieve the required services
            drive, sheets, apps_script = authorize(pushbullet=self.pushbullet,
                                                   open_browser=self.browser_path)

            # Hide the sheet
            body = {"requests": [{
                "updateSheetProperties": {
                    "properties": {"sheetId": int(self.sheets_data["sheet_id"]),
                                   "hidden": hide},
                    "fields": "hidden"
                }
            }], "includeSpreadsheetInResponse": False}
            sheets.spreadsheets().batchUpdate(spreadsheetId=self.sheets_data["spreadsheet_id"],
                                              body=body).execute()
            if hide:
                LOGGER.info("Sheet %s has been hidden successfully.", self.sheets_data["sheet_id"])
            else:
                LOGGER.info("Sheet %s has been un-hidden successfully.", self.sheets_data["sheet_id"])
        else:
            LOGGER.info("There is no Google Sheet to hide.")

    def log(self) -> None:
        """Logs the object to the log."""

        LOGGER.info(json.dumps(self.__dict__, cls=CustomEncoder))


def authorize(pushbullet: dict,
              open_browser: Optional[Union[bool, str]] = True) -> tuple:
    """Authorizes access to the user's Drive, Sheets, and Apps Script.

    :param pushbullet: the configuration for Pushbullet
    :type pushbullet: dict
    :param open_browser: whether to open the browser
    :type open_browser: Optional[Union[bool, str]]
    :return: the authenticated services
    :rtype: tuple
    """

    def authorize_in_browser():
        """Authorize in the browser, with a timeout."""

        if open_browser is False:
            # Tell the user to go and authorize it themselves
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, SCOPES,
                redirect_uri="urn:ietf:wg:oauth:2.0:oob")
            auth_url, _ = flow.authorization_url(prompt="consent")
            print("Please visit this URL to authorize this application: %s" % auth_url)
            push_url("Google Authorization URL", auth_url, pushbullet)
            pyperclip.copy(auth_url)
            print("The URL has been copied to the clipboard.")

            # Start the timer
            try:
                code = inputimeout(prompt="Enter the authorization code: ",
                                   timeout=AUTHORIZATION_TIMEOUT)
                flow.fetch_token(code=code)
                return flow.credentials
            except TimeoutOccurred:
                raise TimeoutError("Waited %d seconds to authorize Google APIs." %
                                   AUTHORIZATION_TIMEOUT)

        else:
            # Run the timer
            try:
                inputimeout(prompt="Press enter to open your browser: ",
                            timeout=AUTHORIZATION_TIMEOUT)
            except TimeoutOccurred:
                raise TimeoutError("Waited %d seconds to authorize Google APIs." %
                                   AUTHORIZATION_TIMEOUT)

            # Open the browser for the user to authorize it
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, SCOPES)
            print("Your browser should open automatically.")
            return flow.run_local_server(port=0)

    LOGGER.info("Authenticating the user to access Google APIs...")
    credentials = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, "rb") as token:
            credentials = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        LOGGER.info("There are no credentials or they are invalid.")
        if credentials and credentials.refresh_token:
            try:
                credentials.refresh(Request())
            except RefreshError:
                os.remove(TOKEN_PICKLE_FILE)
                credentials = authorize_in_browser()

        else:
            credentials = authorize_in_browser()

    # If we do have valid credentials then refresh them
    else:
        credentials.refresh(Request())

    # Save the credentials for the next run
    with open(TOKEN_PICKLE_FILE, "wb") as token:
        pickle.dump(credentials, token)
    LOGGER.info("Credentials saved to %s successfully.",
                TOKEN_PICKLE_FILE)

    # Build the services and return them as a tuple
    drive_service = build("drive", "v3", credentials=credentials,
                          cache_discovery=False)
    sheets_service = build("sheets", "v4", credentials=credentials,
                           cache_discovery=False)
    apps_script_service = build("script", "v1", credentials=credentials,
                                cache_discovery=False)
    LOGGER.info("Services built successfully.")
    return drive_service, sheets_service, apps_script_service


def push_url(title: str, url: str, config: dict):
    """Pushes the URL to Pushbullet using the config.

    This pushes the given URL to Pushbullet using the config. It will
    catch any Pushbullet-associated exceptions.

    :param title: the title of the URL
    :type title: str
    :param url: the URL to push
    :type url: str
    :param config: the configuration
    :type config: dict
    """

    # Stop if the user doesn't want to use Pushbullet
    if str(config["access_token"]).lower() == "false":
        return

    # Attempt to authenticate
    try:
        LOGGER.info("Authenticating with Pushbullet")
        pb = Pushbullet(config["access_token"])
        LOGGER.info("Authenticated with Pushbullet.")
    except InvalidKeyError:
        LOGGER.exception("InvalidKeyError raised when authenticating Pushbullet.")
        print("InvalidKeyError raised when authenticating Pushbullet.")
        traceback.print_exc()

    # If successfully authenticated then attempt to push
    else:
        try:
            if str(config["device"]).lower() == "false":
                pb.get_device(str(config["device"])).push_link(title, url)
                LOGGER.info("Pushed URL %s with title %s to all devices." %
                            (url, title))
                print("The URL has been successfully pushed to all devices.")
            else:
                pb.push_link(title, url)
                LOGGER.info("Pushed URL %s with title %s to device %s." %
                            (url, title, config["device"]))
                print("The URL has been successfully pushed to %s." %
                      config["device"])
        except InvalidKeyError:
            LOGGER.exception("InvalidKeyError raised when pushing to Pushbullet.")
            print("InvalidKeyError raised when pushing to Pushbullet.")
            traceback.print_exc()
        except PushError:
            LOGGER.exception("PushError raised when pushing to Pushbullet.")
            print("PushError raised when pushing to Pushbullet.")
            traceback.print_exc()


def main(app_gui: gui) -> None:
    """Runs the program.

    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    """

    # Check that the config file exists
    try:
        open(CONFIG_FILENAME)
        LOGGER.info("Loaded config %s.", CONFIG_FILENAME)
    except FileNotFoundError:
        print("The config file doesn't exist!")
        LOGGER.info("Could not find config %s, exiting.", CONFIG_FILENAME)
        time.sleep(5)
        raise FileNotFoundError("The config file doesn't exist!")

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read(CONFIG_FILENAME)
    config = parser["ledger_fetcher"]
    expense365 = parser["eXpense365"]

    # Get the ledger
    print("Downloading the PDF...")
    ledger = Ledger(config=config, expense365=expense365, app_gui=app_gui)

    # Ask the user if they want to open it
    if app_gui.yesNoBox("Open PDF?",
                        "Do you want to open the ledger?") is True:
        ledger.open_pdf_in_browser()

    # Ask the user if they want to convert it to an XLSX spreadsheet
    if app_gui.yesNoBox("Convert to XLSX?",
                        ("Do you want to convert the PDF ledger to an XLSX " +
                         "spreadsheet, and then upload it to %s?" %
                         config["destination_sheet_name"])) is True:

        # If so then convert it and upload it
        LOGGER.info("User chose to convert and upload the ledger.")
        print("Converting the ledger...")
        ledger.convert_to_xlsx()
        print("Uploading the ledger to Google Sheets...")
        sheets_data = ledger.get_sheets_data()
        print("Ledger uploaded to Google Sheets successfully. "
              "Find it in the sheet named %s." % sheets_data["name"])

        # Ask the user if they want to open the new ledger in Google Sheets
        if app_gui.yesNoBox("Open %s?" % config["destination_sheet_name"],
                            ("Do you want to open the uploaded ledger in " +
                             "Google Sheets?")) is True:
            # If so then open it in the prescribed browser
            LOGGER.info("User chose to open the uploaded ledger in Sheets.")
            ledger.open_sheet_in_browser()

    time.sleep(5)


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(filename="ledger_fetcher.log",
                        filemode="a",
                        format="%(asctime)s | %(levelname)s : %(message)s",
                        level=logging.DEBUG)
    LOGGER = logging.getLogger(__name__)

    # Create the GUI
    appjar_gui = gui(showIcon=False)
    appjar_gui.setOnTop()
    appjar_gui.setFont(size=12)

    main(appjar_gui)

else:
    LOGGER = logging.getLogger(__name__)
