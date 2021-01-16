"""Regularly checks the ledger for any changes, and notifies the user.

This module is primarily designed to check the ledger on a regular basis
for any new changes, and notify the user by email of any changes found.

This module specifically contains functions to execute an Apps Script
function, process the output from that function, and create and send an
email to notify the user of any changes.

This module relies on ledger_fetcher.py. It also relies on two Apps
Scripts, ledger-comparison.gs and ledger-checker.gs, to neatly format
and compare the ledger to an older version. These should be created in
an Apps Script project linked to the Google Sheet that the ledger is
uploaded to.
"""

import base64
import configparser
import email.utils
import imaplib
import logging
import os
import pickle
import re
import smtplib
import socket
import ssl
import time
import traceback
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import googleapiclient
import html2text
import timeago
from babel.numbers import format_currency
from googleapiclient.discovery import build
from jinja2 import Environment, FileSystemLoader

# noinspection PyCompatibility
from exceptions import AppsScriptApiException
# noinspection PyUnresolvedReferences
from ledger_fetcher import download_pdf, convert_to_xlsx, upload_ledger, \
    authorize, update_pdf_ledger

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# This is the number of consecutive failed attempts that the program
# can make before it sends an error email
ATTEMPTS = [3, 8, 15]

# The name of the config file
CONFIG_FILENAME = "config.ini"


class LedgerCheckerSaveFile:
    """Represents the save file used to maintain persistence."""

    def __init__(self, save_data_filepath: str, log: logging):

        # Try to open the save file if it exists
        try:
            with open(save_data_filepath, "rb") as f:

                # Load it and set the attributes for self from it
                inst = pickle.load(f)
                for k in inst.__dict__.keys():
                    setattr(self, k, getattr(inst, k))

            if not isinstance(inst, LedgerCheckerSaveFile):
                raise FileNotFoundError("The save data file is invalid.")
            pass

        # If it fails then just initialise with blank values
        except FileNotFoundError:

            self.save_data_filepath = save_data_filepath
            self.stack_traces = []
            self.changes = None
            self.sheet_id = None
            self.timestamp = None

        self.log = log
        self.save_data()

    def save_data(self):
        """Save to the actual file."""

        with open(self.save_data_filepath, "wb") as save_file:
            pickle.dump(self, save_file)
            save_file.close()

    def new_check_success(self, changes: dict, sheet_id: int,
                          timestamp: datetime):
        """Runs when a check was successful.

        :param changes: the changes made to the ledger
        :type changes: dict
        :param sheet_id: the ID of the new sheet
        :type sheet_id: int
        :param timestamp: when the last check was run
        :type timestamp: datetime
        """

        self.changes = changes
        self.sheet_id = sheet_id
        self.timestamp = timestamp
        self.stack_traces.clear()
        self.save_data()
        self.log.debug("Successful check saved to %s",
                       self.save_data_filepath)

    def new_check_fail(self, stack_trace: str):
        """Runs when a check failed.

        :param stack_trace: the stack trace
        :type stack_trace: str
        """

        date = datetime.now().strftime("%A %d %B %Y AT %H:%M:%S")
        self.stack_traces.append("ERROR ON %s\n%s" % (date, stack_trace))
        self.save_data()
        self.log.debug("Failed check saved to %s",
                       self.save_data_filepath)

    def get_data(self) -> tuple:
        """Gets and returns the saved data."""

        return self.changes, self.sheet_id, self.timestamp

    def get_stack_traces(self) -> list:
        """Gets and returns the list of stack traces."""

        return self.stack_traces

    def get_timestamp(self) -> datetime:
        """Gets and returns the timestamp."""

        return self.timestamp


def prepare_email_body(changes: dict, sheet_url: str, pdf_url: str,
                       old_timestamp: datetime):
    """Prepares an email detailing the changes to the ledger.

    :param changes: the changes made to the ledger
    :type changes: dict
    :param sheet_url: the URL of the Google Sheet to link to
    :type sheet_url: str
    :param pdf_url: the URL of the PDF to link to
    :type pdf_url: str
    :param old_timestamp: when the previous check was made
    :type old_timestamp: datetime.datetime
    :return: the plain-text and html
    :rtype: tuple
    """

    # Remove the cost codes with no change
    LOGGER.info("Creating the email...")
    for name, value in changes["costCodes"].copy().items():
        if value["changeInBalance"] == 0:
            changes["costCodes"].pop(name)

    # Calculate a grand total change
    changes["grandTotal"]["changeInTotalBalance"] = 0
    for name, value in changes["costCodes"].items():
        changes["grandTotal"]["changeInTotalBalance"] += value["changeInBalance"]

    # Format all of the money values in each cost code
    cost_code_items = ["balance", "changeInBalance", "moneyIn", "moneyOut"]
    for name, value in changes["costCodes"].items():

        # For the cost code totals
        for item in cost_code_items:
            changes["costCodes"][name][item] = \
                format_currency(value[item], "GBP", locale="en_GB")

        # For each entry in the cost code
        for i in range(0, len(value["entries"])):
            changes["costCodes"][name]["entries"][i]["money"] = \
                format_currency(changes["costCodes"][name]["entries"][i]["money"],
                                "GBP", locale="en_GB")

    # Format the grand total values
    grand_total_items = ["balanceBroughtForward", "totalIn", "totalBalance",
                         "totalOut", "changeInTotalBalance"]
    for item in grand_total_items:
        changes["grandTotal"][item] = \
            format_currency(changes["grandTotal"][item], "GBP", locale="en_GB")

    # Prepare when the last check was made
    if old_timestamp is not None:
        last_check = " since the last check %s on %s" % (timeago.format(old_timestamp),
                                                         old_timestamp.strftime("%A %d %B %Y at %H:%M:%S"))
    else:
        last_check = ""

    # Render the template
    root = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader(root))
    template = env.get_template("email-template.html")
    html = template.render(changes=changes,
                           sheet_url=sheet_url,
                           pdf_url=pdf_url,
                           last_check=last_check)

    # Create the plain-text version of the message
    text_maker = html2text.HTML2Text()
    text_maker.ignore_links = True
    text_maker.bypass_tables = False
    text = text_maker.handle(html)
    LOGGER.info("Email HTML and plain-text created successfully.")

    return text, html


def send_success_email(config: configparser.SectionProxy, changes: dict,
                       pdf_filepath: str, sheet_url: str, pdf_url: str,
                       old_timestamp: datetime):
    """Sends an email detailing the changes.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param changes: the changes made to the ledger
    :type changes: dict
    :param pdf_filepath: the path to the PDF to attach
    :type pdf_filepath: str
    :param sheet_url: the URL of the Google Sheet to link to
    :type sheet_url: str
    :param pdf_url: the URL of the PDF to link to
    :type pdf_url: str
    :param old_timestamp: when the previous check was made
    :type old_timestamp: datetime.datetime
    """

    message = MIMEMultipart("alternative")
    message["Subject"] = changes["societyName"] + " Ledger Update"
    message["To"] = config["to"]
    message["From"] = config["from"]

    # Prepare the email
    text, html = prepare_email_body(changes=changes,
                                    sheet_url=sheet_url,
                                    pdf_url=pdf_url,
                                    old_timestamp=old_timestamp)

    # Turn these into plain/html MIMEText objects
    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(MIMEText(text, "plain"))
    message.attach(MIMEText(html, "html"))

    # Attach the ledger
    with open(pdf_filepath, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "pdf")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    head, filename = os.path.split(pdf_filepath)
    part.add_header("Content-Disposition",
                    "attachment; filename=\"%s\"" % filename)
    message.attach(part)

    # Send the email
    send_email(config=config, message=message)


def delete_sheet(sheets_service: googleapiclient.discovery.Resource,
                 spreadsheet_id: str, sheet_id: id):
    """Deletes the named Google Sheet in the specified spreadsheet.

    :param sheets_service: the authenticated service for Google Sheets
    :type sheets_service: googleapiclient.discovery.Resource
    :param spreadsheet_id: the ID of the spreadsheet to use
    :type spreadsheet_id: str
    :param sheet_id: the id of the sheet to delete
    :type sheet_id: int
    """

    if sheet_id is not None:
        LOGGER.info("Deleting the sheet with ID %d", sheet_id)
        body = {"requests": {"deleteSheet": {"sheetId": sheet_id}}}
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                  body=body).execute()
        LOGGER.info("Sheet %d has been deleted successfully.", sheet_id)
    else:
        LOGGER.info("sheet_id was None, so there was nothing to delete.")


def send_error_email(config: configparser.SectionProxy,
                     save_data: LedgerCheckerSaveFile):
    """Used to email the user about a fatal exception.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param save_data: the save data object
    :type save_data: LedgerCheckerSaveFile
    """

    # Get the count of errors and stack traces
    error_count = len(save_data.get_stack_traces())
    stack_traces = ""
    for e in save_data.get_stack_traces():
        stack_traces += e
        stack_traces += "\n\n"

    # Create the message
    message = MIMEMultipart("alternative")
    message["Subject"] = "ERROR with ledger_checker.py!"
    message["To"] = config["to"]
    message["From"] = config["from"]
    message["X-Priority"] = "1"

    # Prepare the email
    if error_count >= ATTEMPTS[-1]:
        future_attempts = " and no further error emails will be sent."
    else:
        future_attempts = "."
        for i in ATTEMPTS:
            if i > error_count:
                future_attempts = (" and another email will be sent if %d "
                                   "consecutive and fatal errors occur "
                                   "(including the %d below)."
                                   % (i, error_count))
                break

    text = ("There were %d consecutive and fatal errors with ledger_checker.py! "
            "Please see the stack traces below and check the log. Note that "
            "future executions will continue as scheduled%s\n\n %s "
            "———\nThis email was sent automatically by a "
            "computer program (https://github.com/cmenon12/contemporary-choir). "
            "If you want to leave some feedback "
            "then please reply directly to it." % (error_count,
                                                   future_attempts,
                                                   stack_traces))
    message.attach(MIMEText(text, "plain"))

    # Send the email
    send_email(config=config, message=message)


def send_email(config: configparser.SectionProxy, message: MIMEMultipart):
    """Sends the message using the config with SSL.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param message: the message to send
    :type message: MIMEMultipart
    """

    # Add a few headers to the message
    message["Date"] = email.utils.formatdate()
    message["Message-ID"] = email.utils.make_msgid(domain=config["smtp_host"])

    # Create the SMTP connection and send the email
    LOGGER.info("Connecting to the SMTP server to send the email...")
    with smtplib.SMTP_SSL(config["smtp_host"],
                          int(config["smtp_port"]),
                          context=ssl.create_default_context()) as server:
        server.login(config["username"], config["password"])
        LOGGER.info("Sending the email...")
        server.sendmail(re.findall("(?<=<)\\S+(?=>)", config["from"])[0],
                        re.findall("(?<=<)\\S+(?=>)", config["to"]),
                        message.as_string())
    LOGGER.info("Email sent successfully!")

    # If asked then manually save it to the Sent folder
    if config["save_to_sent"] == "True":
        LOGGER.info("Connecting to the IMAP server to save the email...")
        with imaplib.IMAP4_SSL(config["imap_host"],
                               int(config["imap_port"])) as server:
            server.login(config["username"], config["password"])
            server.append('INBOX.Sent', '\\Seen',
                          imaplib.Time2Internaldate(time.time()),
                          message.as_string().encode('utf8'))
        LOGGER.info("Email saved successfully!")


def check_ledger(save_data: LedgerCheckerSaveFile,
                 parser: configparser.ConfigParser):
    """Runs a check of the ledger.

    :param save_data: the save data object
    :type save_data: LedgerCheckerSaveFile
    :param parser: the whole config
    :type parser: configparser.ConfigParser
    :raises: AppsScriptApiException
    """

    LOGGER.info("\n")

    # Fetch info from the config
    config = parser["ledger_checker"]
    expense365 = parser["eXpense365"]

    # Prepare the authentication
    data = expense365["email"] + ":" + expense365["password"]
    auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

    # Download the ledger, convert it, and upload it to Google Sheets
    print("Downloading the PDF...")
    pdf_filepath = download_pdf(auth=auth,
                                group_id=expense365["group_id"],
                                subgroup_id=expense365["subgroup_id"],
                                filename_prefix=config["filename_prefix"],
                                dir_name=config["dir_name"],
                                app_gui=None)
    print("Converting the ledger...")
    xlsx_filepath = convert_to_xlsx(pdf_filepath=pdf_filepath,
                                    app_gui=None,
                                    dir_name=config["dir_name"])
    print("Uploading the ledger to Google Sheets...")
    sheet_url, sheet_name = upload_ledger(dir_name=config["dir_name"],
                                          destination_sheet_id=config["destination_sheet_id"],
                                          xlsx_filepath=xlsx_filepath)
    print("Ledger downloaded, converted, and uploaded successfully.")

    # Connect to the Apps Script service and attempt to execute it
    socket.setdefaulttimeout(600)
    drive, sheets, apps_script = authorize()
    print("Executing the Apps Script function (this may take some time)...")
    LOGGER.info("Starting the Apps Script function...")
    body = {"function": config["function"], "parameters": sheet_name}
    response = apps_script.scripts().run(body=body,
                                         scriptId=config["script_id"]).execute()

    # Catch and then raise an error during execution
    if "error" in response:
        LOGGER.error("There was an error with the Apps Script function!")
        LOGGER.error(response["error"])
        error = response["error"]["details"][0]
        print("Script error message: " + error["errorMessage"])
        if "scriptStackTraceElements" in error:
            # There may not be a stacktrace if the script didn't start
            # executing.
            print("Script error stacktrace:")
            for trace in error["scriptStackTraceElements"]:
                print("\t{0}: {1}".format(trace["function"],
                                          trace["lineNumber"]))
        raise AppsScriptApiException(response["error"])

    # Otherwise save the data that the Apps Script returns
    changes = response["response"].get("result")
    LOGGER.debug(changes)
    print("The Apps Script function executed successfully!")
    LOGGER.info("The Apps Script function executed successfully.")

    # Get the old values of changes
    old_changes, old_sheet_id, previous_check = save_data.get_data()

    # Get the datetime that the request was made for the save file
    head, filename = os.path.split(pdf_filepath)
    timestamp = datetime.strptime(filename,
                                  parser["ledger_checker"]["filename_prefix"] +
                                  " %d-%m-%Y at %H.%M.%S.pdf")

    # If there were no changes then do nothing
    # The sheet will have been deleted by the Apps Script
    if changes == "False":
        print("Changes is False, so we'll do nothing.")
        LOGGER.info("Changes is False, so we'll do nothing.")
        os.remove(pdf_filepath)
        os.remove(xlsx_filepath)
        LOGGER.info("The local PDF and XLSX have been deleted successfully.")
        save_data.new_check_success(changes=old_changes,
                                    sheet_id=old_sheet_id,
                                    timestamp=timestamp)

    # If the returned changes aren't actually new to us then
    # just delete the new sheet we just made
    # This compares the total income & expenditure
    elif old_changes is not None and \
            format_currency(changes["grandTotal"]["totalIn"],
                            "GBP", locale="en_GB") == old_changes["grandTotal"]["totalIn"] and \
            format_currency(changes["grandTotal"]["totalOut"],
                            "GBP", locale="en_GB") == old_changes["grandTotal"]["totalOut"]:
        print("The new changes is the same as the old.")
        LOGGER.info("The new changes is the same as the old.")
        LOGGER.info("Deleting the new sheet (that's the same as the old one)...")
        delete_sheet(sheets_service=sheets,
                     spreadsheet_id=config["destination_sheet_id"],
                     sheet_id=changes["sheetId"])
        os.remove(pdf_filepath)
        os.remove(xlsx_filepath)
        LOGGER.info("The local PDF and XLSX have been deleted successfully!")
        save_data.new_check_success(changes=old_changes,
                                    sheet_id=old_sheet_id,
                                    timestamp=timestamp)

    # Otherwise these changes are new
    # Update the PDF ledger in the user's Google Drive
    # Notify the user (via email) and delete the old sheet
    # Save the new data to the save file
    else:
        print("We have some new changes!")
        LOGGER.info("We have some new changes.")
        pdf_url = update_pdf_ledger(dir_name=config["dir_name"],
                                    pdf_ledger_id=config["pdf_ledger_id"],
                                    pdf_ledger_name=config["pdf_ledger_name"],
                                    pdf_filepath=pdf_filepath)
        old_timestamp = save_data.get_timestamp()
        send_success_email(config=parser["email"], changes=changes,
                           pdf_filepath=pdf_filepath,
                           sheet_url=sheet_url, pdf_url=pdf_url,
                           old_timestamp=old_timestamp)
        print("Email sent successfully!")
        LOGGER.info("Deleting the old sheet...")
        delete_sheet(sheets_service=sheets,
                     spreadsheet_id=config["destination_sheet_id"],
                     sheet_id=old_sheet_id)
        save_data.new_check_success(changes=changes,
                                    sheet_id=changes["sheetId"],
                                    timestamp=timestamp)


def main():
    """Manages the checks.

    This function runs the checks of the ledger regularly and gracefully
    handles any errors that occur.
    """

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read(CONFIG_FILENAME)

    save_data = LedgerCheckerSaveFile(parser["ledger_checker"]["save_data_filepath"], LOGGER)

    try:
        check_ledger(save_data=save_data, parser=parser)
    except Exception:
        save_data.new_check_fail(stack_trace=traceback.format_exc())
        traceback.print_exc()
        LOGGER.exception("That check went wrong!")
        LOGGER.error("This is consecutive failed attempt no. %d.",
                     len(save_data.get_stack_traces()))
        print("This is consecutive error number %d"
              % len(save_data.get_stack_traces()))

        if len(save_data.get_stack_traces()) in ATTEMPTS:
            send_error_email(config=parser["email"], save_data=save_data)
            print("Email sent successfully!")


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(filename="ledger_checker.log",
                        filemode="a",
                        format="%(asctime)s | %(levelname)s : %(message)s",
                        level=logging.DEBUG)
    LOGGER = logging.getLogger(__name__)

    main()

else:
    LOGGER = logging.getLogger(__name__)
