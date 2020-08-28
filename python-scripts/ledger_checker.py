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
import logging
import os
import pickle
import re
import smtplib
import socket
import time
import traceback
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import googleapiclient
import html2text
from babel.numbers import format_currency
from googleapiclient.discovery import build
from jinja2 import Environment, FileSystemLoader
from ledger_fetcher import download_pdf, convert_to_xlsx, upload_ledger, authorize

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# This is the time in seconds between each check. Note that the actual
# time between checks may be sightly longer.
# 1800 is 30 minutes
INTERVAL = 1800

# This is the number of consecutive failed attempts that the program
# can make before it stops trying to run the checks.
ATTEMPTS = 3


def prepare_email_body(config: configparser.SectionProxy,
                       changes: list, url: str):
    """Prepares an email detailing the changes to the ledger.

        :param config: the configuration for the email
        :type config: configparser.SectionProxy
        :param changes: the changes made to the ledger
        :type changes: list
        :param url: the URL to link to
        :type url: str
        :return: the plain-text and html
        :rtype: tuple
        """

    # Calculate the total change for each cost code
    LOGGER.info("Creating the email...")
    cost_code_totals = {}
    for item in changes:
        # If the item is the totals for each cost code then skip it
        if isinstance(item[0], list):
            continue
        # Add new cost code (key)
        if item[0] not in cost_code_totals.keys():
            cost_code_totals[item[0]] = 0
        # Save the income and +ve and expenditure as -ve
        if item[3] != "":
            cost_code_totals[item[0]] += item[3]
        else:
            cost_code_totals[item[0]] -= item[4]

    # Calculate a grand total change
    grand_total = 0
    for key, value in cost_code_totals.items():
        grand_total += value
        value = format_currency(int(value), 'GBP', locale='en_GB')
        cost_code_totals[key] = value

    # Process each entry for each cost code
    cost_codes = {}
    for item in changes:
        # If the item is the totals for each cost code then skip it
        if isinstance(item[0], list):
            continue
        # Save the income and +ve and expenditure as -ve
        if item[3] == "":
            item[3] = -int(item[4])
        item[3] = format_currency(int(item[3]), 'GBP', locale='en_GB')
        # Add new cost code (key)
        if item[0] not in cost_codes.keys():
            cost_codes[item[0]] = []
        # Add the formatted entry to its cost code
        cost_codes[item[0]].append(item[1:4])

    # cost_codes is structured as
    # {"cost code name": [["date", "description", "£amount"],
    #                     ["date", "description", "£amount"],
    #                     ["date", "description", "£amount"]]}

    # Add the current balance for each cost code to cost_code_totals
    for key, value in cost_code_totals.items():
        total = [value]
        for item in changes[-1]:
            if item[0] == key:
                total.append(format_currency(int(item[3]),
                                             'GBP', locale='en_GB'))
        cost_code_totals[key] = total
    # Add the total changes and total balance
    cost_code_totals["grand_total"] = [format_currency(int(grand_total),
                                                       'GBP', locale='en_GB'),
                                       format_currency(int(changes[-1][-1][3]),
                                                       'GBP', locale='en_GB')]

    # cost_code_totals is structured as
    # {"cost code name": ["£change", "£balance"],
    #  "grand_total": ["£total change", "£total balance"]}

    # Render the template
    root = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader(root))
    template = env.get_template('email-template.html')
    html = template.render(society_name=config["society_name"],
                           cost_codes=cost_codes,
                           cost_code_totals=cost_code_totals,
                           url=url)

    # Create the plain-text version of the message
    text_maker = html2text.HTML2Text()
    text_maker.ignore_links = True
    text_maker.bypass_tables = False
    text = text_maker.handle(html)
    LOGGER.info("Email HTML and plain-text created successfully.")

    return text, html


def send_email(config: configparser.SectionProxy, changes: list,
               pdf_filepath: str, url: str):
    """Sends an email detailing the changes.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param changes: the changes made to the ledger
    :type changes: list
    :param pdf_filepath: the path to the PDF to attach
    :type pdf_filepath: str
    :param url: the URL to link to
    :type url: str
    """

    # Create the SMTP connection
    LOGGER.info("Connecting to the email server...")
    with smtplib.SMTP(config["host"], int(config["port"])) as server:
        server.starttls()
        server.login(config["username"], config["password"])
        message = MIMEMultipart("alternative")
        message["Subject"] = config["society_name"] + " Ledger Update"
        message["To"] = config["to"]
        message["From"] = config["from"]

        # Prepare the email
        text, html = prepare_email_body(config=config,
                                        changes=changes,
                                        url=url)

        # Turn these into plain/html MIMEText objects
        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(MIMEText(text, "plain"))
        message.attach(MIMEText(html, "html"))

        # Attach the ledger
        LOGGER.info("Attaching the PDF ledger...")
        with open(pdf_filepath, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        # Add header as key/value pair to attachment part
        head, filename = os.path.split(pdf_filepath)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add the attachment to message and send the message
        message.attach(part)
        LOGGER.info("Sending the email...")
        server.sendmail(re.findall("(?<=<)\\S+(?=>)", config["from"])[0],
                        re.findall("(?<=<)\\S+(?=>)", config["to"]),
                        message.as_string())
    LOGGER.info("Email sent successfully!")


def check_ledger():
    """Runs a check of the ledger.
    """

    # Only run between 07:00 and 23:00
    if 7 <= time.localtime().tm_hour < 23:
        # Make a note of the start time
        print("\nIt's %s and we're doing a check." %
              time.strftime("%d %b %Y at %H:%M:%S"))
    else:
        print("\nIt's %s so we're not doing a check." %
              time.strftime("%d %b %Y at %H:%M:%S"))
        return
    LOGGER.info("\n")

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
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
    new_url, sheet_name = upload_ledger(dir_name=config["dir_name"],
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
    if 'error' in response:
        LOGGER.error("There was an error with the Apps Script function!")
        LOGGER.error(response["error"])
        error = response['error']['details'][0]
        print("Script error message: " + error['errorMessage'])
        if 'scriptStackTraceElements' in error:
            # There may not be a stacktrace if the script didn't start
            # executing.
            print("Script error stacktrace:")
            for trace in error['scriptStackTraceElements']:
                print("\t{0}: {1}".format(trace['function'],
                                          trace['lineNumber']))
        raise Exception(response["error"])

    # Otherwise save the data that the Apps Script returns
    changes = response["response"].get("result")
    print("The Apps Script function executed successfully!")
    LOGGER.info("The Apps Script function executed successfully.")

    # Get the old values of changes
    with open(config["save_data_filepath"], "rb") as save_file:
        old_changes, old_sheet_id = pickle.load(save_file)

    # If there were no changes then do nothing
    # The sheet will have been deleted by the Apps Script
    if changes == "False":
        print("Changes is False, so we'll do nothing.")
        LOGGER.info("Changes is False, so we'll do nothing.")
        os.remove(pdf_filepath)
        os.remove(xlsx_filepath)
        LOGGER.info("The local PDF and XLSX have been deleted successfully.")

    # If the returned changes aren't actually new to us then
    # just delete the new sheet we just made
    # This compares the total income & expenditure
    elif changes[-1][-1][1] == old_changes[-1][-1][1] and \
            changes[-1][-1][2] == old_changes[-1][-1][2]:
        print("The new changes is the same as the old.")
        LOGGER.info("The new changes is the same as the old.")
        LOGGER.info("Deleting the new sheet (that's the same as the old one)...")
        sheet_id = changes.pop(0)
        delete_sheet(sheets_service=sheets,
                     spreadsheet_id=config["destination_sheet_id"],
                     sheet_id=sheet_id)
        os.remove(pdf_filepath)
        os.remove(xlsx_filepath)
        LOGGER.info("The local PDF and XLSX have been deleted successfully!")

    # Otherwise these changes are new
    # Notify the user (via email) and delete the old sheet
    # Save the new data to the save file
    else:
        print("We have some new changes!")
        LOGGER.info("We have some new changes.")
        sheet_id = changes.pop(0)
        send_email(config=parser["email"], changes=changes, pdf_filepath=pdf_filepath, url=new_url)
        LOGGER.info("Deleting the old sheet...")
        delete_sheet(sheets_service=sheets,
                     spreadsheet_id=config["destination_sheet_id"],
                     sheet_id=old_sheet_id)
        LOGGER.info("Saving the new changes and sheet_id...")
        with open(config["save_data_filepath"], "wb") as save_file:
            pickle.dump([changes, sheet_id], save_file)
        LOGGER.info("Save data updated successfully!\n")

    # Note the end time and that we're now finished
    print("It's %s and the check is complete.\n" %
          time.strftime("%d %b %Y at %H:%M:%S"))


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


def send_error_email(config: configparser.SectionProxy, error_stack: str,
                     fails_text: str):
    """Used to email the user about a fatal exception.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param error_stack: the stack trace of the exception
    :type error_stack: str
    :param fails_text: info about the no. of failed attempts
    :type fails_text: str
    """

    # Connect to the server
    LOGGER.info("Connecting to the email server...")
    with smtplib.SMTP(config["host"], int(config["port"])) as server:
        server.starttls()
        server.login(config["username"], config["password"])
        message = MIMEMultipart("alternative")
        message["Subject"] = "ERROR with ledger_checker.py!"
        message["To"] = config["to"]
        message["From"] = config["from"]
        message["X-Priority"] = "1"

        # Prepare the email
        text = ("There was a fatal error with ledger_checker.py! %s"
                "Please see the stack trace below and check the logs.\n\n %s "
                "\n\n\n———\nThis email was sent automatically by a "
                "computer program. If you want to leave some feedback "
                "then please reply directly to it." % (fails_text,
                                                       error_stack))
        message.attach(MIMEText(text, "plain"))

        # Send the email
        LOGGER.info("Sending the exception email...")
        server.sendmail(re.findall("(?<=<)\\S+(?=>)", config["from"])[0],
                        re.findall("(?<=<)\\S+(?=>)", config["to"]),
                        message.as_string())
    LOGGER.info("The email about the exception was sent successfully!")


def main():
    """Manages the checks.

    This function runs the checks of the ledger regularly and gracefully
    handles any errors that occur.
    """

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    save_data_filepath = parser["ledger_checker"]["save_data_filepath"]

    # If the save file doesn't exist then create it
    if not os.path.exists(save_data_filepath):
        LOGGER.warning("The save file doesn't exist, creating a blank one...")
        with open(save_data_filepath, "wb") as save_file:
            pickle.dump([[[[0, 0, 0]]], None], save_file)

    # Run the checker
    # fails counts the number of consecutive failed attempts
    fails = 0
    while fails < ATTEMPTS:
        try:
            check_ledger()

            # If no exception occurred then reset fails
            fails = 0

        # Catch any exception that occurs
        except Exception:
            fails += 1
            error_stack = traceback.format_exc()
            LOGGER.exception("We had a problem and had to stop the checks!")
            LOGGER.error("This is error no. %d.", fails)
            print("It's %s and we've hit a fatal error! "
                  "Go check the log to find out more." %
                  time.strftime("%d %b %Y at %H:%M:%S"))
            if fails == ATTEMPTS:
                fails_text = "This is consecutive failed attempt no. " \
                             "%d so we won't try to " \
                             "make any further checks. " % fails
            else:
                fails_text = "This is consecutive failed attempt no. " \
                             "%d so we will try again " \
                             "in %d seconds. " % (fails, INTERVAL)
            print(fails_text)
            print(traceback.format_exc())

            # Attempt to email the user about the exception
            try:
                print("Sending an email about the exception...")
                send_error_email(config=parser["email"],
                                 error_stack=error_stack,
                                 fails_text=fails_text)
                print("Email sent successfully!\n")
            except Exception:
                print("The email was not sent successfully!")
                print(traceback.format_exc())
                LOGGER.exception("We couldn't send the email about the exception!")

        # Always wait before the next one
        finally:
            time.sleep(INTERVAL)


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
