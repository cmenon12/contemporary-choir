import base64
import configparser
import os
import pickle
import re
import smtplib
import socket
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import googleapiclient
import html2text
from babel.numbers import format_currency
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import errors
from googleapiclient.discovery import build
from jinja2 import Environment, FileSystemLoader
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


def authorize() -> tuple:
    """Authorizes access to the user's scripts and Google Sheets.

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
    apps_script_service = build("script", "v1", credentials=credentials,
                                cache_discovery=False)
    sheets_service = build("sheets", "v4", credentials=credentials,
                           cache_discovery=False)
    return apps_script_service, sheets_service


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
    print("Creating the email...")
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
    print(format_currency(int(grand_total), 'GBP', locale='en_GB'))

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
    print("Connecting to the email server...")
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
        print("Attaching the ledger...")
        with open(pdf_filepath, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        # Add header as key/value pair to attachment part
        filename = pdf_filepath.split("\\")[-1]
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add the attachment to message and send the message
        message.attach(part)
        print("Sending the email...")
        server.sendmail(re.findall("(?<=<)\\S+(?=>)", config["from"])[0],
                        re.findall("(?<=<)\\S+(?=>)", config["to"]),
                        message.as_string())
    print("Email sent successfully!")


def run_task():

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

    # Connect to the Apps Script service and attempt to execute it
    socket.setdefaulttimeout(600)
    apps_script_service, sheets_service = authorize()
    print("Connected to the Apps Script service.")
    try:
        print("Executing the Apps Script function (this may take some time)...")
        body = {"function": config["function"], "parameters": sheet_name}
        response = apps_script_service.scripts().run(body=body,
                                                     scriptId=config["script_id"]).execute()

        # Catch and print an error during execution
        if 'error' in response:
            error = response['error']['details'][0]
            print("Script error message: " + error['errorMessage'])
            if 'scriptStackTraceElements' in error:
                # There may not be a stacktrace if the script didn't start
                # executing.
                print("Script error stacktrace:")
                for trace in error['scriptStackTraceElements']:
                    print("\t{0}: {1}".format(trace['function'],
                                              trace['lineNumber']))
            raise SystemExit()

        # Otherwise save the data that the Apps Script returns
        changes = response["response"].get("result")

    # Catch an error making the request to the API
    except errors.HttpError as err:
        raise SystemExit(err)
    print("The Apps Script function executed successfully!")

    if changes != "False":
        send_email(config=parser["email"], changes=changes, pdf_filepath=pdf_filepath, url=new_url)
    else:
        print("No changes were found.")

    return sheet_name, changes


if __name__ == "__main__":
    run_task()
