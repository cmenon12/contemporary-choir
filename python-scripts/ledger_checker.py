"""Checks the ledger for any changes and notifies the user via email.

This script is designed to check the ledger for any new changes, and
notify the user by email of any changes found.

The LedgerCheckerSaveFile class is used to represent the save file that
maintains persistence between executions. This ensures that duplicate
emails aren't sent. It saves the most recent changes and copy of the
ledger, allowing this old copy to be attached to a future email. It
also saves the stack traces, allowing the user to be notified of
consecutive errors.

This script contains standalone functions to run the check. This
includes separate functions to prepare and send the emails.

Before running this script, make sure you've got a valid config file
and a valid set of credentials (saved to config.ini and
credentials.json respectively).

This script relies on the classes in ledger_fetcher.py and
custom_exceptions.py. It also relies on two Apps Scripts,
ledger-comparison.gs and ledger-checker.gs, to neatly format the ledger
and compare it to an older version. These should be created in an Apps
Script project linked to the Google Sheet that the ledger is uploaded
to. Finally, it relies on email-template.html to form the email.
"""

import configparser
import email.utils
import imaplib
import json
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
from typing import Optional

import html2text
import timeago
from babel.numbers import format_currency
from jinja2 import Environment, FileSystemLoader

from custom_exceptions import AppsScriptApiError
from ledger_fetcher import Ledger, CustomEncoder, authorize

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

    def __init__(self, save_data_filepath: str):
        """Constructor.

        :param save_data_filepath: where the save data is saved
        :type save_data_filepath: str
        """

        # Try to open the save file if it exists
        try:
            with open(save_data_filepath, "rb") as save_file:

                # Load it and set the attributes for self from it
                inst = pickle.load(save_file)
                for k in inst.__dict__.keys():
                    setattr(self, k, getattr(inst, k))

            if not isinstance(inst, LedgerCheckerSaveFile):
                raise FileNotFoundError("The save data file is invalid.")

        # If it fails then just initialise with blank values
        except (FileNotFoundError, EOFError):

            self.save_data_filepath = save_data_filepath
            self.stack_traces = []
            self.changes = None
            self.most_recent_ledger = None
            self.changes_ledger = None
            self.error_email_id = None
            self.success_email_id = None

        self.save_data()

    def save_data(self) -> None:
        """Save to the actual file."""

        with open(self.save_data_filepath, "wb") as save_file:
            pickle.dump(self, save_file)
            save_file.close()

        self.log()

    def new_check_success(self, new_ledger: Ledger, changes: dict = None) -> None:
        """Runs when a check was successful.

        :param new_ledger: the ledger that was just checked
        :type new_ledger: Ledger
        :param changes: the changes made to the ledger
        :type changes: dict
        """

        # This means that we have new changes (and an email was sent)
        if changes is not None:
            self.changes = changes
            self.changes_ledger = new_ledger

        self.most_recent_ledger = new_ledger
        self.stack_traces.clear()
        self.error_email_id = None
        self.save_data()
        LOGGER.info("Successful check saved to %s",
                    self.save_data_filepath)

    def new_check_fail(self, stack_trace: str) -> None:
        """Runs when a check failed.

        :param stack_trace: the stack trace
        :type stack_trace: str
        """

        date = datetime.now().strftime("%A %d %B %Y AT %H:%M:%S")
        self.stack_traces.append(("ERROR ON %s\n%s" % (date.upper(), stack_trace)))
        self.save_data()
        LOGGER.info("Failed check saved to %s",
                    self.save_data_filepath)

    def update_error_email_id(self, email_id: str) -> None:
        """Sets and saves the ID of the last error email.

        :param email_id: the ID to save
        :type email_id: str
        """

        self.error_email_id = email_id
        self.save_data()

    def get_error_email_id(self) -> Optional[str]:
        """Gets and returns the ID of the last error email, if it exists."""

        return self.error_email_id

    def update_success_email_id(self, email_id: str) -> None:
        """Sets and saves the ID of the last success email.

        :param email_id: the ID to save
        :type email_id: str
        """

        self.success_email_id = email_id
        self.save_data()

    def get_success_email_id(self) -> Optional[str]:
        """Gets and returns the ID of the last success email, if it exists."""

        return self.success_email_id

    def get_changes(self) -> dict:
        """Gets and returns changes."""

        return self.changes

    def get_most_recent_ledger(self) -> Ledger:
        """Gets and returns the most recent ledger."""

        return self.most_recent_ledger

    def get_changes_ledger(self) -> Ledger:
        """Gets and returns the ledger associated with changes."""

        return self.changes_ledger

    def get_stack_traces(self) -> list:
        """Gets and returns the list of stack traces."""

        return self.stack_traces

    def log(self) -> None:
        """Logs the object to the log."""

        LOGGER.info(json.dumps(self.__dict__, cls=CustomEncoder))


def prepare_email_body(changes: dict, sheet_url: str, pdf_url: str,
                       last_check: str, ledger_plurality: str = "(s)") -> tuple:
    """Prepares an email detailing the changes to the ledger.

    :param changes: the changes made to the ledger
    :type changes: dict
    :param sheet_url: the URL of the Google Sheet to link to
    :type sheet_url: str
    :param pdf_url: the URL of the PDF to link to
    :type pdf_url: str
    :param last_check: when the previous check was made
    :type last_check: str
    :param ledger_plurality: indicates if there's multiple PDF ledgers
    :type ledger_plurality: str
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

    # Render the template
    root = os.path.dirname(os.path.abspath(__file__))
    env = Environment(loader=FileSystemLoader(root))
    template = env.get_template("email-template.html")
    html = template.render(changes=changes,
                           sheet_url=sheet_url,
                           pdf_url=pdf_url,
                           last_check=last_check,
                           ledger_plurality=ledger_plurality)

    # Create the plain-text version of the message
    text_maker = html2text.HTML2Text()
    text_maker.ignore_links = True
    text_maker.bypass_tables = False
    text = text_maker.handle(html)
    LOGGER.info("Email HTML and plain-text created successfully.")

    return text, html


def send_success_email(config: configparser.SectionProxy,
                       save_data: LedgerCheckerSaveFile, changes: dict,
                       new_ledger: Ledger, old_ledger: Ledger) -> None:
    """Sends an email detailing the changes.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param save_data: the save data object
    :type save_data: LedgerCheckerSaveFile
    :param changes: the changes made to the ledger
    :type changes: dict
    :param new_ledger: the ledger with these new changes
    :type new_ledger: Ledger
    :param old_ledger: the ledger just before these new changes
    :type old_ledger: Ledger
    """

    message = MIMEMultipart("alternative")
    message["Subject"] = changes["societyName"] + " Ledger Update"
    message["To"] = config["to"]
    message["From"] = config["from"]
    if save_data.get_changes() is not None and \
            changes["oldLedgerTimestamp"] == save_data.get_changes()["oldLedgerTimestamp"]:
        message["In-Reply-To"] = save_data.get_success_email_id()

    # Prepare the email
    if old_ledger is not None:
        old_timestamp = old_ledger.get_timestamp()
        last_check = " since the last check %s on %s" % (timeago.format(old_timestamp.replace(tzinfo=None)),
                                                         old_timestamp.strftime("%A %d %B %Y at %H:%M:%S"))
        ledger_plurality = "s"
    else:
        last_check = ", although we don't know how new these changes are"
        ledger_plurality = ""
    text, html = prepare_email_body(changes=changes,
                                    sheet_url=new_ledger.get_sheets_data(convert=False,
                                                                         save=False,
                                                                         upload=False)["url"],
                                    pdf_url=new_ledger.get_drive_pdf_url(save=False,
                                                                         upload=False),
                                    last_check=last_check,
                                    ledger_plurality=ledger_plurality)

    # Turn these into plain/html MIMEText objects
    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(MIMEText(text, "plain"))
    message.attach(MIMEText(html, "html"))

    # Attach the new ledger
    part = MIMEBase("application", "pdf")
    part.set_payload(new_ledger.get_pdf_file())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition",
                    "attachment; filename=\"NEW %s\"; creation-date=\"%s\"; "
                    "modification-date=\"%s\"; read-date=\"%s\"" %
                    (new_ledger.get_pdf_filename(),
                     email.utils.format_datetime(new_ledger.get_timestamp()),
                     email.utils.format_datetime(new_ledger.get_timestamp()),
                     email.utils.format_datetime(new_ledger.get_timestamp())))
    part.add_header("Content-Description",
                    "NEW %s" % new_ledger.get_pdf_filename())
    message.attach(part)

    # Attach the old ledger if it exists
    if old_ledger is not None:
        part = MIMEBase("application", "pdf")
        part.set_payload(old_ledger.get_pdf_file())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        "attachment; filename=\"OLD %s\"; creation-date=\"%s\"; "
                        "modification-date=\"%s\"; read-date=\"%s\"" %
                        (old_ledger.get_pdf_filename(),
                         email.utils.format_datetime(old_ledger.get_timestamp()),
                         email.utils.format_datetime(old_ledger.get_timestamp()),
                         email.utils.format_datetime(old_ledger.get_timestamp())))
        part.add_header("Content-Description",
                        "OLD %s" % old_ledger.get_pdf_filename())
        message.attach(part)

    # Send the email and save the ID
    save_data.update_success_email_id(send_email(config=config, message=message))


def send_error_email(config: configparser.SectionProxy,
                     save_data: LedgerCheckerSaveFile) -> None:
    """Used to email the user about a fatal exception.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param save_data: the save data object
    :type save_data: LedgerCheckerSaveFile
    """

    # Get the count of errors and stack traces
    error_count = len(save_data.get_stack_traces())
    stack_traces = ""
    for error in save_data.get_stack_traces():
        stack_traces += error
        stack_traces += "\n\n"

    # Create the message
    message = MIMEMultipart("alternative")
    message["Subject"] = "ERROR with ledger_checker.py!"
    message["To"] = config["to"]
    message["From"] = config["from"]
    message["X-Priority"] = "1"
    if save_data.get_error_email_id() is not None:
        message["In-Reply-To"] = save_data.get_error_email_id()

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

    # Find out when the most recent successful check was (if ever)
    if save_data.get_most_recent_ledger() is not None:
        most_recent = save_data.get_most_recent_ledger() \
            .get_timestamp().strftime("%A %d %B %Y at %H:%M:%S")
    else:
        most_recent = "never"

    text = ("There were %d consecutive and fatal errors with ledger_checker.py! "
            "The most recent successful check was %s.\n\n"
            "Please see the stack traces below and check the log. Note that "
            "future executions will continue as scheduled%s\n\n\n%s"
            "———\nThis email was sent automatically by a "
            "computer program (https://github.com/cmenon12/contemporary-choir). "
            "If you want to leave some feedback "
            "then please reply directly to it." % (error_count,
                                                   most_recent,
                                                   future_attempts,
                                                   stack_traces))
    message.attach(MIMEText(text, "plain"))

    # Send the email and save the ID
    save_data.update_error_email_id(send_email(config=config, message=message))


def send_email(config: configparser.SectionProxy, message: MIMEMultipart) -> str:
    """Sends the message using the config with SSL.

    :param config: the configuration for the email
    :type config: configparser.SectionProxy
    :param message: the message to send
    :type message: MIMEMultipart
    :return: the email ID
    :rtype: str
    """

    # Add a few headers to the message
    message["Date"] = email.utils.formatdate()
    email_id = email.utils.make_msgid(domain=config["smtp_host"])
    message["Message-ID"] = email_id

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
            server.append("INBOX.Sent", "\\Seen",
                          imaplib.Time2Internaldate(time.time()),
                          message.as_string().encode("utf8"))
        LOGGER.info("Email saved successfully!")

    return email_id


def check_ledger(save_data: LedgerCheckerSaveFile,
                 parser: configparser.ConfigParser) -> None:
    """Runs a check of the ledger.

    :param save_data: the save data object
    :type save_data: LedgerCheckerSaveFile
    :param parser: the whole config
    :type parser: configparser.ConfigParser
    :raises: AppsScriptApiError
    """

    # Fetch info from the config
    config = parser["ledger_checker"]
    expense365 = parser["eXpense365"]

    # Download the ledger, convert it, and upload it to Google Sheets
    print("Downloading the PDF...")
    ledger = Ledger(config=config, expense365=expense365)
    print("Converting the ledger...")
    ledger.get_xlsx_filepath()
    print("Uploading the ledger to Google Sheets...")
    sheets_data = ledger.get_sheets_data()
    print("Ledger downloaded, converted, and uploaded successfully.")

    # Connect to the Apps Script service and attempt to execute it
    socket.setdefaulttimeout(600)
    _, _, apps_script = authorize(pushbullet=ledger.pushbullet,
                                  open_browser=ledger.browser_path)
    print("Executing the Apps Script function (this may take some time)...")
    LOGGER.info("Starting the Apps Script function...")
    body = {"function": config["function"],
            "parameters": [sheets_data["name"], config["compare_sheet_id"], config["compare_sheet_name"]]}
    response = apps_script.scripts().run(body=body,
                                         scriptId=config["deployment_id"]).execute()

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
        raise AppsScriptApiError(response["error"])

    # Otherwise save the data that the Apps Script returns
    changes = response["response"].get("result")
    LOGGER.info(changes)
    print("The Apps Script function executed successfully!")
    LOGGER.info("The Apps Script function executed successfully.")

    # Get the old values of changes
    old_changes = save_data.get_changes()

    # If there were no changes then do nothing
    if changes is False or changes.lower() == "false":
        print("Changes is False, so we'll do nothing.")
        LOGGER.info("Changes is False, so we'll do nothing.")
        ledger.delete_pdf()
        ledger.delete_xlsx()
        ledger.delete_sheet()
        save_data.new_check_success(new_ledger=ledger)

    else:
        changes = json.loads(response["response"].get("result"))

        # If the returned changes aren't actually new to us then
        # just delete the new sheet we just made
        # This compares the total income & expenditure
        if old_changes is not None and \
                format_currency(changes["grandTotal"]["totalIn"], "GBP", locale="en_GB") == \
                old_changes["grandTotal"]["totalIn"] and \
                format_currency(changes["grandTotal"]["totalOut"], "GBP", locale="en_GB") == \
                old_changes["grandTotal"]["totalOut"] and \
                format_currency(changes["grandTotal"]["balanceBroughtForward"], "GBP", locale="en_GB") \
                == old_changes["grandTotal"]["balanceBroughtForward"]:
            print("The new changes is the same as the old.")
            LOGGER.info("The new changes is the same as the old.")
            ledger.delete_pdf()
            ledger.delete_xlsx()
            ledger.delete_sheet()
            save_data.new_check_success(new_ledger=ledger)

        # Otherwise these changes are new
        # Update the PDF ledger in the user's Google Drive
        # Notify the user (via email) and hide the old sheet
        # Save the new data to the save file
        else:
            print("We have some new changes!")
            LOGGER.info("We have some new changes.")
            ledger.update_drive_pdf()
            send_success_email(config=parser["email"],
                               save_data=save_data,
                               changes=changes,
                               new_ledger=ledger,
                               old_ledger=save_data.get_most_recent_ledger())
            print("Email sent successfully!")
            LOGGER.info("Hiding the old sheet...")
            if save_data.get_changes_ledger() is not None:
                save_data.get_changes_ledger().hide_sheet()
            save_data.new_check_success(new_ledger=ledger, changes=changes)


def main() -> None:
    """Manages the checks.

    This function runs the checks of the ledger regularly and gracefully
    handles any errors that occur.
    """

    LOGGER.info("\n")

    # Check that the config file exists
    try:
        open(CONFIG_FILENAME, "rb")
        LOGGER.info("Loaded config %s.", CONFIG_FILENAME)
    except FileNotFoundError as e:
        print("The config file doesn't exist!")
        LOGGER.info("Could not find config %s, exiting.", CONFIG_FILENAME)
        time.sleep(5)
        raise FileNotFoundError("The config file doesn't exist!") from e

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read(CONFIG_FILENAME)

    save_data = LedgerCheckerSaveFile(parser["ledger_checker"]["save_data_filepath"])

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

    time.sleep(5)


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(filename="ledger_checker.log",
                        filemode="a",
                        format="%(asctime)s | %(levelname)s : %(message)s",
                        level=logging.INFO)
    LOGGER = logging.getLogger(__name__)

    main()

else:
    LOGGER = logging.getLogger(__name__)
