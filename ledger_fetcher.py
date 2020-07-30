import base64
import configparser
import webbrowser
from datetime import datetime

import requests
from appJar import gui
from dateutil import tz

import ledger_uploader

# 30 is used for the ledger, 31 for the balance
REPORT_ID = "30"

# 35778 refers to Contemporary Choir
GROUP_ID = "35778"

# Don't know what this is, likely to be unused
SUB_GROUP_ID = "0"


def download_pdf(auth: str, dir_name: str, app_gui: gui) -> str:
    """Downloads the ledger from expense365.com.

    :param auth: the authentication header with the email and password
    :type auth: str
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
            ',"UserGroupID":' + GROUP_ID +
            ',"SubGroupID":' + SUB_GROUP_ID + '}')

    # Make the request and check it was successful
    print("Making the HTTP request...")
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
    file_name = "C30 Ledger " + date_string

    # Attempt to save the PDF (fails gracefully if the user cancels)
    try:
        print("Saving the PDF...")
        pdf_filepath = app_gui.saveBox("Save ledger",
                                       dirName=dir_name,
                                       fileName=file_name,
                                       fileExt=".pdf",
                                       fileTypes=[("PDF file", "*.pdf")],
                                       asFile=True).name
        with open(pdf_filepath, "wb") as pdf_file:
            pdf_file.write(response.content)
        app_gui.removeAllWidgets()

    except Exception as err:
        print("There was an error saving the PDF ledger!")
        raise SystemExit(err)

    # If successful then return the file path
    else:
        print("PDF ledger saved successfully!")
        return pdf_filepath


def main(app_gui: gui):
    """Runs the program.

    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    """

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    email = parser.get("ledger_fetcher", "email")
    password = parser.get("ledger_fetcher", "password")
    dir_name = parser.get("ledger_fetcher", "dir_name")
    browser_path = parser.get("ledger_fetcher", "browser_path")

    # Prepare the authentication
    data = email + ":" + password
    auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

    # Download the PDF, returning the file path
    pdf_filepath = download_pdf(auth=auth, dir_name=dir_name, app_gui=app_gui)

    # Ask the user if they want to open it
    if app_gui.yesNoBox("Open PDF?",
                        "Do you want to open the ledger?") is True:
        # If so then open it in the prescribed browser
        open_path = "file://///" + pdf_filepath
        webbrowser.register("my-browser",
                            None,
                            webbrowser.BackgroundBrowser(browser_path))
        webbrowser.get(using="my-browser").open(open_path)

    # Ask the user if they want to open pdftoexcel.com
    if app_gui.yesNoBox("Open pdftoexcel.com?",
                        "Do you want to open https://www.pdftoexcel.com/?") is True:
        # If so then open it in the prescribed browser
        open_path = "https://www.pdftoexcel.com/"
        webbrowser.register("my-browser",
                            None,
                            webbrowser.BackgroundBrowser(browser_path))
        webbrowser.get(using="my-browser").open(open_path)

        # If they have converted the spreadsheet then
        # Ask the user if they want to upload it to New Ledger Macro
        if app_gui.yesNoBox("Upload the spreadsheet?",
                            "Do you want to upload the spreadsheet to New Ledger Macro?\n"
                            "Click YES when you've downloaded it.") is True:
            # If so then run the ledger_uploader
            ledger_uploader.main(app_gui)


if __name__ == "__main__":
    # Create the GUI
    appjar_gui = gui(showIcon=False)
    appjar_gui.setOnTop()

    main(appjar_gui)
