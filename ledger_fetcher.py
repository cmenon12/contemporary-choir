import base64
import configparser
import webbrowser
from datetime import datetime

import requests
from appJar import gui
from dateutil import tz

# 30 is used for the ledger, 31 for the balance
REPORT_ID = "30"

# 35778 refers to Contemporary Choir
GROUP_ID = "35778"

# Don't know what this is, likely to be unused
SUB_GROUP_ID = "0"

# Create the GUI
APP_GUI = gui(showIcon=False)
APP_GUI.setOnTop()


def download_pdf(auth: str, dir_name: str) -> str:
    """Downloads the ledger from expense365.com.

    :param auth: the authentication header with the email and password
    :type auth: str
    :param dir_name: the default directory to save the PDF
    :type dir_name: str
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
    print("Making the HTTP request.")
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
        pdf_filepath = APP_GUI.saveBox("Save ledger",
                                       dirName=dir_name,
                                       fileName=file_name,
                                       fileExt=".pdf",
                                       fileTypes=[("PDF file", "*.pdf")],
                                       asFile=True).name
        with open(pdf_filepath, "wb") as pdf_file:
            pdf_file.write(response.content)
        APP_GUI.removeAllWidgets()

    except Exception as err:
        print("There was an error saving the PDF report.")
        raise SystemExit(err)

    # If successful then return the file path
    else:
        print("PDF report saved successfully.")
        return pdf_filepath


def main():
    """Runs the program."""

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    email = parser.get("expense365", "email")
    password = parser.get("expense365", "password")
    dir_name = parser.get("expense365", "dir_name")
    browser_path = parser.get("expense365", "browser_path")

    # Prepare the authentication
    data = email + ":" + password
    auth = "Basic " + str(base64.b64encode(data.encode("utf-8")).decode())

    # Download the PDF, returning the file path
    pdf_filepath = download_pdf(auth=auth, dir_name=dir_name)

    # Ask the user if they want to open it
    if APP_GUI.yesNoBox("Open PDF?",
                        "Do you want to open the ledger?") is True:
        # If so then open it in the prescribed browser
        open_path = "file://///" + pdf_filepath
        webbrowser.register('my-browser',
                            None,
                            webbrowser.BackgroundBrowser(browser_path))
        webbrowser.get(using='my-browser').open(open_path)


if __name__ == "__main__":
    main()
