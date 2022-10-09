from __future__ import print_function

import configparser
import csv
import logging
import os
import pickle
import sys
import time
from datetime import datetime

# noinspection PyUnresolvedReferences
import pyexcel_xlsx
import requests
from appJar import gui

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

from bs4 import BeautifulSoup

from requests import HTTPError

# The name of the config file
CONFIG_FILENAME = "config.ini"


def get_authorised_session(config: configparser.SectionProxy, use_file: bool = True) -> requests.Session:
    """Gets an authorised session."""

    # Create the session
    session = requests.Session()
    session.headers.update({"User-Agent": config["user_agent"]})

    # Load the cookies from the file
    if os.path.exists(config["cookie_file"]) and use_file:
        with open(config["cookie_file"], "rb") as f:
            session.cookies.update(pickle.load(f))
            LOGGER.info("Loaded cookies from %s.", config["cookie_file"])

    else:
        # Get the magic link to login
        print(f"Visit {config['magic_link_url']} and request a magic link.")
        auth_link = input("Enter the magic link: ")

        # Login
        response = session.get(auth_link)
        response.raise_for_status()
        LOGGER.info("Logged in using the magic link.")

    return session


def download_knowledgebase(session: requests.Session, config: configparser.SectionProxy,
                           app_gui: gui, min: int = 1, max: int = 200) -> None:
    """Download a list of articles in the Knowledgebase."""

    LOGGER.info("Downloading knowledgebase from %d to %d...", min, max)

    result = []
    for i in range(min, max + 1):
        url = f"{config['knowledgebase_url']}{i}"
        try:
            # Download the article
            LOGGER.debug("Downloading article %d..." % i)
            response = session.get(url)
            response.raise_for_status()

            # Extract the title and dates from it
            page_soup = BeautifulSoup(response.text, "lxml")
            title = page_soup.find_all("h3", {"class": "text-2xl font-bold tracking-tight text-gray-900"})[0].text
            dates = page_soup.find_all("p", {"class": "flex justify-start items-center text-gray-500 text-sm truncate"})[0].text

            # Add the article to the list
            result.append([i, title, dates, url])

        except HTTPError as error:
            # Add the error to the list
            result.append([i, f"{error.response.status_code} {error.response.reason}", "", url])

    LOGGER.debug("Downloaded %d articles, %d of which were valid.", len(result), len([x for x in result if x[2] != ""]))

    # Ask the user where to save it
    filename = f"guild-knowledgebase-{datetime.now().isoformat().replace(':', '_')}.csv"
    csv_file_box = app_gui.saveBox("Save Knowledgebase",
                                   fileName=filename,
                                   fileExt=".csv",
                                   fileTypes=[("Comma-Separated Values", "*.csv")],
                                   asFile=True)
    if csv_file_box is None:
        LOGGER.warning("The user cancelled saving the CSV.")
        raise SystemExit("User cancelled saving the CSV.")
    filepath = csv_file_box.name
    app_gui.removeAllWidgets()

    LOGGER.debug("Saving knowledgebase to %s...", filepath)

    # Save to a CSV
    with open(filepath, "w", encoding="UTF8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Index", "Title", "Dates", "URL"])
        writer.writerows(result)

    LOGGER.info("Saved knowledgebase to %s.", filepath)


def main(app_gui: gui) -> None:
    """Runs the program.

    :param app_gui: the appJar GUI to use
    :type app_gui: appJar.appjar.gui
    """

    # Check that the config file exists
    try:
        open(CONFIG_FILENAME)
        LOGGER.info("Loaded config %s.", CONFIG_FILENAME)
    except FileNotFoundError as e:
        print("The config file doesn't exist!")
        LOGGER.info("Could not find config %s, exiting.", CONFIG_FILENAME)
        time.sleep(5)
        raise FileNotFoundError("The config file doesn't exist!") from e

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read(CONFIG_FILENAME)
    config = parser["memplus"]


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(filename="memplus_base.log",
                        filemode="a",
                        format="%(asctime)s | %(levelname)5s in %(module)s.%(funcName)s() on line %(lineno)-3d | %(message)s",
                        level=logging.INFO)
    LOGGER = logging.getLogger(__name__)

    # Create the GUI
    appjar_gui = gui(showIcon=False)
    appjar_gui.setOnTop()
    appjar_gui.setFont(size=12)

    main(appjar_gui)

else:
    LOGGER = logging.getLogger(__name__)