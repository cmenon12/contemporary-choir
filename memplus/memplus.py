from __future__ import print_function

import cgi
import configparser
import csv
import logging
import os
import pickle
import sys
import time
from datetime import datetime
from io import StringIO
from typing import Optional

import numpy as np
import pandas as pd
# noinspection PyUnresolvedReferences
import pyexcel_xlsx
import requests
from appJar import gui
from progress.bar import Bar

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
        with open(config["cookie_file"], "rb") as file:
            session.cookies.update(pickle.load(file))
        LOGGER.info("Loaded cookies from %s.", config["cookie_file"])

    else:
        # Get the magic link to login
        print(f"Visit {config['domain']}student/auth/magic-link/ and request a magic link.")
        auth_link = input("Enter the magic link: ")

        # Login
        response = session.get(auth_link)
        response.raise_for_status()
        LOGGER.info("Logged in using the magic link.")

        # Save the cookies
        with open(config["cookie_file"], "wb") as file:
            pickle.dump(session.cookies, file)
        LOGGER.info("Saved cookies to %s.", config["cookie_file"])

    return session


def download_knowledgebase(session: requests.Session, config: configparser.SectionProxy,
                           app_gui: gui, start: int = 1, end: int = 200) -> None:
    """Download a list of articles in the Knowledgebase."""

    LOGGER.info("Downloading knowledgebase from %d to %d...", start, end)

    pbar = Bar("Downloading", max=end)
    result = []
    for i in range(start, end + 1):
        url = f"{config['domain']}auth/committee/knowledge-base/{i}"
        try:
            # Download the article
            LOGGER.debug("Downloading article %d...", i)
            response = session.get(url)
            response.raise_for_status()

            # Extract the title and dates from it
            page_soup = BeautifulSoup(response.text, "lxml")
            title = page_soup.find_all("h3", {"class": "text-2xl font-bold tracking-tight text-gray-900"})[0].text
            date_element = page_soup.find_all("p", {
                "class": "flex justify-start items-center text-gray-500 text-sm truncate"})
            dates = ". ".join([x.strip() for x in date_element[0].text.split(".")])

            # Add the article to the list
            result.append([i, title, dates, url])

        except HTTPError as error:
            # Add the error to the list
            result.append([i, f"{error.response.status_code}", "", url])

        pbar.next()

    pbar.finish()
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


def download_members(session: requests.Session, config: configparser.SectionProxy,
                     app_gui: gui, society_id: Optional[str] = None) -> None:
    """Download a comprehensive CSV of members."""

    LOGGER.info("Downloading members for society %s...", str(society_id if society_id else config['society_id']))

    # Download the CSV of members
    url = f"{config['domain']}auth/committee/group/{society_id if society_id else config['society_id']}/members/export"
    response = session.get(url)
    response.raise_for_status()
    society_name = cgi.parse_header(response.headers["content-disposition"])[1]["filename"].split(".")[0]
    df: pd.DataFrame = pd.read_csv(StringIO(response.text), sep=",")

    # Add empty columns
    new_columns = ["MemPlus ID", "Student Type", "Active", "Started", "Expires", "Paid", "Current Position"]
    for column in new_columns:
        df[column] = ""

    # Download the page of members
    url = f"{config['domain']}auth/committee/group/{society_id if society_id else config['society_id']}/members"
    response = session.get(url)
    response.raise_for_status()

    # Extract all the members
    page_soup = BeautifulSoup(response.text, "lxml")
    members = page_soup.find_all("table")[0].find_all("tr")[1:]
    pbar = Bar("Downloading", max=len(members))
    for member in members:

        # Download that member's page
        member_url = member.find_all("a")[0]["href"]
        member_response = session.get(member_url)
        member_response.raise_for_status()
        member_soup = BeautifulSoup(member_response.text, "lxml")

        # Find their MemPlus ID and student ID
        student_id = member_soup.find(lambda tag: tag.name == "dt" and "Student ID" in tag.text).find_next_sibling(
            "dd").text.strip()
        df.loc[df["Student ID"] == student_id, "MemPlus ID"] = int(member_url.split("/")[-1])

        # Find and save all other attributes
        for div in member_soup.find_all("section", {"class": "grid lg:grid-cols-2 gap-8"})[0].find_all("div"):
            if not div.find_all("dt", recursive=False):
                continue
            key = div.find_all("dt", recursive=False)[0].text.strip()
            if key == "Student Type":
                df.loc[df["Student ID"] == student_id, "Student Type"] = div.find_all("code")[0].text.strip()
            elif key in ["Started", "Expires"]:
                df.loc[df["Student ID"] == student_id, key] = div.find_all("dd")[0].text.strip()[:19]
            elif key == "Current Position":
                df.loc[df["Student ID"] == student_id, "Current Position"] = div.find_all("dd")[0].text.strip()
            elif key in ["Active", "Paid"]:
                df.loc[df["Student ID"] == student_id, key] = bool(div.find_all("span", {
                    "class": "text-green-400 flex items-center"}))

        pbar.next()

    pbar.finish()
    LOGGER.debug("Downloaded %d members.", len(members))

    # Ask the user where to save it
    filename = f"{society_name}-members-{datetime.now().isoformat().replace(':', '_')}.csv"
    csv_file_box = app_gui.saveBox("Save Members",
                                   fileName=filename,
                                   fileExt=".csv",
                                   fileTypes=[("Comma-Separated Values", "*.csv")],
                                   asFile=True)
    if csv_file_box is None:
        LOGGER.warning("The user cancelled saving the CSV.")
        raise SystemExit("User cancelled saving the CSV.")
    filepath = csv_file_box.name
    app_gui.removeAllWidgets()

    LOGGER.debug("Saving members to %s...", filepath)

    # Save to a CSV
    df.to_csv(filepath, index=False)

    LOGGER.info("Saved members to %s.", filepath)


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

    # Run
    session = get_authorised_session(config)
    download_knowledgebase(session, config, app_gui)
    download_members(session, config, app_gui)


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(handlers=[
        logging.FileHandler("memplus_base.log"),
        logging.StreamHandler(sys.stdout)
    ],
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
