from __future__ import print_function

import configparser
import logging
import os
import pickle
import time

# noinspection PyUnresolvedReferences
import pyexcel_xlsx
import requests
from appJar import gui

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

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