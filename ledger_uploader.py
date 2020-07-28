import configparser

from appJar import gui

# Create the GUI
APP_GUI = gui(showIcon=False)
APP_GUI.setOnTop()


def upload_ledger(dir_name: str):
    """Uploads the ledger to the Google Sheet with the macro.

    :param dir_name: the default directory to open the xlsx file from
    :type dir_name: str
    """

    # Open the spreadsheet
    try:
        xlsx_filepath = APP_GUI.openBox(title="Open spreadsheet",
                                        dirName=dir_name,
                                        fileTypes=[("Office Open XML Workbook", "*.xlsx")],
                                        asFile=True).name
        with open(xlsx_filepath, "wb") as xlsx_file:
            pass  # might need to save content
        APP_GUI.removeAllWidgets()

    except Exception as err:
        print("There was an error opening the spreadsheet.")
        raise SystemExit(err)

    # upload as a Google Sheet


def main():
    """Runs the program."""

    # Fetch info from the config
    parser = configparser.ConfigParser()
    parser.read("config.ini")
    dir_name = parser.get("expense365", "dir_name")

    upload_ledger(dir_name)


if __name__ == "__main__":
    main()
