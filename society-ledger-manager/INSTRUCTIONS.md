# Instructions
Note that you can (and probably should for convenience) use the same Google Cloud Platform Project throughout.

## Python Scripts
1. Clone this repository to your local computer.
2. Download and install Python. This was built on Python 3.7, but may work on other versions.
3. Open a terminal window within the `python-scripts` folder of this repository.
4. Install the required packages using `pip install -r requirements.txt`.
5. Rename [`python-scripts/config-template.ini`](python-scripts/config-template.ini) to `python-scripts/config.ini` and update the values within it with your own data. 
   1. For the `destination_sheet_id` and `destination_sheet_name`, enter the ID and name of a Google Sheet to upload the ledger to. The ID is the long string of characters in the URL (i.e. https://docs.google.com/spreadsheets/d/thisistheid/edit), and the name is just the name of the file.
   2. For the `pdf_ledger_id` and `pdf_ledger_name`, enter the ID and name of a copy of the PDF ledger already in Drive. 
   3. For Pushbullet, get the access token from [https://www.pushbullet.com/#settings/account](https://www.pushbullet.com/#settings/account), and the device name as it appears in the list of devices. This feature is optional and can be set to `false` to disable.  
   4. You don't need to fill out the `[ledger_checker]` and `[email]` sections if you're only using [`ledger_fetcher.py`](python-scripts/ledger_fetcher.py).
6. Create your own Google Cloud Platform Project [here](https://console.cloud.google.com/projectcreate).
7. Enable the Google Drive and Google Sheets APIs [here](https://console.cloud.google.com/apis/library).
8. Configure the OAuth consent screen for your project [here](https://console.cloud.google.com/apis/credentials/consent). 
   1. The user type is External, and you can choose all the other details yourself.  * Don't enter any scopes.
   2. Add yourself as a test user.
   3. Click `BACK TO DASHBOARD` and then click `PUBLISH APP`. This won't suddenly make it available for everyone to use, but having the app in 'production' as opposed to in 'testing' ensures that credentials aren't revoked every seven days. Don't click `PREPARE FOR VERIFICATION`.
9. Create some OAuth Credentials [here](https://console.cloud.google.com/apis/credentials). 
   1. Click on `+ CREATE CREDENTIALS` at the top and select `OAuth client ID`. 
   2. Select `Desktop app` as the client type and enter a name (e.g. `Python scripts for checking the ledger`). 
   3. Click `CREATE` and close the box that pops up with the client ID and client secret. 
   4. Find the credentials you just created in the list, and click the download icon on the right to download the client secret. 
   5. Save this as `credentials.json` in the [`python-scripts`](python-scripts) folder.

Additionally, for [`ledger_checker.py`](python-scripts/ledger_checker.py):

10. Repeat step 7 above and enable the Google Apps Script API.
11. Open the spreadsheet that you used in step 5.i, and open the script editor by going to Tools --> Script editor, which will open in a new tab.
12. The project will be called `Untitled project`, so rename it to something more useful (e.g. the name of the spreadsheet with `Scripts` on the end).
13. Create script files for [`ledger-comparison.gs`](google-apps-scripts/the-new-ledger/ledger-comparison.gs), [`ledger-checker.gs`](google-apps-scripts/the-new-ledger/ledger-checker.gs), and [`ledger-checker-classes.gs`](google-apps-scripts/the-new-ledger/ledger-checker-classes.gs) in the editor, copy their contents in, and save them.
14. Click `Project Settings` on the left and check `Show "appsscript.json" manifest file in editor`.
15. Scroll down. Click `Change project` to change the Google Cloud Platform project from the default to the project you created in step 6 using the project number (which you can find [here](https://console.cloud.google.com/home/dashboard)).
16. Go back to the editor, and replace the content of the default `appsscript.json` file with that from [`appsscript.json`](google-apps-scripts/the-new-ledger/appsscript.json) and save it.
17. Click the blue `DEPLOY` button top-right, and click `New deployment`.
18. Click the cog next to `Select type` and make sure only `API executable` is ticked.
    1. Enter a brief description, allow anyone with a Google Account to use it, and click `Deploy`. 
    2. Copy the deployment ID and paste it into `config.ini`.
19. Fill out the `[ledger_checker]` and `[email]` sections of the config. 
    1. To compare the ledger, you need to have an older version of it in another spreadsheet. 
        1. You can create this using [`ledger_fetcher.py`](python-scripts/ledger_fetcher.py) and copying it to another Google Sheet file. 
	      2. The ID of this Google Sheet file should be added to the config under `compare_spreadsheet_id`, and the name of the sheet (the specific sheet within the spreadsheet, not the spreadsheet file itself) added under `compare_sheet_name`.
    2. You can use the same values as in `[ledger_fetcher]` where duplicates exist. 

You can run either of them from a terminal window within the [`python-scripts`](python-scripts) folder using `python ledger_fetcher.py` or `python_checker.py`. The first time you do, you'll be asked to authorise access to your Google Account. You can safely skip past the `Google hasn’t verified this app` warning screen, because you created the app in step 6, so this isn't a problem.


## Google Apps Scripts

### [The New Ledger](google-apps-scripts/the-new-ledger)
To use the scripts in [`google-apps-scripts/the-new-ledger`](google-apps-scripts/the-new-ledger) to format the ledger neatly, highlight new entries, and copy it to another sheet:
1. Follow steps 6 to 8 and 10 to 12 above. You can re-use the same spreadsheet and Apps Script project.
2. Create script files for [`addon-cards.gs`](google-apps-scripts/the-new-ledger/addon-cards.gs), [`addon-main.gs`](google-apps-scripts/the-new-ledger/addon-main.gs), [`ledger-comparison.gs`](google-apps-scripts/the-new-ledger/ledger-comparison.gs), and [`ledger-comparison-menu.gs`](google-apps-scripts/the-new-ledger/ledger-comparison-menu.gs) in the editor, copy their contents in, and save them.
3. Follow steps 14 to 16 above, again, you can reuse the same spreadsheet and Apps Script project.
4. Click the blue `DEPLOY` button top-right, and click `Test deployments`.
5. Make sure `Test latest code` is selected and click `Install`.
6. Open (or refresh) Google Sheets, and the add-on will appear in the sidebar on the right.

To use the add-on and associated scripts:
1. Open it, click `AUTHORISE ACCESS`, and grant the requested permissions.
   1. You can safely skip past the `Google hasn’t verified this app` warning screen, because you created the app in step 1 above, so this isn't a problem.
2. Enter the URL of the spreadsheet that you want to compare with, and click `VALIDATE URL`.
   1. If it's valid, you'll be asked to select a specific sheet from that spreadsheet.
3. Fill in the other preferences, and click `SAVE`.
4. A green **`Success!`** message will appear once everything's valid, otherwise a red **`Uh oh!`** message will appear, and you'll need to fix any errors before continuing.
5. You can now run the scripts themselves using the `Scripts` menu in the Google Sheet (next to `Help`).


### [Society Ledger Downloader Apps Script Add-On](google-apps-scripts/society-ledger-downloader)
To use the scripts in [`google-apps-scripts/society-ledger-downloader`](google-apps-scripts/society-ledger-downloader) to download the ledger to Google Drive:
1. Follow steps 6 to 8 and 10 above (for the Python scripts).
2. Create a new standalone Apps Script project here: [https://script.google.com/home/projects/create](https://script.google.com/home/projects/create).
3. Create script files for [`cards.gs`](google-apps-scripts/society-ledger-downloader/cards.js), [`ledger-fetcher.gs`](google-apps-scripts/society-ledger-downloader/ledger-fetcher.js), and [`main.gs`](google-apps-scripts/society-ledger-downloader/main.js) in the editor, copy their contents in, and save them.
3. Follow steps 14 to 15 above (for the Python scripts).
4. Go back to the editor, and replace the content of the default `appsscript.json` file with that from [`appsscript.json`](google-apps-scripts/society-ledger-downloader/appsscript.json) and save it.
4. Click the blue `DEPLOY` button top-right, and click `Test deployments`.
5. Make sure `Test latest code` is selected and click `Install`.
6. Open (or refresh) Google Drive, and the add-on will appear in the sidebar on the right.

To use the add-on and associated scripts:
1. Open it, click `AUTHORISE ACCESS`, and grant the requested permissions.
   1. You can safely skip past the `Google hasn’t verified this app` warning screen, because you created the app in step 1 above (for the Python scripts), so this isn't a problem.

You can now either:
* Select a folder to save the ledger to it. 
   * You'll need to enter the society group ID and your eXpense365 credentials.
   * You'll also need to authorise access to that specific folder to allow the ledger to be saved to it.
* Select a PDF file to update it with a new copy of the ledger.
   * The new PDF ledger will be saved as a new version of the file - previous versions can be accessed by right-clicking the file and selecting `Manage versions`.
   * You can choose whether to rename the file or not.
   * You'll need to enter the society group ID and your eXpense365 credentials.
   * You'll also need to authorise access to that specific file to allow it to be updated.