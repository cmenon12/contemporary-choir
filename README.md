# Contemporary Choir

When I was the Treasurer of [Contemporary Choir](https://exetercontemporarychoir.com/) (a University of Exeter Students' Guild society), I found myself spending quite a bit of time working with various spreadsheets and data from various different sources. As a result, I wrote scripts to help automate some tasks and ultimately make my life easier. In addition, I also led the team to produce a [website for the society](https://exetercontemporarychoir.com), during which I was able to draw on and apply skills learnt during my degree as well as learn some new skills (such as domain management and SEO) too. 

This repository contains scripts that I wrote to automatically download and process our society ledger, and notify me of any new changes via email. It's also got some custom HTML that's used on the website, as well as a plugin that I've tweaked myself for the website.

*Please note that this repository is managed by me personally as an individual, and not by Contemporary Choir.*

*Icon made by [photo3idea_studio](https://www.flaticon.com/authors/photo3idea-studio)
from [www.flaticon.com](https://www.flaticon.com/).*

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.4725442.svg)](https://doi.org/10.5281/zenodo.4725442)
[![GitHub license](https://img.shields.io/github/license/cmenon12/contemporary-choir?style=flat)](https://github.com/cmenon12/contemporary-choir/blob/master/LICENSE)

## You can find instructions for getting started in [`INSTRUCTIONS.md`](society-ledger-manager/INSTRUCTIONS.md).


## Python Scripts
**[`requirements.txt`](society-ledger-manager/python-scripts/requirements.txt)** contains all the requirements for both of these scripts.

### [`ledger_fetcher.py`](society-ledger-manager/python-scripts/ledger_fetcher.py)
This can download the society ledger from eXpense365 to your computer (instead of having to use the app), convert it to an XLSX, and upload it to Google Drive.

#### Classes
* **`CustomEncoder`** is a custom JSON encoder that's used for logging. For bytes objects, it returns a string with their length instead of the actual bytes. For all other objects it uses the default JSON encoder, falling back on the built-in `str(obj)` method where needed.
* **`Ledger`** represents a ledger, which is downloaded upon instantiation. It includes methods to convert it to an XLSX, upload the PDF or XLSX to Drive, save or delete the PDF or XLSX file to the local filesystem, open the PDF or uploaded Google Sheet in the web browser, and refresh the ledger to a more up-to-date version. It also has numerous getter methods that incorporate these methods as required.

#### Functions
* **`authorize()`** is used to authorize access to Google Drive, Sheets, and Apps Script, returning the three services in a tuple. It has an inner function, `authorize_in_browser()` that handles the authorization flow by opening the browser if requested by the user and timing out after 300 seconds.
* **`push_url()`** is used to push a given URL via Pushbullet to the user's device(s). This is used by the `authorize()` function to send them the URL if they don't want to open the browser on the current device. Note that it catches all Pushbullet-related exceptions, because this functionality is not required for the rest of the program to function.
* **`main()`** downloads the ledger from eXpense365 and gives the user the option to open it in the browser. It also gives the user the option to convert it to an XLSX and upload it to Google Sheets.


### [`ledger_checker.py`](society-ledger-manager/python-scripts/ledger_checker.py)
This checks the ledger and notifies the user via email of any changes.
* An example email can be found [here](https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/Example%20email%20from%20ledger_checker.py.jpg). 
* The user is only ever notified of each change once by serialising them to a file that maintains persistence.
* The user is also notified via email if multiple consecutive exceptions occur. 
* The program is designed to be run multiple times via `cron` or a similar tool. It can also be run once to make an ad-hoc check.

#### Classes
* **`LedgerCheckerSaveFile`** represents the save file for this script, and is used to maintain persistence between checks. It contains the filepath to the actual file, a list of stacktraces from failed executions, the most recent new changes and the associated `Ledger`, the most recently downloaded `Ledger`, and the ID of the last success and error emails. It includes methods to save this data to a file and update itself after a successful or unsuccessful check. It also has several getter methods which are used by other functions in this script. 

#### Functions
* **`prepare_email_body()`** is used to populate the [email template](society-ledger-manager/python-scripts/email-template.html) with details of the new changes. The HTML template is written with Jinja2 placeholders that this function uses.
* **`send_success_email()`** is used to prepare and send an email about a successful check. It creates the HTML message using the `prepare_email_body()` function, attaches the new PDF ledger (as well as the one from the previous check if available), and sends the message using the `send_email()` function.
* **`send_error_email()`** is used to prepare and send an email after a certain number of consecutive exceptions. This email contains the stacktraces and timestamps of each of these exceptions, the timestamp of the last successful check, and when a future email will be sent if these exceptions continue consecutively. It's written in plain-text instead of HTML to reduce the risk of an exception occurring within this function itself (which would prevent the user being notified). The email is sent using the `send_email()` function.
* **`send_email()`** is used to send an email from `send_success_email()` or `send_error_email()`. It adds the date and a unique ID to the email and sends it via SMTP. It also optionally manually saves it to the Sent folder using IMAP.
* **`check_ledger()`** is used to run a single check of the ledger. It downloads the ledger from eXpense365, converts it to a PDF, and uploads it to Google Sheets. It then executes an Apps Script function (namely `checkForNewTotals(sheetName)` in [`ledger-checker.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-checker.gs))to identify any changes. If it does identify any changes then it will email these to the user along with the PDF ledger itself using `send_success_email()`.
* **`main()`** runs `check_ledger()` and catches any exceptions that occur. It saves them to the save file, and emails the user if multiple consecutive exceptions occur.


### [`custom_exceptions.py`](society-ledger-manager/python-scripts/custom_exceptions.py)
This contains a variety of custom exceptions that the above two scripts can throw.


### [`archive.py`](society-ledger-manager/python-scripts/archive.py)
This contains code that is no longer actively used, but may be of use in the future. Note that its dependencies are not necessarily listed in
[`requirements.txt`](society-ledger-manager/python-scripts/requirements.txt).

#### Classes
* **`PDFToXLSXConverter`** represents an online PDF-to-XLSX converter. It currently supports [pdftoexcel.com](https://www.pdftoexcel.com/) and [pdftoexcelconverter.net](https://www.pdftoexcelconverter.net/), both of which are very similar. A converter is chosen upon instantiation, either randomly or by the user. It includes methods to upload the PDF, check the conversion status, and download the resulting XLSX file.
* **`ConversionTimeoutError`** and **`ConversionRejectedError`** are both exceptions used by `PDFToXLSXConverter`.


## Google Apps Scripts
All of these (except for the [Society Ledger Downloader](#society-ledger-downloader-apps-script-add-on)) are for Google Sheets, and should therefore be created in a [bound Apps Script project](https://developers.google.com/apps-script/guides/bound).

**[`google-apps-scripts/macmillan-fundraising.gs`](fundraising/macmillan-fundraising.gs)** updates how much has been fundraised for Macmillan from a GoFundMe page. It fetches the page, extracts the total fundraised and the total number of donors, applies a reduction due to payment processor fees & postage, and then updates a pre-defined named range in the sheet with the total.

**[`google-apps-scripts/yd-fundraising.gs`](fundraising/yd-fundraising.gs)** updates how much has been fundraised for Young Devon. 
* It currently fetches the totals from multiple Enthuse pages, and updates several named ranges in the sheet with the totals.
* It also calculates the fees that Enthuse charges.
* `EnthuseFundraisingSource` is a class used to represent an Enthuse fundraising source (which can consist of one or more pages). 
  * Each source has an amount, number of donors, and fees, as well as it's named range. 


### [The New Ledger](society-ledger-manager/google-apps-scripts/the-new-ledger)
This folder contains scripts that process the ledger in Google Sheets once it's been uploaded. The scripts must be
created in an Apps Script project that's bound to the relevant Google Sheet.

**[`google-apps-scripts/the-new-ledger/ledger-comparison.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-comparison.gs)** is used to process the ledger that has been uploaded by [`ledger_fetcher.py`](society-ledger-manager/python-scripts/ledger_fetcher.py). It's preferences are set using the Google Workspace add-on defined in [`addon-cards.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/addon-cards.gs) and [`addon-main.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/addon-main.gs).
* **`formatNeatly(thisSheet, sheetName)`** is used to format the ledger neatly by renaming the sheet, resizing the columns, removing unnecessary headers, and removing excess columns & rows.
* **`compareLedgers(newSheet, oldSheet, colourCountdown, newRowColour, newLedger = null)`** is used to compare the ledger in `newSheet` with that in `oldSheet`. Any new or differing entries in the newer version will be highlighted in `newRowColour`. It can optionally save the changes to the `newLedger` object (which is used by [`ledger-checker.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-checker.gs)).
* **`copyToLedger(thisSheet, destSpreadsheet, newSheet)`** is used to copy `thisSheet` to the `destSpreadsheet`, either replacing `newSheet` or renaming it to `newSheet` (if it's a Sheet or String respectively).

**[`google-apps-scripts/the-new-ledger/ledger-comparison-menu.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-comparison-menu.gs)** is used to create the menu within Google Sheets for the user to trigger the functions in [`ledger-comparison.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-comparison.gs). It includes a function to create the menu, as well as functions that the menu items trigger to call the functions in [`ledger-comparison.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-comparison.gs) using the saved data.

**[`google-apps-scripts/the-new-ledger/addon-cards.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/addon-cards.gs)** is used to build the Google Workspace add-on for the user to set their preferences for the functions in [`ledger-comparison.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-comparison.gs). It includes functions to build the homepage card and its constituent sections.

**[`google-apps-scripts/the-new-ledger/addon-main.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/addon-main.gs)** is used as the backend for the Google Workspace add-on. It includes functions to process the submitted form and manage the user's saved data.

**[`google-apps-scripts/the-new-ledger/ledger-checker.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-checker.gs)** is used by [`ledger_checker.py`](society-ledger-manager/python-scripts/ledger_checker.py) to identify any changes in the uploaded ledger. It's designed to be executed by the API and is therefore not reliant on determining the active sheet.
* **`checkForNewTotals(sheetName, compareSheetId, compareSheetName)`** is used to check the named sheet in the linked spreadsheet for any new changes compared the other sheet. If any are found then it will return them along with the current total for each cost code and the grand total, otherwise it will return `"False"`.
* **`getCostCodeTotals(sheet, ledger)`** is used to retrieve the total income, expenditure, and balance for each cost code, as well as the grand total for the entire ledger from `sheet`. It'll add these to `ledger` and return it.

**[`google-apps-scripts/the-new-ledger/ledger-checker-classes.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-checker-classes.gs)** contains classes used by [`ledger-checker.gs`](society-ledger-manager/google-apps-scripts/the-new-ledger/ledger-checker.gs). `Ledger`, `CostCode`, `Entry`, and `GrandTotal` are classes used to represent the ledger, each cost code, each entry within each cost code, and the grand total respectively. Using these classes makes it much easier to handle and process this data, both by the Apps Script and by the Python Script.


### [Society Ledger Downloader Apps Script Add-On](society-ledger-manager/google-apps-scripts/society-ledger-downloader)
This is a Google Workspace add-on that allows you to download your society ledger and save it straight to Drive. *Note that all of these files should have a `.gs` file extension instead of `.js`but [`clasp`](https://developers.google.com/apps-script/guides/clasp) changes this for me.*

## Website Custom HTML
**[`website-custom-html/`](website-custom-html)** contains various snippets of HTML (including CSS and JavaScript) that are used on Contemporary Choir's websites. These are all incorporated using the custom HTML block in WordPress.
* **[`onesignal-subscribe.html`](website-custom-html/onesignal-subscribe.html)** is used to produce an in-text link for the user to subscribe and unsubscribe to OneSignal push notifications. It’s used on several pages.
  * It will show the text “Finally, you can also subscribe to push notifications through your web browser.”, replacing subscribe with unsubscribe if they’re already subscribed. 
  * When the user clicks on the link, the text will change between subscribe and unsubscribe (as they subscribe and unsubscribe).
  * The whole sentence will be completely hidden if they’re using an unsupported browser (e.g. Safari).	
* **[`would-i-lie-to-you.html`](website-custom-html/would-i-lie-to-you.html)** was used for a WILTY social.
* **[`pictionary-generator.html`](website-custom-html/pictionary-generator.html)** was used for a Pictionary social.
* **[`addthis-buttons.html`](website-custom-html/addthis-buttons.html)** adds the AddThis buttons, but also customises them so that they’re bigger and aligned in the centre.
* **[`recent-posts.html`](website-custom-html/recent-posts.html)** rearranges the information in the ‘Recent Posts’ widget above it to look nicer.
* **[`latest-update.html`](website-custom-html/latest-update.html)** is very similar to the one above, except it’s designed for only one post (our latest update), instead of a small list of them.


## WordPress Plugins
**[`wordpress-plugins/password-protected/`](wordpress-plugins/password-protected)** is my own customised version of [Password Protected by Ben Huson](https://wordpress.org/plugins/password-protected/). I have modified it to include a Google CAPTCHA on the password page, which is implemented using the plugin [Advanced noCaptcha & invisible Captcha (v2 & v3) by Shamim Hasan](https://wordpress.org/plugins/advanced-nocaptcha-recaptcha/). It also has a custom logo (instead of the WordPress one) and some brief text for people arriving at the site.

## License
[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)
