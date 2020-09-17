# Contemporary Choir

As the Treasurer of [Contemporary Choir](https://exetercontemporarychoir.com/) (a University of Exeter Students' Guild society), I've found myself spending quite a bit of time working with various spreadsheets and data from various different sources. As a result, I've started writing scripts to help automate some tasks and ultimately make my life easier. 
<br><br>In addition, I've also led the team to produce a [brand-new website for the society](https://exetercontemporarychoir.com), during which I've been able to draw on and apply skills learnt during my degree as well as learn some new skills (such as domain management and SEO) too. This repo currently has a plugin that I've tweaked myself especially for our website.
<br><br>*Please note that this repository is managed by me personally as an individual, and not by Contemporary Choir.*

[![GitHub issues](https://img.shields.io/github/issues/cmenon12/contemporary-choir?style=flat)](https://github.com/cmenon12/contemporary-choir/issues)
[![GitHub license](https://img.shields.io/github/license/cmenon12/contemporary-choir?style=flat)](https://github.com/cmenon12/contemporary-choir/blob/master/LICENSE)


## Getting Started
For the Python Scripts and the Google Apps Scripts:
* Rename [`python-scripts/config-template.ini`](python-scripts/config-template.ini) to `python-scripts/config.ini` and update the values within it with your own data.
* Install the required packages using `pip install -r requirements.txt`.
* Create your own Google Cloud Platform Project, enable the Apps Script, Drive, and Sheets APIs, and create & download some OAuth 2.0 Client ID credentials as `credentials.json`. 
  * You mustn't share `credentials.json` or the generated tokens with anyone.
  * Use of these APIs is (at the time of writing) free.
* Add the Apps Scripts to an Apps Script project created within Google Sheets.
  * Follow steps 1 to 3 of [these instructions](https://developers.google.com/apps-script/api/how-tos/execute#general_procedure) to allow the functions to be executed by the API. This is only necessary to run `ledger_checker.py`.


## Python Scripts

**[`python-scripts/ledger_fetcher.py`](python-scripts/ledger_fetcher.py)** 
can be used  to download the society ledger from eXpense365 to your computer (instead of having to use the app). It can also then convert it from a PDF to an XLSX spreadsheet (using [pdftoexcel.com](https://www.pdftoexcel.com/)) and then upload the newly-converted spreadsheet to a Google Sheet (as a new sheet within a pre-existing spreadsheet). *Please note that I'm not affiliated with pdftoexcel.com and that use of their service is bound by their terms & privacy policy - it's just a handy service that I've found can convert the ledger accurately.*


**[`python-scripts/ledger_checker.py`](python-scripts/ledger_checker.py)** is designed to check the ledger on a regular basis and notify the user via email of any changes. It relies on [`ledger_fetcher.py`](python-scripts/ledger_fetcher.py) to download the ledger, convert it, and upload it to Google Sheets. It will then run an Apps Script function (namely `checkForNewTotals(sheetName)` in [`ledger_checker.gs`](apps-script/ledger_checker.gs)) to identify any changes. If it does identify any changes then it will email these to the user along with the PDF ledger itself (check [here](https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/Example%20email%20from%20ledger_checker.py.jpg) for an example). The user is only ever notified of each change once by serialising them to a file that maintains persistence.

## Google Apps Scripts (for Google Sheets)
**[`apps-script/ledger-comparison.gs`](google-apps-scripts/ledger-comparison.gs)** can be used to process the ledger that has been uploaded by [`ledger_fetcher.py`](python-scripts/ledger_fetcher.py). 
* `formatNeatlyWithSheet(sheet)` will format the ledger neatly by renaming the sheet, resizing the columns, removing unnecessary headers, and removing excess columns & rows.
  * `formatNeatly()` is the same as this, except it will use the active sheet in the active spreadsheet.
* `compareLedgersGetUrl()` and `compareLedgers(url)` will compare the ledger with that in the Google Sheet at a given URL. The sheet at the URL must be an older version named `Original`. Any new or differing entries in the newer version will be highlighted in red. 
* `copyToLedgerGetUrl()` and `copyToLedger(url)` will copy the ledger to the Google Sheet at the given URL. The function will replace the sheet called `Original` at this URL.
* `processWithDefaultUrl()` will do all of the above at once using the URL in the named range named `DefaultUrl`.

**[`apps-script/ledger-checker.gs`](google-apps-scripts/ledger-checker.gs)** is used by [`ledger_checker.py`](python-scripts/ledger_checker.py) to identify any changes in the uploaded ledger. It's designed to be executed by the API and is therefore not reliant on determining the active sheet.
* `checkForNewTotals(sheetName)` will check the named sheet in the linked spreadsheet for any new changes compared with the Google Sheet at the default URL called `Original`. If any are found then it will return them along with the current total for each cost code, otherwise it will return `"False"`.
* `compareLedgersWithCostCodes(newSheet, oldSheet, costCodes)` will search for changes in the `newSheet` compared with the `oldSheet` (not vice-versa). It will categorise them by cost code and return them.
* `getCostCodeTotals(sheet)` will retrieve the total income, expenditure, and balance for each cost code, as well as the grand total for the entire ledger.

**[`apps-script/fundraising.gs`](google-apps-scripts/fundraising.gs)** updates how much has been fundraised from a GoFundMe page. It fetches the page, extracts the total fundraised and the total number of donors, applies a reduction due to payment processor fees & postage, and then updates a pre-defined named range in the sheet with the total.

## Wordpress Plugins 
**[`wordpress-plugins/password-protected/`](wordpress-plugins/password-protected/)** is my own customised version of [Password Protected by Ben Huson](https://wordpress.org/plugins/password-protected/). I have modified it to include a Google CAPTCHA on the password page, which is implemented using the plugin [Advanced noCaptcha & invisible Captcha (v2 & v3) by Shamim Hasan](https://wordpress.org/plugins/advanced-nocaptcha-recaptcha/). It also has a custom logo (instead of the Wordpress one) and some brief text for people arriving at the site.

## License
[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)


