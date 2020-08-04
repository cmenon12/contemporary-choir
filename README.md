# Contemporary Choir
As the Treasurer of [Contemporary Choir](https://www.exeterguild.org/societies/contemporarychoir/) (a University of Exeter Students' Guild society), I've found myself spending quite a bit of time working with various spreadsheets and data from various different sources. As a result, I've started writing scripts to help automate some tasks and ultimately make my life easier.

## Python Scripts
**[`ledger_fetcher.py`](ledger_fetcher.py)** 
can be used  to download the society ledger from eXpense365 to your computer (instead of having to use the app). It can also then convert it from a PDF to an XLSX spreadsheet (using [pdftoexcel.com](https://www.pdftoexcel.com/)) and then upload the newly-converted spreadsheet to a pre-existing Google Sheet (as a new sheet within a spreadsheet). *Please note that I'm not affiliated with pdftoexcel.com and that use of their service is bound by their terms & privacy policy - it's just a handy service that I've found can convert the ledger accurately.*

## Google Apps Scripts (for Google Sheets)
**[`apps-script/ledger-comparison.gs`](apps-scripts/ledger-comparison.gs)** can be used to process the ledger that has been uploaded by `ledger_fetcher.py`. 
* `formatNeatly()` will format the ledger neatly by renaming the sheet, resizing the columns, removing unnecesary headers, and removing excess columns & rows.
* `compareLedgersGetUrl()` and `compareLedgers(url)` will compare the ledger with that in the Google Sheet at a given URL. The sheet at the URL must be an older version named `Original`. Any new or differing entries in the newer version will be highlighted in red. 
* `copyToLedgerGetUrl()` and `copyToLedger(url)` will copy the ledger to the Google Sheet at the given URL. The function will replace the sheet called `Original` at this URL.
* `processWithDefaultUrl()` will do all of the above at once using the URL in the named range named `DefaultUrl`.


**[`apps-script/fundraising.gs`](apps-scripts/fundraising.gs)** updates how much has been fundraised from a GoFundMe page. It fetches the page, extracts the total fundraised and the total number of donors, applies a reduction due to payment processor fees & postage, and then updates a pre-defined named range in the sheet with the total.

## License
[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)


