/**
 * Run formatNeatly() with the saved data.
 */
function formatNeatlyMenu() {
  formatNeatly(SpreadsheetApp.getActiveSheet(),
    getUserProperties().formatSheetName);

  SpreadsheetApp.getActiveSpreadsheet().toast("Formatting complete!");
}


/**
 * Run compareLedgers() with the saved data.
 */
function compareLedgersMenu() {

  const compareSpreadsheet = SpreadsheetApp.openByUrl(getUserProperties().sheetURL);

  compareLedgers(SpreadsheetApp.getActiveSheet(),
    compareSpreadsheet.getSheetByName(getUserProperties().sheetName),
    true,
    getUserProperties().highlightNewRowColour,
    null);

  SpreadsheetApp.getActiveSpreadsheet().toast("Highlighting complete!");
}


/**
 * Run copyToLedger with the saved data.
 */
function copyToLedgerMenu() {

  const compareSpreadsheet = SpreadsheetApp.openByUrl(getUserProperties().sheetURL);

  if (getUserProperties().copyNewSheetName === "") {
    copyToLedger(SpreadsheetApp.getActiveSheet(),
      compareSpreadsheet,
      compareSpreadsheet.getSheetByName(getUserProperties().sheetName));

  } else {
    copyToLedger(SpreadsheetApp.getActiveSheet(),
      compareSpreadsheet,
      getUserProperties().copyNewSheetName);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Formatting complete!");

}


/**
 * Does everything at once: formatting, highlighting, and copying.
 */
function processEntirelyMenu() {

  formatNeatlyMenu();
  compareLedgersMenu();
  copyToLedgerMenu();

}


/**
 * Creates the menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts")
    .addItem("Format neatly", "formatNeatlyMenu")
    .addItem("Highlight new entries", "compareLedgersMenu")
    .addItem("Copy to the ledger", "copyToLedgerMenu")
    .addSeparator()
    .addItem("Format, highlight, and copy", "processEntirelyMenu")
    .addToUi();
}