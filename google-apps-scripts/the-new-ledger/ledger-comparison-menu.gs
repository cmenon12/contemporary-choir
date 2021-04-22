/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Run formatNeatly() with the saved data.
 */
function formatNeatlyMenu() {
  if (validatePreferences(getUserProperties())) {
    formatNeatly(SpreadsheetApp.getActiveSheet(),
      getUserProperties().formatSheetName);

    SpreadsheetApp.getActiveSpreadsheet().toast("Formatting complete!");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("You need to set your preferences in the add-on first.");
  }
}


/**
 * Run compareLedgers() with the saved data.
 */
function compareLedgersMenu() {

  if (validatePreferences(getUserProperties())) {
    const compareSpreadsheet = SpreadsheetApp.openByUrl(getUserProperties().sheetURL);

    compareLedgers(SpreadsheetApp.getActiveSheet(),
      compareSpreadsheet.getSheetByName(getUserProperties().sheetName),
      true,
      getUserProperties().highlightNewRowColour,
      null);

    SpreadsheetApp.getActiveSpreadsheet().toast("Highlighting complete!");

  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("You need to set your preferences in the add-on first.");
  }
}


/**
 * Run copyToLedger with the saved data.
 */
function copyToLedgerMenu() {

  if (validatePreferences(getUserProperties())) {

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

  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("You need to set your preferences in the add-on first.");
  }

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