/**
 * Run formatNeatly() with the saved data.
 */
function formatNeatlyMenu() {

}


/**
 * Run compareLedgers() with the saved data.
 */
function compareLedgersMenu() {

}


/**
 * Run copyToLedger with the saved data.
 */
function copyToLedgerMenu() {

}


/**
 * Does everything at once: formatting, highlighting, and copying.
 */
function processEntirely() {

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