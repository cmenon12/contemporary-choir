/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Checks for any new entries in sheetName compared to the other sheet,
 * and returns these as a JSON Ledger object.
 *
 * @param {String} sheetName the name of the newer sheet.
 * @param {String} compareSheetId the ID of the sheet to compare with
 * @param {String} compareSheetName the name of the sheet to compare with
 * @return {Ledger|String} the Ledger object, or "False" if no changes
 * were found.
 */
function checkForNewTotals(sheetName, compareSheetId, compareSheetName) {

  // Get the spreadsheet & sheet
  const spreadsheet = SpreadsheetApp.getActive();
  const newSheet = spreadsheet.getSheetByName(sheetName);
  Logger.log(`Looking at the sheet called ${newSheet.getName()}`);

  // Create the ledger
  let ledger = new Ledger(newSheet.getSheetId());

  // Format the sheet neatly
  formatNeatly(newSheet, "{{datetime}} auto");

  // Find the income & expenditure for each cost code
  ledger = getCostCodeTotals(newSheet, ledger);

  // Find the totals in the sheet that we are comparing against
  const oldSpreadsheet = SpreadsheetApp.openById(compareSheetId);
  Logger.log(`Script has opened spreadsheet ${compareSheetId}`);
  const oldSheet = oldSpreadsheet.getSheetByName(compareSheetName);
  let oldLedger = new Ledger(oldSheet.getSheetId());
  oldLedger = getCostCodeTotals(oldSheet, oldLedger);

  // Save timestamp of the old ledger to the new ledger
  ledger.setOldLedgerTimestamp(oldSheet.getRange("A3").getValue());

  // If they're equal then stop
  if (Ledger.compareLedgers(ledger, oldLedger) === true) {
    Logger.log("There is no difference in the total income, expenditure, or balance brought forward.");
    return "False";
  }

  // If there is a difference then find the new entries and return the Ledger
  Logger.log("There is a difference in the total income and/or expenditure!");
  ledger = compareLedgers(newSheet, oldSheet, "orange", false, ledger);
  return JSON.stringify(ledger);

}


/**
 * This function retrieves the total income, expenditure, and balance for
 * each cost code, as well as the grand totals for the entire ledger.
 * It saves these to the Ledger object, which it then returns.
 * It also sets the society name.
 *
 * @param {Sheet} sheet the sheet to search.
 * @param {Ledger} ledger the Ledger object to update.
 * @returns {Ledger} the updated Ledger.
 */
function getCostCodeTotals(sheet, ledger) {

  // Search for the total for each cost code (but not the grand total)
  let costCode;
  const finder = sheet.createTextFinder("Totals for ").matchEntireCell(false);
  const foundRanges = finder.findAll();
  for (let i = 0; i < foundRanges.length - 1; i++) {

    // Get the name of the cost code
    costCode = String(foundRanges[i].getValue()).replace("Totals for ", "");

    // Add this to the Ledger object
    ledger.addCostCode(costCode,
      Number(foundRanges[i].offset(0, 1).getValue()),
      Number(foundRanges[i].offset(0, 2).getValue()),
      Number(foundRanges[i].offset(1, 2).getValue()),
      foundRanges[i].getRow());
  }

  // Save the grand total
  ledger.setGrandTotal(Number(foundRanges[foundRanges.length - 1].offset(0, 1).getValue()),
    Number(foundRanges[foundRanges.length - 1].offset(0, 2).getValue()),
    Number(foundRanges[foundRanges.length - 1].offset(2, 2).getValue()),
    Number(foundRanges[foundRanges.length - 1].offset(3, 2).getValue()),
    foundRanges[foundRanges.length - 1].getRow());

  // Save the society name
  ledger.setSocietyName(sheet.getRange("A2").getValue());

  ledger.log();
  return ledger;

}
