/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Represents a single entry on a ledger.
 */
class Entry {

  /**
   * Constructor, sets everything needed.
   *
   * @param {String} costCodeName
   * @param {String} date
   * @param {String} description
   * @param {Number} money
   */
  constructor(costCodeName, date, description, money) {
    Object.assign(this, {costCodeName, date, description, money});
  }

}


/**
 * Represents a cost code on a ledger (but doesn't contain its entries).
 */
class CostCode {

  /**
   * Constructor, sets everything needed.
   * Checks that balance===moneyIn-moneyOut.
   *
   * @param {String} name
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   * @param {Number} balance
   * @param {Number} lastRowNumber
   */
  constructor(name, moneyIn, moneyOut, balance, lastRowNumber) {
    if (balance !== moneyIn - moneyOut) {
      Logger.log(`balance=${balance}; moneyIn=${moneyIn}; moneyOut=${moneyOut}`);
      // throw new Error("The balance is invalid (balance!==moneyIn-moneyOut).");
    }
    Object.assign(this, {name, moneyIn, moneyOut, balance, lastRowNumber});
  }

}


/**
 * Represents the grand total for a cost code.
 */
class GrandTotal {

  /**
   * Constructor, sets everything needed.
   * Checks that balance===balanceBroughtForward+totalIn-totalOut.
   *
   * @param {Number} totalIn
   * @param {Number} totalOut
   * @param {Number} balanceBroughtForward
   * @param {Number} totalBalance
   * @param {Number} lastRowNumber
   */
  constructor(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    if (totalBalance !== balanceBroughtForward + totalIn - totalOut) {
      Logger.log(`totalBalance=${totalBalance}; balanceBroughtForward=${balanceBroughtForward}; totalIn=${totalIn}; totalOut=${totalOut}`);
      // throw new Error("The total balance is invalid (totalBalance!==balanceBroughtForward+totalIn-totalOut).");
    }
    Object.assign(this, {totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber});
  }

}


/**
 * Represents a ledger.
 * It contains the sheet ID, society name, an array of entries, an
 * array of cost codes, and the grand total.
 */
class Ledger {

  /**
   * Constructor, sets the sheetID, creates empty arrays for entries
   * and costCodes, and sets grandTotal as undefined.
   *
   * @param {Number} sheetId the sheet ID that this represents
   */
  constructor(sheetId) {
    this.sheetId = sheetId;
    this.entries = [];
    this.costCodes = [];
    this.grandTotal = undefined;
  }

  /**
   * Calculates a single value for the money based on the money in
   * or money out (which is how the ledger represents it).
   *
   * @param {Number} moneyIn the money in
   * @param {Number} moneyOut the money out
   * @return {Number} the actual change
   */
  static calculateMoney(moneyIn, moneyOut) {

    if (Number(moneyIn) === 0) {
      return -moneyOut;
    } else {
      return moneyIn;
    }

  }

  /**
   * Compares the two ledgers.
   * Returns true if their grand totals have the same totalIn,
   * totalOut, and balanceBroughtForward.
   *
   * @param {Ledger} ledgerA one ledger.
   * @param {Ledger} ledgerB another ledger.
   * @return {boolean} true if the same, otherwise false.
   */
  static compareLedgers(ledgerA, ledgerB) {

    return ledgerA.grandTotal.totalIn === ledgerB.grandTotal.totalIn &&
      ledgerA.grandTotal.totalOut === ledgerB.grandTotal.totalOut &&
      ledgerA.grandTotal.balanceBroughtForward === ledgerB.grandTotal.balanceBroughtForward;

  }

  /**
   * Adds a new entry on this ledger.
   *
   * @param {String} costCodeName
   * @param {String} date
   * @param {String} description
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   */
  addEntry(costCodeName, date, description, moneyIn = undefined, moneyOut = undefined) {
    const money = Ledger.calculateMoney(moneyIn, moneyOut);
    this.entries.push(new Entry(costCodeName, date, description, money));
  }

  /**
   * Adds a new cost code.
   *
   * @param {String} name
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   * @param {Number} balance
   * @param {Number} lastRowNumber
   */
  addCostCode(name, moneyIn, moneyOut, balance, lastRowNumber) {
    this.costCodes.push(new CostCode(name, moneyIn, moneyOut, balance, lastRowNumber));
  }

  /**
   * Sets the name of the society.
   *
   * @param {String} name
   */
  setSocietyName(name) {
    this.societyName = name;
  }

  /**
   * Sets the grand total. There can only be one grand total.
   *
   * @param {Number} totalIn
   * @param {Number} totalOut
   * @param {Number} balanceBroughtForward
   * @param {Number} totalBalance
   * @param {Number} lastRowNumber
   */
  setGrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    this.grandTotal = new GrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber);
  }

  /**
   * Returns each cost code with its last row number.
   * It's returned as an array of cost codes, with each element as
   * [name, lastRowNumber].
   *
   * @return {[{String}|{Number}]} the array of cost codes.
   */
  getCostCodeRows() {

    const costCodeRows = [];
    for (let i = 0; i < this.costCodes.length; i++) {
      costCodeRows.push([this.costCodes[i].name,
        this.costCodes[i].lastRowNumber]);
    }

    return costCodeRows;

  }

  /**
   * Saves the whole thing to the log.
   */
  log() {
    Logger.log(`The Ledger object is: ${this}`);
  }
}


/**
 * Checks for any new entries in the sheet compared to Original, and
 * returns these as a Ledger object.
 *
 * @param {String} sheetName the name of the newer sheet.
 * @return {Ledger|String} the Ledger object, or "False" if no changes
 * were found.
 */
function checkForNewTotals(sheetName) {

  // Get the spreadsheet & sheet
  const spreadsheet = SpreadsheetApp.getActive();
  const newSheet = spreadsheet.getSheetByName(sheetName);
  Logger.log(`Looking at the sheet called ${newSheet.getName()}`);

  // Create the ledger
  let ledger = new Ledger(newSheet.getSheetId());

  // Find the income & expenditure for each cost code
  ledger = getCostCodeTotals(newSheet, ledger);

  // Format the sheet neatly, and rename it to reflect that this is automated
  formatNeatlyWithSheet(newSheet);
  newSheet.setName(`${newSheet.getName()} auto`);

  // Locate the named range with the URL to the old sheet
  const namedRanges = spreadsheet.getNamedRanges();
  let url;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === "DefaultUrl") {
      url = namedRanges[i].getRange().getValue();
    }
  }

  // Find the totals in the Original (the one that we are comparing against)
  const oldSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log(`Script has opened spreadsheet ${url}`);
  const oldSheet = oldSpreadsheet.getSheetByName("Original");
  let oldLedger = new Ledger(oldSheet.getSheetId());
  oldLedger = getCostCodeTotals(oldSheet, oldLedger);

  // If they're equal then stop and delete the sheet
  if (Ledger.compareLedgers(ledger, oldLedger) === true) {
    Logger.log("There is no difference in the total income, expenditure, or balance brought forward.");
    spreadsheet.deleteSheet(newSheet);
    return "False";
  }

  // If there is a difference then find the new entries and return the Ledger
  Logger.log("There is a difference in the total income and/or expenditure!");
  ledger = findNewEntries(newSheet, oldSheet, ledger);
  return ledger;

}


/**
 * Finds any entries in newSheet that aren't in oldSheet, and saves
 * them to newLedger.
 * These entries will be highlighted in orange on the spreadsheet
 * newLedger must have the cost codes saved to it.
 *
 * @param {Sheet} newSheet the newer sheet.
 * @param {Sheet} oldSheet the older sheet.
 * @param {Ledger} newLedger the Ledger object to save the entries to.
 * @return {Ledger} newLedger with the new entries added.
 */
function findNewEntries(newSheet, oldSheet, newLedger) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  const oldSheetValues = oldSheet.getSheetValues(1, 1, oldSheet.getLastRow(), 4);

  let passedHeader = false;
  let cell;
  let cellValue;
  const costCodeRows = newLedger.getCostCodeRows();
  for (let row = 1; row <= newSheet.getLastRow(); row += 1) {
    cell = newSheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't passed the first header row then skip it
    if (passedHeader === false) {
      if (cellValue === "Date") {
        passedHeader = true;
      }

    } else {
      // Compare it with the original/old sheet
      // Comparing all rows allows us to identify changes in the totals too
      const isNew = compareWithOld(row, oldSheetValues, newSheet);

      // If it is a new row and has a date then save it with its cost code
      if (isNew && isADate(newSheet.getRange(row, 1).getValue())) {
        Logger.log(`Row ${row} is a new row!`);
        newSheet.getRange(row, 1, 1, 4).setBackground("#FFA500");

        // Identify the relevant cost code and save it
        for (let i = 0; i < costCodeRows.length; i++) {
          if (row < costCodeRows[i][1]) {
            newLedger.addEntry(costCodeRows[i][0],
              newSheet.getRange(row, 1).getValue(),
              newSheet.getRange(row, 2).getValue(),
              Number(newSheet.getRange(row, 3).getValue()),
              Number(newSheet.getRange(row, 4).getValue()));
            break;
          }
        }
      }
    }
  }

  Logger.log("Finished comparing sheets!");
  newLedger.log();
  return newLedger;

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
  ledger.setSocietyName(String(foundRanges[foundRanges.length - 1].getValue()).replace("Totals for ", ""));

  ledger.log();
  return ledger;

}
