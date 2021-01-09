/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


class Entry {

  constructor(costCodeName, date, description, money) {
    Object.assign(this, {costCodeName, date, description, money});
  }

}


class CostCode {

  constructor(name, moneyIn, moneyOut, balance, lastRowNumber) {
    if (balance !== moneyIn - moneyOut) {
      throw new Error("The balance is invalid (balance!==moneyIn-moneyOut).");
    }
    Object.assign(this, {name, moneyIn, moneyOut, balance, lastRowNumber});
  }

}


class GrandTotal {

  constructor(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    if (totalBalance !== balanceBroughtForward + totalIn - totalOut) {
      throw new Error("The total balance is invalid (totalBalance!==balanceBroughtForward+totalIn-totalOut).");
    }
    Object.assign(this, {totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber});
  }

}


class Changes {

  constructor(sheetId) {
    this.sheetId = sheetId;
    this.entries = [];
    this.costCodes = [];
    this.grandTotal = undefined;
  }

  static calculateMoney(moneyIn, moneyOut) {

    if (moneyIn !== undefined && moneyOut !== undefined) {
      throw new Error("Only moneyIn or moneyOut can be defined, not both.")
    } else if (moneyOut !== undefined) {
      return -moneyOut;
    } else {
      return moneyIn;
    }

  }

  addEntry(costCodeName, date, description, moneyIn = undefined, moneyOut = undefined) {
    const money = Changes.calculateMoney(moneyIn, moneyOut);
    this.entries.push(new Entry(costCodeName, date, description, money));
  }

  addCostCode(name, moneyIn, moneyOut, balance, lastRowNumber) {
    this.costCodes.push(new CostCode(name, moneyIn, moneyOut, balance, lastRowNumber));
  }

  setSocietyName(name) {
    this.societyName = name;
  }

  setGrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    this.grandTotal = new GrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber);
  }

  log() {
    Logger.log(`The Changes object is: ${this}`);
  }

}


/**
 * This function checks if there are any new entries in the sheet,
 * and if so returns a list of them all, tagged with their cost code.
 * It also returns the new sheet ID and the cost code totals.
 *
 * The returned array (changes) has the structure:
 * [sheetId,
 *  [Entry cost code, Entry date, Entry description, £in, £out],
 *  [Entry cost code, Entry date, Entry description, £in, £out],
 *  [Entry cost code, Entry date, Entry description, £in, £out],
 *  [[Cost code 1, £in, £out, £balance, lastRowNumber],
 *   [Cost code 2, £in, £out, £balance, lastRowNumber],
 *   [Cost code 3, £in, £out, £balance, lastRowNumber],
 *   [Society name, £totalIn, £totalOut, £totalBalance, lastRowNumber, balanceBroughtForward]]]
 */
function checkForNewTotals(sheetName) {

  // Get the spreadsheet & sheet
  const spreadsheet = SpreadsheetApp.getActive();
  const newSheet = spreadsheet.getSheetByName(sheetName);
  Logger.log(`Looking at the sheet called ${newSheet.getName()}`);

  // Create changes
  let changes = new Changes(newSheet.getSheetId());

  // Find the income & expenditure for each cost code
  changes = getCostCodeTotals(newSheet, changes);

  // Format the sheet neatly, and rename it to reflect that this is automated
  formatNeatlyWithSheet(newSheet);
  newSheet.setName(`${newSheet.getName()} auto`);

  // Locate the named range with the URL to the old sheet
  const namedRanges = spreadsheet.getNamedRanges();
  let url;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() == "DefaultUrl") {
      url = namedRanges[i].getRange().getValue();
    }
  }

  // Find the total income and expense in the Original (the one that we are comparing against)
  const oldSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log(`Script has opened spreadsheet ${url}`);
  const oldSheet = oldSpreadsheet.getSheetByName("Original");
  const oldCostCodeTotals = getCostCodeTotals(oldSheet);

  // If there's no difference then stop and delete the sheet
  if (newCostCodeTotals[newCostCodeTotals.length - 1][1] == oldCostCodeTotals[oldCostCodeTotals.length - 1][1] &&
    newCostCodeTotals[newCostCodeTotals.length - 1][2] == oldCostCodeTotals[oldCostCodeTotals.length - 1][2]) {
    Logger.log("There is no difference in the total income or expenditure.");
    spreadsheet.deleteSheet(newSheet);
    return "False";
  }

  // If there is a difference then make comparisons and return the changes
  Logger.log("There is a difference in the total income and/or expenditure!");
  const changeso = compareLedgersWithCostCodes(newSheet, oldSheet, newCostCodeTotals);
  changeso.unshift(newSheet.getSheetId());
  changeso.push(newCostCodeTotals);
  Logger.log(`changes is: ${changeso}`);
  return changeso;

}


/**
 * This function is used to search for changes in the newSheet compared
 * with the oldSheet (not vice-versa). It will categorise them by cost
 * code and return them.
 *
 * The returned array (changes) has the structure:
 * [[Entry cost code, Entry date, Entry description, £in, £out],
 *  [Entry cost code, Entry date, Entry description, £in, £out],
 *  [Entry cost code, Entry date, Entry description, £in, £out]]
 */
function compareLedgersWithCostCodes(newSheet, oldSheet, costCodes) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  const oldSheetValues = oldSheet.getSheetValues(1, 1, oldSheet.getLastRow(), 4);

  let passedHeader = false;
  let cell;
  let cellValue;
  const changes = [];
  for (let row = 1; row <= newSheet.getLastRow(); row += 1) {
    cell = newSheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't passed the first header row then skip it
    if (passedHeader == false) {
      if (cellValue == "Date") {
        passedHeader = true;
      }


      // Compare it with the original/old sheet
      // Comparing all rows allows us to identify changes in the totals too
    } else {
      const isNew = compareWithOld(row, oldSheetValues, newSheet);

      // If it is a new row and has a date then save it with its cost code
      if (isNew && isADate(newSheet.getRange(row, 1).getValue())) {
        Logger.log(`Row ${row} is a new row!`);
        newSheet.getRange(row, 1, 1, 4).setBackground("#FFA500");

        // Identify the relevant cost code and save it
        for (let i = 0; i < costCodes.length; i++) {
          if (row < costCodes[i][4]) {
            changes.push([costCodes[i][0],
              newSheet.getRange(row, 1).getValue(),
              newSheet.getRange(row, 2).getValue(),
              newSheet.getRange(row, 3).getValue(),
              newSheet.getRange(row, 4).getValue()]);
            break;
          }
        }
      }
    }
  }
  Logger.log("Finished comparing sheets!");
  Logger.log(`changes is: ${changes}`);
  return changes;

}


/**
 * This function retrieves the total income, expenditure, and balance for
 * each cost code, as well as the grand total for the entire ledger.
 *
 * The returned array (costCodeTotals) has the structure:
 * [[Cost code 1, £in, £out, £balance, lastRowNumber],
 *  [Cost code 2, £in, £out, £balance, lastRowNumber],
 *  [Cost code 3, £in, £out, £balance, lastRowNumber],
 *  [Society name, £totalIn, £totalOut, £totalBalance, lastRowNumber, balanceBroughtForward]]
 */
function getCostCodeTotals(sheet, changes) {

  const costCodeTotals = [];
  let costCode;

  // Search for the total for each cost code (but not the grand total)
  const finder = sheet.createTextFinder("Totals for ").matchEntireCell(false);
  const foundRanges = finder.findAll();
  for (let i = 0; i < foundRanges.length-1; i++) {

    // Get the name of the cost code
    costCode = String(foundRanges[i].getValue()).replace("Totals for ", "");

    // Add this to changes
    changes.addCostCode(costCode,
      foundRanges[i].offset(0, 1).getValue(),
      foundRanges[i].offset(0, 2).getValue(),
      foundRanges[i].offset(1, 2).getValue(),
      foundRanges[i].getRow())
  }

  // Save the grand total
  changes.setGrandTotal(foundRanges.slice(-1).offset(0, 1).getValue(),
    foundRanges.slice(-1).offset(0, 2).getValue(),
    foundRanges.slice(-1).offset(2, 2).getValue(),
    foundRanges.slice(-1).offset(3, 2).getValue(),
    foundRanges.slice(-1).getRow())

  changes.log();
  return changes;

}
