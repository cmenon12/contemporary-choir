/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

// The name of the sheet with these named ranges.
const NAMED_RANGES_SHEET_NAME = "Values";

// The name of the sheet with the imported ledger
const LEDGER_SHEET_NAME = "Original";


/**
 * This can be used to delete and then recreate most of the named ranges.
 * It's designed for one-time use after something bad happens.
 */
function recreateNamedRanges() {

  // Delete all of the named ranges on this sheet
  const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  for (let i = 0; i < namedRanges.length; i++) {
    namedRanges[i].remove();
  }

  // Recreate the named ranges in the selection
  const thisRange = SpreadsheetApp.getActiveRange();
  for (let i = 0; i < thisRange.getValues().length; i++) {
    Logger.log(thisRange.getValues()[i][1]);
    SpreadsheetApp.getActiveSpreadsheet()
      .setNamedRange(thisRange.getValues()[i][1], thisRange.getCell(i + 1, 1));
  }
}


/**
 * Locates and returns the named ranges that we'll want to update.
 * These are the named ranges in the Values sheet that will also appear
 * in the PDF ledger.
 *
 * @return {{}} the named ranges as an Object
 */
function getNamedRangesFromSheet() {

  // Define the sheet and get the named ranges.
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(NAMED_RANGES_SHEET_NAME);
  const namedRanges = sheet.getNamedRanges();

  // Prepare to save the values
  const namedRangesObj = {};

  // For each named range in this sheet
  for (let i = 0; i < namedRanges.length; i++) {

    // If the named range name doesn't include Pre,
    // and the A1 notation doesn't include a : (so it's a single cell)
    if (!namedRanges[i].getName().includes("Pre") && !namedRanges[i]
      .getRange().getA1Notation().includes(":")) {

      // Add it to our object, using the name as the key.
      namedRangesObj[namedRanges[i].getName()] = namedRanges[i].getRange();

    }
  }

  Logger.log("We found " + Object.keys(namedRangesObj).length +
    " named ranges.");
  Logger.log(`namedRangesObj is: ${JSON.stringify(namedRangesObj)}`);

  return namedRangesObj;

}


/**
 * Extracts the income, expenditure, and balance for each cost code.
 * This also includes the total income, expenditure, and balance,
 * as well as the balance brought forward.
 * It also groups all of the individual TOMS cost codes together.
 *
 * If addNotesToLedger is true, it will add notes to the totals on the
 * imported ledger, labelling it with the named ranges. This allows you
 * to refer back and see where the program got its values from.
 *
 * @param {Boolean} addNotesToLedger whether to label the ledger with notes.
 * @return {{}} the cost code values.
 */
function getNewCostCodeValues(addNotesToLedger = true) {

  // Define the sheet and get the named ranges.
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(LEDGER_SHEET_NAME);

  // Prepare to save the values
  const costCodeValues = {};
  costCodeValues["TOMSIncome"] = 0;
  costCodeValues["TOMSExpenditure"] = 0;
  costCodeValues["TOMSBalance"] = 0;

  // Define the required variables
  let costCodeName;
  let income;
  let expenditure;
  let balance;
  let balanceBroughtForward;

  // Search for the total for each cost code (and the grand total)
  const finder = sheet.createTextFinder("Totals for ").matchEntireCell(false);
  const foundRanges = finder.findAll();
  for (let i = 0; i < foundRanges.length; i++) {

    // Get the name of the cost code
    costCodeName = String(foundRanges[i].getValue())
      .replace("Totals for ", "");

    // Get these ranges
    income = foundRanges[i].offset(0, 1);
    expenditure = foundRanges[i].offset(0, 2);
    balance = foundRanges[i].offset(1, 2);

    // If it's a TOMS cost code then add it to the existing TOMS values
    // This keeps all of the TOMS grouped together
    if (costCodeName.includes("TOMS")) {
      costCodeValues["TOMSIncome"] += income.getValue();
      costCodeValues["TOMSExpenditure"] += expenditure.getValue();
      costCodeValues["TOMSBalance"] += balance.getValue();

      // Add notes to the imported ledger if requested
      if (addNotesToLedger) {
        income.setNote("Part of TOMSIncome");
        expenditure.setNote("Part of TOMSExpenditure");
        balance.setNote("Part of TOMSBalance");
      }


      // If it's the last element then this must be the grand total
      // This uses the name Total and includes the balance brought forward
    } else if (i === foundRanges.length - 1) {

      // Update these values, specifically for the grand total
      balanceBroughtForward = foundRanges[i].offset(2, 2);
      balance = foundRanges[i].offset(3, 2);

      // Add the values
      costCodeValues["TotalIncome"] = income.getValue();
      costCodeValues["TotalExpenditure"] = expenditure.getValue();
      costCodeValues["BalanceBroughtForward"] = balanceBroughtForward.getValue();
      costCodeValues["TotalBalance"] = balance.getValue();

      // Add notes to the imported ledger if requested
      if (addNotesToLedger) {
        income.setNote("TotalIncome");
        expenditure.setNote("TotalExpenditure");
        balanceBroughtForward.setNote("BalanceBroughtForward");
        balance.setNote("TotalBalance");
      }


      // Otherwise just add the standalone cost code itself
    } else {

      // Remove the spaces from the cost code name
      costCodeName = costCodeName.replace(/\s/g, "");

      // Add the values
      costCodeValues[`${costCodeName}Income`] = income.getValue();
      costCodeValues[`${costCodeName}Expenditure`] = expenditure.getValue();
      costCodeValues[`${costCodeName}Balance`] = balance.getValue();

      // Add notes to the imported ledger if requested
      if (addNotesToLedger) {
        income.setNote(`${costCodeName}Income`);
        expenditure.setNote(`${costCodeName}Expenditure`);
        balance.setNote(`${costCodeName}Balance`);
      }
    }
  }

  Logger.log(`We found ${Object.keys(costCodeValues).length} cost code values`);
  Logger.log(`costCodeValues is: ${JSON.stringify(costCodeValues)}`);

  return costCodeValues;

}


/**
 * Compare the named ranges with the new values, and return those that
 * have changed.
 *
 * @param {{}} namedRanges the named ranges with the values.
 * @param {{}} newValues the new values for those ranges.
 * @return {{}} the ranges that have changed.
 */
function compareNamedRanges(namedRanges, newValues) {

  const changedRanges = {};

  for (let key in namedRanges) {

    // If the value has changed then save it
    if (namedRanges[key].getValue() !== newValues[key]) {

      // This saves [old value, new value, Range]
      changedRanges[key] = [namedRanges[key].getValue(),
        newValues[key], namedRanges[key]];

      Logger.log(key + " has changed: old:" +
        namedRanges[key].getValue() + ", new:" + newValues[key]);
    }
  }

  Logger.log(`There are ${Object.keys(changedRanges).length} ` +
    `values that have changed.`);

  return changedRanges;
}


/**
 * Update the named ranges with the new values.
 * If includeDate is true, then the current date & time will be added
 * as a note to the source.
 *
 * @param {{}} changedRanges the ranges that have changed.
 * @param {Boolean} includeDate whether to add the timestamp as a note.
 */
function updateNamedRanges(changedRanges, includeDate = true) {

  // Get the current date and time
  const now = new Date().toLocaleString("en-GB");

  // For each key in changedRanges
  for (let key in changedRanges) {

    // Set the range to the new value
    changedRanges[key][2].setValue(changedRanges[key][1]);

    // Set the source to reflect that this is an automatic update
    changedRanges[key][2].offset(0, 3).setValue("PDF (auto)");

    // Add the date & time as a note to the source if requested
    if (includeDate) {
      changedRanges[key][2].offset(0, 3).setNote(now);
    }

    Logger.log(`Changed ${key} in ${changedRanges[key][2].getA1Notation()}` +
      ` from ${changedRanges[key][0]} to ${changedRanges[key][1]}`);

  }

}


/**
 * Run the script and make any changes without asking the user first.
 */
function runWithoutConsent() {

  const namedRanges = getNamedRangesFromSheet();
  const newValues = getNewCostCodeValues();
  const changedRanges = compareNamedRanges(namedRanges, newValues);

  // If there are new values then get consent and update them
  if (Object.keys(changedRanges).length > 0) {
    updateNamedRanges(changedRanges, true);

    // Notify the user if no new values were found
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("No new values were found.");
  }
}


/**
 * Creates the menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts")
    .addItem("Update values", "runWithoutConsent")
    // .addItem("Recreate named ranges (make selection first)", "recreateNamedRanges")
    .addToUi();
}
