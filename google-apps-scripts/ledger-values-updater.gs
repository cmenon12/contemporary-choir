/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Locates and returns the named ranges that we'll want to update.
 * These are the named ranges in the Values sheet that will also appear
 * in the PDF ledger.
 */
function getNamedRangesFromSheet() {

  // The name of the sheet with these named ranges.
  const NAMED_RANGES_SHEET_NAME = "Values";

  // Define the sheet and get the named ranges.
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(NAMED_RANGES_SHEET_NAME);
  const namedRanges = sheet.getNamedRanges();

  // Prepare to save the values
  const namedRangesDict = {};

  // For each named range in this sheet
  for (let i = 0; i < namedRanges.length; i++) {

    // If the named range name doesn't include Pre,
    // and the A1 notation doesn't include a : (so it's a single cell)
    if (!namedRanges[i].getName().includes("Pre") && !namedRanges[i]
        .getRange().getA1Notation().includes(":")) {

      // Add it to our dictionary, using the name as the key.
      namedRangesDict[namedRanges[i].getName()] = namedRanges[i].getRange();

    }
  }

  Logger.log("We found " + Object.keys(namedRangesDict).length +
      " named ranges.");
  Logger.log(JSON.stringify(namedRangesDict));

  return namedRangesDict;

}


/**
 * Extracts the income, expenditure, and balance for each cost code.
 * This also includes the total income, expenditure, and balance, as well as the balance brought forward.
 * It also groups all of the individual TOMS cost codes together.
 */
function getNewCostCodeValues() {

  // The name of the sheet with the imported ledger
  const LEDGER_SHEET_NAME = "Original";

  // Define the sheet and get the named ranges.
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(LEDGER_SHEET_NAME);

  // Prepare to save the values
  const costCodeValues = {};
  costCodeValues["TOMSIncome"] = 0
  costCodeValues["TOMSExpenditure"] = 0
  costCodeValues["TOMSBalance"] = 0

  let costCodeName;

  // Search for the total for each cost code (and the grand total)
  const finder = sheet.createTextFinder("Totals for ").matchEntireCell(false);
  const foundRanges = finder.findAll();
  for (let i = 0; i < foundRanges.length; i++) {

    // Get the name of the cost code
    costCodeName = String(foundRanges[i].getValue())
        .replace("Totals for ", "");

    // If it's a TOMS cost code then add it to the existing TOMS values
    // This keeps all of the TOMS grouped together
    if (costCodeName.includes("TOMS")) {
      costCodeValues["TOMSIncome"] += foundRanges[i].offset(0, 1).getValue()
      costCodeValues["TOMSExpenditure"] += foundRanges[i].offset(0, 2).getValue()
      costCodeValues["TOMSBalance"] += foundRanges[i].offset(1, 2).getValue()

      // If it's the last element then this must be the grand total
      // This uses the cost code name Total and includes the balance brought forward
    } else if (i == foundRanges.length - 1) {
      costCodeValues["TotalIncome"] = foundRanges[i].offset(0, 1).getValue()
      costCodeValues["TotalExpenditure"] = foundRanges[i].offset(0, 2).getValue()
      costCodeValues["BalanceBroughtForward"] = foundRanges[i].offset(2, 2).getValue()
      costCodeValues["TotalBalance"] = foundRanges[i].offset(3, 2).getValue()

      // Otherwise just add the standalone cost code itself
    } else {
      costCodeValues[costCodeName.replace(/\s/g, '') + "Income"] =
          foundRanges[i].offset(0, 1).getValue()
      costCodeValues[costCodeName.replace(/\s/g, '') + "Expenditure"] =
          foundRanges[i].offset(0, 2).getValue()
      costCodeValues[costCodeName.replace(/\s/g, '') + "Balance"] =
          foundRanges[i].offset(1, 2).getValue()
    }
  }

  Logger.log("We found " + Object.keys(costCodeValues).length +
      " cost code values.");
  Logger.log(JSON.stringify(costCodeValues));

  return costCodeValues;

}


/**
 * Compare the named ranges with the new values,
 * and return those that have changed.
 */
function compareNamedRanges(namedRanges, newValues) {

  const changedRanges = {};

  for (let key in namedRanges) {

    // If the value has changed then save it
    if (namedRanges[key].getValue() != newValues[key]) {

      // This saves [old value, new value, Range]
      changedRanges[key] = [namedRanges[key].getValue(), newValues[key], namedRanges[key]]

      Logger.log(key + " has changed");
      Logger.log("old:" + namedRanges[key].getValue() + ", new:" + newValues[key]);
    }
  }

  return changedRanges;
}


function getUserConsent(changedValues) {
}


/**
 * Update the named ranges with the new values.
 */
function updateNamedRanges(changedRanges) {

  // Get the current date and time
  const now = new Date();

  // For each key in changedRanges
  for (let key in changedRanges) {

    // Set the range to the new value
    changedRanges[key][2].setValue(changedRanges[key][1])

    // Set the source to reflect that this is an automatic update
    changedRanges[key][2].offset(0, 3).setValue("PDF (auto)")
    //changedRanges[key][2].offset(0, 3).setNote(now)

    Logger.log("Changed " + key + " in " +
        changedRanges[key][2].getA1Notation() + " from " +
        changedRanges[key][0] + " to " + changedRanges[key][1]);

  }

}

/**
 * Run the script, and ask the user to confirm the changes to the values.
 */
function runWithConsent() {

  const namedRanges = getNamedRangesFromSheet();
  const newValues = getNewCostCodeValues();
  const changedRanges = compareNamedRanges(namedRanges, newValues);

  if (Object.keys(changedRanges).length > 0 && getUserConsent(changedRanges) === true) {
    updateNamedRanges(changedRanges)
  }
}


function runWithoutConsent() {

  const namedRanges = getNamedRangesFromSheet();
  const newValues = getNewCostCodeValues();
  const changedRanges = compareNamedRanges(namedRanges, newValues);

  if (Object.keys(changedRanges).length > 0) {
    updateNamedRanges(changedRanges)
  }
}


/**
 * Creates the menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts")
      .addItem("getNamedRangesFromSheet", "getNamedRangesFromSheet")
      .addItem("getNewCostCodeValues", "getNewCostCodeValues")
      .addToUi();
}