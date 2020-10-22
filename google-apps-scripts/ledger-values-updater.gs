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

  const NAMED_RANGES_SHEET_NAME = "Values";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMED_RANGES_SHEET_NAME);
  const namedRanges = sheet.getNamedRanges();
  let namedRangesArray = []

  // For each named range in this sheet
  for (let i = 0; i < namedRanges.length; i++) {

    // If the named range name doesn't include Pre,
    // and the A1 notation doesn't include a : (so it's a single cell)
    if (!namedRanges[i].getName().includes("Pre") && !namedRanges[i].getRange().getA1Notation().includes(":")) {

      // Add it to our array
      namedRangesArray.push([namedRanges[i].getName(), namedRanges[i].getRange()]);

    }
  }

  Logger.log("We found " + namedRangesArray.length + " named ranges.");
  Logger.log(namedRangesArray);

  return namedRangesArray;

}

function findNamedRangesFromLedger() {

  const SHEET_NAME = "Original";
}

/**
 * Creates the menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts")
      .addItem("getNamedRangesFromSheet", "getNamedRangesFromSheet")
      .addToUi();
}


