/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Fetches our fundraising total from Enthuse,
 * and updates the total in the sheet.
 */
function enthuse() {

  // Constants
  const URL = "https://exeterguild.enthuse.com/execontempchoir/profile";
  const RANGE_NAME = "SantaRun"
  let amount;

  // Make the GET request
  const response = UrlFetchApp.fetch(URL);
  const status = response.getResponseCode();
  const html = response.getContentText();

  // If successful then extract the amount
  if (status == 200) {
    let regex = /£\d+/;
    let match = html.match(regex)[0];
    regex = /\d+/
    amount = Number(match.match(regex)[0]);
    Logger.log(amount);

  // Otherwise log the error and stop
  } else {
    Logger.log(status);
    Logger.log(html);
    return;
  }

  updateRange(RANGE_NAME, amount)

}


/**
 * Update the named range with the name to the value (if it exists).
 */
function updateRange(name, value) {

  // Locate the named range
  const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  let range;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() == name) {
      range = namedRanges[i].getRange();
    }
  }

  // If the range has been located then update it
  if (range != undefined) {

    range.setValue(value);
    Logger.log(`Updated the named range ${name} to ${value} successfully.`);

  } else {
    Logger.log(`Could not find the named range called ${name}`);
    SpreadsheetApp.getActiveSpreadsheet()
        .toast(`Could not find the named range called ${name}`, "ERROR")
  }
}


/**
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Update Santa Run total (Enthuse)", "enthuse");
  menu.addItem("Update all", "updateAll");
  menu.addToUi();
}


/**
 * Run all the update functions.
 */
function updateAll() {

  enthuse()

}