/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Fetch the fundraising total from the Enthuse page at url.
 */
function getEnthuseAmount(url) {

  let amount;

  // Make the GET request
  const response = UrlFetchApp.fetch(url);
  const status = response.getResponseCode();
  const html = response.getContentText();

  // If successful then extract the amount
  if (status == 200) {
    let regex = /£\d+/;
    let match = html.match(regex)[0];
    regex = /\d+/;
    amount = Number(match.match(regex)[0]);

  // Otherwise log the error and stop
  } else {
    Logger.log(`There was an error fetching ${URL}.`)
    SpreadsheetApp
        .getActiveSpreadsheet()
        .toast(`There was an error (status ${status}) fetching ${URL}.`,
            "ERROR", -1);
    Logger.log(status);
    Logger.log(html);
    return;
  }

  return amount;
}


/**
 * Update the named range with the name to the value (if it exists).
 */
function updateRange(name, value) {

  // Locate the named range
  const namedRanges = SpreadsheetApp
      .getActiveSpreadsheet().getNamedRanges();
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
        .toast(`Could not find the named range called ${name}`, "ERROR", -1);
  }
}


/**
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Update all amounts", "updateAll");
  menu.addToUi();
}


/**
 * Run all the update functions.
 */
function updateAll() {

  const enthuseTotal = getEnthuseAmount("https://exeterguild.enthuse.com/execontempchoir/profile");
  const adventCalendarTotal = getEnthuseAmount("https://exeterguild.enthuse.com/cf/contemporary-choir-s-advent-calendar")
  updateRange("SantaRun", enthuseTotal-adventCalendarTotal)
  updateRange("AdventCalendar", adventCalendarTotal)

}