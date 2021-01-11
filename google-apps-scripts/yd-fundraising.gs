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
 *
 * @param {String} url the URL of the fundraising page.
 * @returns {Number} the amount extracted.
 */
function getEnthuseAmount(url) {

  // Make the GET request
  const response = UrlFetchApp.fetch(url);
  const status = response.getResponseCode();
  const html = response.getContentText();

  // If successful then extract the amount
  if (status === 200) {

    // Determine the index to use
    let index = 0;
    if (url.includes("pf/")) {
      index = 1;
    }

    // Extract and return the total
    let match = html.match(/£\d+/g)[index];
    return Number(match.match(/\d+/)[0]);

    // Otherwise log the error and stop
  } else {
    Logger.log(`There was an error fetching ${URL}.`);
    SpreadsheetApp
      .getActiveSpreadsheet()
      .toast(`There was an error (status ${status}) fetching ${URL}.`,
        "ERROR", -1);
    Logger.log(status);
    Logger.log(html);
    throw new Error(`There was an error (status ${status}) fetching ${URL}.`);
  }
}


/**
 * Fetch the total fees and number of donors.
 *
 * @param {String} profileUrl the URL of the fundraising profile.
 * @returns {{}} the total in fees and total donors.
 */
function getEnthuseFeesAndDonors(profileUrl) {

  let donors;
  let amount;

  // Make the GET request
  const response = UrlFetchApp.fetch(profileUrl);
  const status = response.getResponseCode();
  const html = response.getContentText();

  // If successful then extract the amount and number of donors
  if (status === 200) {

    // Get the total amount
    let match = html.match(/£\d+/)[0];
    amount = Number(match.match(/\d+/)[0]);

    // Get the total number of donors
    match = html.match(/<span id="js-total-donations">\d+</)[0];
    donors = Number(match.match(/\d+/)[0]);

    // Calculate the fees and return
    const fees = (amount * 0.005) + (0.2 * donors);
    return {"fees": fees, "donors": donors};

    // Otherwise log the error and stop
  } else {
    Logger.log(`There was an error fetching ${URL}.`);
    SpreadsheetApp
      .getActiveSpreadsheet()
      .toast(`There was an error (status ${status}) fetching ${URL}.`,
        "ERROR", -1);
    Logger.log(status);
    Logger.log(html);
    throw new Error(`There was an error (status ${status}) fetching ${URL}.`);
  }
}


/**
 * Get the range with the name from the spreadsheet.
 * Returns the range if found, otherwise undefined.
 *
 * @param {String} name the name of the range.
 * @param {Spreadsheet} spreadsheet the spreadsheet to search.
 * @returns {Range|undefined} the named range.
 */
function getNamedRange(name, spreadsheet) {

  const namedRanges = spreadsheet.getNamedRanges();
  let range;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === name) {
      range = namedRanges[i].getRange();
      break;
    }
  }
  return range;
}


/**
 * Fetches the sum fundraising total for all of the URLs in the array.
 *
 * @param {String[]} urls the enthuse URLs.
 * @returns {Number} the sum of the total for all the URLs.
 */
function getMultipleEnthuseAmounts(urls) {

  let total = 0;
  for (let i = 0; i < urls.length; i++) {
    total = total + getEnthuseAmount(urls[i]);
  }
  return total;
}


/**
 * Update the named range with the name to the value (if it exists).
 *
 * @returns {String} the newly-updated toast message.
 */
function updateRange(name, value, toastMsg) {

  const range = getNamedRange(name, SpreadsheetApp.getActiveSpreadsheet());

  // If the range has been located
  if (range !== undefined) {

    // If the value has changed then update it and update the toast
    const oldValue = range.getValue();
    if (value !== oldValue) {
      toastMsg = toastMsg.concat(`${name} has increased by £${value - oldValue}. `);
      range.setValue(value);
      Logger.log(`Updated the named range ${name} to ${value} successfully.`);

      // If it hasn't changed then just log it
    } else {
      Logger.log(`Range ${name} already had a value of ${value}.`);
    }

    // If the range could not be found
  } else {
    Logger.log(`Could not find the named range called ${name}.`);
    toastMsg = toastMsg.concat(`Could not find the named range called ${name}. `);
  }

  // Return the possibly-updated toast message
  return toastMsg;
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

  // Retrieve the totals
  const enthuseTotal = getEnthuseAmount("https://exeterguild.enthuse.com/execontempchoir/profile");
  const adventCalendarTotal = getEnthuseAmount("https://exeterguild.enthuse.com/cf/contemporary-choir-s-advent-calendar");
  const virtualChoirTotal = getEnthuseAmount("https://exeterguild.enthuse.com/cf/virtual-choir-performances");

  // Get the Santa Run total
  const santaRunUrls = ["https://exeterguild.enthuse.com/pf/jamie-harvey",
    "https://exeterguild.enthuse.com/pf/rebekah-lydia",
    "https://exeterguild.enthuse.com/pf/ellie-doherr",
    "https://exeterguild.enthuse.com/pf/bethany-piper-edd11",
    "https://exeterguild.enthuse.com/pf/tom-joshi-cale",
    "https://exeterguild.enthuse.com/pf/chris-menon-santa-run-2020-9558f"];
  const santaRunTotal = getMultipleEnthuseAmounts(santaRunUrls);

  // Run the update functions
  let toastMsg = "";
  toastMsg = updateRange("SantaRun", santaRunTotal, toastMsg);
  toastMsg = updateRange("AdventCalendar", adventCalendarTotal, toastMsg);
  toastMsg = updateRange("VirtualChoir", virtualChoirTotal, toastMsg);
  toastMsg = updateRange("OtherEnthuse",
    enthuseTotal - santaRunTotal - adventCalendarTotal - virtualChoirTotal, toastMsg);

  // Update the fees
  const summary = getEnthuseFeesAndDonors("https://exeterguild.enthuse.com/execontempchoir/profile");
  toastMsg = updateRange("EnthuseFees", -summary.fees, toastMsg);
  getNamedRange("EnthuseFeesNotes", SpreadsheetApp.getActiveSpreadsheet())
    .setValue(`Estimate, based on ${summary.donors} donors.`);
  Logger.log(`Updated the number of donors in named range EnthuseFeesNotes to ${summary.donors} successfully.`);

  // Produce the toast to notify the user
  if (toastMsg === "") {
    SpreadsheetApp.getActiveSpreadsheet().toast("No new changes were found.", "");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMsg, "", -1);
  }

  // Update when this script was last run
  const range = getNamedRange("ScriptLastRun", SpreadsheetApp.getActiveSpreadsheet());
  if (range !== undefined) {
    const date = new Date();
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    let minutes = date.getMinutes().toString();
    if (parseInt(minutes) < 10) {
      minutes = "0" + minutes;
    }
    const dateString = `Script last run on ${days[date.getDay()]} ${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()} at ${date.getHours()}:${minutes}.`;
    range.setValue(dateString);
  }

  // Force these changes
  SpreadsheetApp.flush();

}