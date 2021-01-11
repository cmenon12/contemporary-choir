/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

class EnthuseFundraisingSource {

  /**
   * Constructor, returns the completed fundraising source.
   *
   * @param {String} rangeName the name of the named range with the
   * amount.
   * @param {String[]} urls the url(s) of the pages in this source.
   */
  constructor(rangeName, urls) {

    this.urls = urls;

    // Find the named range
    this.rangeName = rangeName;
    this.range = getNamedRange(rangeName);
    if (this.range === undefined) {
      throw new Error(`The named range ${rangeName} doesn't exist.`);
    }

    // Prepare to iterate over the URLs
    let pageAmount;
    this.amount = 0;
    this.donors = 0;
    this.fees = 0;

    // Fetch and set the total amount, donors, and fees
    for (let i = 0; i < urls.length; i++) {
      pageAmount = EnthuseFundraisingSource.fetchAmount(urls[i]);
      const donorsAndFees = EnthuseFundraisingSource.fetchDonorsAndFees(urls[i], pageAmount);
      this.amount = this.amount + pageAmount;
      this.donors = this.donors + donorsAndFees.donors;
      this.fees = this.fees + donorsAndFees.fees;
    }

    this.log();

  }


  /**
   * Fetch the fundraising total from the Enthuse page at url.
   *
   * @param {String} url the URL of the Enthuse page.
   * @returns {Number} the amount extracted.
   */
  static fetchAmount(url) {

    // Make the GET request
    const response = UrlFetchApp.fetch(url);
    const status = response.getResponseCode();
    const html = response.getContentText();

    // If successful then extract the amount
    if (status === 200) {

      // Determine the index to use
      let index = 0;
      if (url.includes("/pf/")) {
        index = 1;
      }

      // Extract and return the total
      return EnthuseFundraisingSource.regexSearchForNumber(html,
        /£\d+/g, index);

      // Otherwise log the error and stop
    } else {
      Logger.log(`There was an error fetching ${url}.`);
      SpreadsheetApp
        .getActiveSpreadsheet()
        .toast(`There was an error (status ${status}) fetching ${url}.`,
          "ERROR", -1);
      Logger.log(status);
      Logger.log(html);
      throw new Error(`There was an error (status ${status}) fetching ${url}.`);
    }
  }

  /**
   * Fetch the number of donors and total fees
   *
   * @param {String} url the URL of the Enthuse page.
   * @param {Number} amount the amount raised.
   * @returns {{}} the total in fees and total donors.
   */
  static fetchDonorsAndFees(url, amount) {

    let donors;

    // Make the GET request
    const response = UrlFetchApp.fetch(url);
    const status = response.getResponseCode();
    const html = response.getContentText();

    // If successful then extract the amount and number of donors
    if (status === 200) {

      // Get the total number of donors
      donors = EnthuseFundraisingSource.regexSearchForNumber(html,
        /<span id="js-total-donations">\d+</, 0)

      // Calculate the fees and return
      const fees = (amount * 0.019) + (0.2 * donors);
      return {"fees": fees, "donors": donors};

      // Otherwise log the error and stop
    } else {
      Logger.log(`There was an error fetching ${url}.`);
      SpreadsheetApp
        .getActiveSpreadsheet()
        .toast(`There was an error (status ${status}) fetching ${url}.`,
          "ERROR", -1);
      Logger.log(status);
      Logger.log(html);
      throw new Error(`There was an error (status ${status}) fetching ${url}.`);
    }
  }

  static regexSearchForNumber(string, regex, index) {
    let match = string.match(regex)[index];
    return Number(match.match(/\d+/)[0]);
  }

  /**
   * Recalculate the amount, donors, and fees by removing the sources.
   * @param {EnthuseFundraisingSource[]} sources the sources to remove.
   * @returns {EnthuseFundraisingSource} this object, for chaining.
   */
  removeAndRecalculate(sources) {

    // Subtract the amount and donors from each of the sources
    for (let i = 0; i < sources.length; i++) {
      this.amount = this.amount - sources[i].getAmount();
      if (this.amount < 0) {
        throw new Error("The amount has dropped below 0.");
      }
      this.donors = this.donors - sources[i].getDonors();
      if (this.donors < 0) {
        throw new Error("The amount has dropped below 0.");
      }
    }

    // Recalculate the fees
    this.fees = (this.amount * 0.019) + (0.2 * this.donors);
    this.log();
    return this;
  }

  /**
   * Returns the amount.
   *
   * @returns {Number} the amount.
   */
  getAmount() {
    return this.amount;
  }

  /**
   * Returns the number of donors.
   *
   * @returns {Number}
   */
  getDonors() {
    return this.donors;
  }

  /**
   * Returns the fees.
   *
   * @returns {Number} the fees.
   */
  getFees() {
    return this.fees;
  }

  /**
   * Returns the named range.
   *
   * @returns {Range} the named range.
   */
  getRange() {
    return this.range;
  }

  /**
   * Returns the name of the named range.
   *
   * @returns {String} the name of the named range.
   */
  getRangeName() {
    return this.rangeName;
  }

  /**
   * Updates the value in the named range.
   *
   * @param {String} toastMsg the toast message to update
   * @returns {String} the newly-updated toast message.
   */
  updateRange(toastMsg) {

    // If the amount has changed then update it and update the toast
    let oldValue = this.range.getValue();
    if (this.amount !== oldValue) {
      toastMsg = toastMsg.concat(`${this.rangeName} has increased by £${value - oldValue}. `);
      this.range.setValue(this.amount);
      Logger.log(`Updated the named range ${this.rangeName} to ${this.amount} successfully.`);

      // If it hasn't changed then just log it
    } else {
      Logger.log(`Range ${this.rangeName} already had a value of ${this.amount}.`);
    }

    // If the fees have changed then update it
    oldValue = this.range.offset(0, 1).getValue();
    if (this.fees !== oldValue) {
      this.range.offset(0, 1).setValue(this.fees);
      Logger.log(`Updated the fees for named range ${this.rangeName} to ${-this.fees} successfully.`);

      // If it hasn't changed then just log it
    } else {
      Logger.log(`The fees for ${this.rangeName} already had a value of ${this.fees}.`);
    }

    // Return the possibly-updated toast message
    return toastMsg;
  }

  /**
   * Saves the whole thing to the log.
   */
  log() {
    Logger.log(`The EnthuseFundraisingSource object is: ${JSON.stringify(this)}`);
  }

}


/**
 * Get the range with the name from the active spreadsheet.
 * Returns the range if found, otherwise undefined.
 *
 * @param {String} name the name of the range.
 * @returns {Range|undefined} the named range.
 */
function getNamedRange(name) {

  const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
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
  const adventCalendar = new EnthuseFundraisingSource("AdventCalendar",
    ["https://exeterguild.enthuse.com/cf/contemporary-choir-s-advent-calendar"]);
  const virtualChoir = new EnthuseFundraisingSource("VirtualChoir",
    ["https://exeterguild.enthuse.com/cf/virtual-choir-performances"]);
  const santaRun = new EnthuseFundraisingSource("SantaRun",
    ["https://exeterguild.enthuse.com/pf/jamie-harvey",
      "https://exeterguild.enthuse.com/pf/rebekah-lydia",
      "https://exeterguild.enthuse.com/pf/ellie-doherr",
      "https://exeterguild.enthuse.com/pf/bethany-piper-edd11",
      "https://exeterguild.enthuse.com/pf/tom-joshi-cale",
      "https://exeterguild.enthuse.com/pf/chris-menon-santa-run-2020-9558f"]);

  const otherEnthuse = new EnthuseFundraisingSource("OtherEnthuse",
    ["https://exeterguild.enthuse.com/execontempchoir/profile"])
    .removeAndRecalculate([adventCalendar, virtualChoir, santaRun]);

  // Update the amounts
  let toastMsg = "";
  toastMsg = adventCalendar.updateRange(toastMsg);
  toastMsg = virtualChoir.updateRange(toastMsg);
  toastMsg = santaRun.updateRange(toastMsg);
  toastMsg = otherEnthuse.updateRange(toastMsg);

  // Produce the toast to notify the user
  if (toastMsg === "") {
    SpreadsheetApp.getActiveSpreadsheet().toast("No new changes were found.", "");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMsg, "", -1);
  }

  // Update when this script was last run
  const range = getNamedRange("ScriptLastRun");
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