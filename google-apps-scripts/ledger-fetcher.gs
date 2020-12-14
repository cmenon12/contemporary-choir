/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Gets and returns the eXpense365 config.
 * If it doesn't exist for this user, then they'll be asked to input it.
 * The user can force them to be updated by passing true for change.
 */
function getExpense365Config(change = false) {

  const GROUP_ID = 35778;
  const SUBGROUP_ID = 0;

  // Get the current saved email and password
  const userProperties = PropertiesService.getUserProperties();
  let email = userProperties.getProperty("expense365_email");
  let password = userProperties.getProperty("expense365_password");

  if (email === null || password === null || change === true) {

    Logger.log("Asking the user for new eXpense365 credentials...");
    const ui = SpreadsheetApp.getUi();

    // Get the username
    let response = ui.prompt("What is your eXpense365 email address?",
                             "This is user-specific, and will only be accessible to you.",
                             ui.Button.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      email = response.getResponseText();
      Logger.log("The user provided an email address.");
    } else {
      Logger.log("ERROR: The user didn't want to provide an email address.");
      throw new Error("The user didn't want to provide an email address.");
    }

    // Get the password
    response = ui.prompt("What is your eXpense365 password?",
                         "This is user-specific, and will only be accessible to you.",
                         ui.Button.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      password = response.getResponseText();
      Logger.log("The user provided a password.");
    } else {
      Logger.log("ERROR: The user didn't want to provide a password.");
      throw new Error("The user didn't want to provide a password.");
    }

    // Save the new properties
    userProperties.setProperty("expense365_email", email)
    .setProperty("expense365_password", password);

  }

  // Return the details
  return {"group_id": GROUP_ID,
          "subgroup_id": SUBGROUP_ID,
          "email": email,
          "password": password};
}


/**
 * Downloads the ledger from eXpense365 and returns it as a blob.
 */
function downloadLedger() {

  const REPORT_ID = 30;
  const expense365Config = getExpense365Config();

  // Prepare the URL, headers, and body
  const url = "https://service.expense365.com/ws/rest/eXpense365/RequestDocument";
  const auth = "Basic " + Utilities.base64Encode(expense365Config.email + ":" + expense365Config.password);
  const headers = {"Host": "service.expense365.com:443",
                   "User-Agent": "eXpense365|1.6.1|Google Pixel XL|Android|10|en_GB",
                   "Authorization": auth,
                   "Accept": "application/json",
                   "If-Modified-Since": "Mon, 1 Oct 1990 05:00:00 GMT",
                   "Content-Type": "text/plain;charset=UTF-8"};
  const data = ("{\"ReportID\":" + REPORT_ID +
                ",\"UserGroupID\":" + expense365Config.group_id +
                ",\"SubGroupID\":" + expense365Config.subgroup_id + "}");

  // Condense the above into a single object
  const params = {"contentType": "text/plain;charset=UTF-8",
                  "headers": headers,
                  "method": "post",
                  "payload": data,
                  "validateHttpsCertificates": false};

  // Make the POST request
  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const html = response.getContentText();

  // If successful then return the content (the PDF)
  if (status == 200) {

    return response.getBlob();

  // Otherwise log the error and stop
  } else {
    Logger.log(`There was a ${status} error fetching ${URL}.`);
    Logger.log(html);
    throw new Error(`There was a ${status} error fetching ${URL}.`);
  }

}
