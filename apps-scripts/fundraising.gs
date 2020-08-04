/*
 * Fetches and parses Maddy's fundraising total,
 * applies fees and postage costs, and updates the
 * named range called MaddyBakes in the sheet.
 *
 * See https://www.gofundme.com/pricing/ for fees info.
 */
function maddyBakes() {

  // Constants
  const POSTAGE_COST = 2;
  const URL = "https://www.gofundme.com/f/letterbox-bakes-for-macmillan-cancer-support";

  // Make the GET request
  var response = UrlFetchApp.fetch(URL);
  var status = response.getResponseCode();
  var html = response.getContentText();

  // If successful then extract the amount and number of donors
  if (status == 200) {
    var regex = /(\"current_amount\":\d{1,})/
    var match = html.match(regex)[0];
    regex = /\d{1,}/
    var amount = Number(match.match(regex)[0])
    Logger.log(amount);

    regex = /(\"donation_count\":\d{1,})/
    match = html.match(regex)[0];
    regex = /\d{1,}/
    var donors = Number(match.match(regex)[0]);
    Logger.log(donors);

  // Otherwise log the error and stop
  } else {
    Logger.log(status);
    Logger.log(html);
    return;
  }

  // Locate the named range
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var range;
  for (var i=0; i<namedRanges.length; i++) {
    if (namedRanges[i].getName() == "MaddyBakes") {
      range = namedRanges[i].getRange();
    }
  }

  // If the range has been located then process the actual amount
  // and fill the range
  if (range != undefined) {

    var amountMinusFeesAndPostage = (amount*0.971)-(0.25*donors)-(POSTAGE_COST*donors);
    range.setValue(amountMinusFeesAndPostage);
    Logger.log("Updated the named range successfully.");

  } else {
    Logger.log("Could not find the named range called MaddyBakes");
  }
}


/*
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Update Maddy's Letterbox Bakes", "maddyBakes");
  menu.addToUi();
}
