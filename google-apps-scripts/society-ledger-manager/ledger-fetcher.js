/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Downloads the ledger from eXpense365 and returns it as a blob.
 */
function downloadLedger(expense365Config) {

  const SUBGROUP_ID = 0;
  const REPORT_ID = 30;

  // Prepare the URL, headers, and body
  const url = "https://service.expense365.com/ws/rest/eXpense365/RequestDocument";
  const auth = "Basic " + Utilities.base64Encode(expense365Config.email + ":" + expense365Config.password);
  const headers = {
    "User-Agent": "eXpense365|1.6.1|Google Pixel XL|Android|10|en_GB",
    "Authorization": auth,
    "Accept": "application/json",
    "If-Modified-Since": "Mon, 1 Oct 1990 05:00:00 GMT",
    "Content-Type": "text/plain;charset=UTF-8"
  };
  const data = ("{\"ReportID\": " + REPORT_ID +
    ",\"UserGroupID\": " + expense365Config.groupId +
    ",\"SubGroupID\": " + SUBGROUP_ID + "}");

  // Condense the above into a single object
  const params = {
    "contentType": "text/plain;charset=UTF-8",
    "headers": headers,
    "method": "post",
    "payload": data,
    "validateHttpsCertificates": false,
    "muteHttpExceptions": true
  };

  // Make the POST request
  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const responseText = response.getContentText();

  // If successful then return the content (the PDF)
  if (status === 200) {

    const fileName = `Ledger at ${response.getHeaders().Date}.pdf`;

    return response.getBlob().setName(fileName);

    // 401 means that they failed to authenticate
  } else if (status === 401) {

    Logger.log("401: User entered incorrect email and/or password");
    throw new Error("Oops! Looks like you entered the wrong email address and/or password.");

    // Otherwise log the error and stop
  } else {
    Logger.log(`There was a ${status} error fetching ${url}.`);
    Logger.log(responseText);
    throw new Error(`There was a ${status} error fetching ${url}. The server returned: ${responseText}`);
  }

}
