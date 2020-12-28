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


/**
 * Saves the blob to the file or folder ID.
 * Will ensure that if it's a file, that it matches the MIME type specified.
 * Returns the outcome as HTML-formatted user friendly text.
 */
function saveToDrive(blob, id, mimeType) {

  let result;

  // Save it to the folder with that ID
  try {
    const folder = DriveApp.getFolderById(id);
    const newFile = folder.createFile(blob);
    Logger.log(`Ledger saved to a new file named ${newFile.getName()} in ${folder.getName()}.`);
    result = `Ledger saved to a new file named <a href="${newFile.getUrl()}" target="_blank">${newFile.getName()}<\/a> in <a href="${folder.getUrl()}" target="_blank">${folder.getName()}<\/a>.<br>`;

    // Try updating the existing file with that ID instead
  } catch (err1) {
    try {

      // Check that the file has the correct MIME type
      const file = Drive.Files.get(id);
      if (file.mimeType !== mimeType) {
        Logger.log(`The file MIME type is ${file.mimeType} which doesn't match ${mimeType}.`)
        result = `Error: The file MIME type is ${file.mimeType} which doesn't match ${mimeType}.`;

      } else {
        // Upload the new revision
        const params = {"newRevision": true, "pinned": true};
        const newFile = Drive.Files.update(file, id, blob, params);
        Logger.log(`Ledger saved to an existing file named ${newFile.title}.`);
        result = `Ledger saved to an existing file named <a href="${newFile.alternateLink}" target="_blank">${newFile.title}<\/a>.<br>`;
      }

      // The ID doesn't match any folder nor file
    } catch (err2) {
      Logger.log(`No folder nor file exists for the ID ${id}.`);
      result = `Error: No folder nor file exists for the ID ${id}.`;

    }
  }
  return result;
}
