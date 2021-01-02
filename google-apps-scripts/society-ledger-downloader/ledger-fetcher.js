/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Downloads the ledger from eXpense365.
 * This is via a simple POST request with the email and password in the
 * headers (no OAuth unfortunately). The PDF ledger is returned as a
 * Blob.
 *
 * @param {eventObject.formInput} formData The details from the form.
 * @returns {(boolean|Blob)[]|(boolean|String)[]} [success, result]
 * where success is true if successful, and result is either the Blob
 * or a String with an error message.
 */
function downloadLedger(formData) {

  const SUBGROUP_ID = 0;
  const REPORT_ID = 30;

  // Prepare the URL, headers, and body
  const url = "https://service.expense365.com/ws/rest/eXpense365/RequestDocument";
  const auth = "Basic " + Utilities.base64Encode(formData.email + ":" + formData.password);
  const headers = {
    "User-Agent": "eXpense365|1.6.1|Google Pixel XL|Android|10|en_GB",
    "Authorization": auth,
    "Accept": "application/json",
    "If-Modified-Since": "Mon, 1 Oct 1990 05:00:00 GMT",
    "Content-Type": "text/plain;charset=UTF-8"
  };
  const data = ("{\"ReportID\": " + REPORT_ID +
    ",\"UserGroupID\": " + formData.groupId +
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
    return [true, response.getBlob().setName(fileName)];

    // 401 means that they failed to authenticate
  } else if (status === 401) {
    Logger.log("401: User entered incorrect email and/or password");
    return [false, "Oops! Looks like you entered the wrong email address and/or password."];

    // Some other error occurred
  } else {
    Logger.log(`There was a ${status} error fetching ${url}.`);
    Logger.log(responseText);
    return [false, `There was a ${status} error fetching the ledger. The server returned: ${responseText}`];
  }

}


/**
 * Saves the PDF blob to either a folder as a new file, or to an
 * existing PDF file. This existing file can optionally be renamed.
 *
 * @param {Blob} pdfBlob The PDF Blob of the ledger.
 * @param {String} id The ID of either the target folder or PDF.
 * @param {Boolean} folder true if the ID is a folder.
 * @param {Boolean} rename true if the PDF should be renamed.
 * @returns {(boolean|String)[]|(boolean|Exception)[]} [success, result]
 * where success is true if successful, and result is either the String
 * URL to the file or the Exception.
 */
function saveToDrive(pdfBlob, id, folder = false, rename = false) {

  try {

    // Save it to the folder with that ID
    if (folder === true) {

      // Upload the file
      const fileId = Drive.Files.generateIds().ids[0];
      let file = Drive.newFile();
      file.id = fileId;
      file.mimeType = "application/pdf";
      file.title = pdfBlob.getName();
      file = Drive.Files.insert(file, pdfBlob);

      // Set the parent folder
      Drive.Files.patch(file, fileId, {"addParents": id});

      return [true, file.alternateLink];

      // Save it to the existing PDF file
    } else {

      // Update the PDF
      const file = Drive.Files.get(id);
      file.originalFilename = pdfBlob.getName();
      const params = {"newRevision": true, "pinned": true};
      let newFile = Drive.Files.update(file, id, pdfBlob, params);

      // Rename it if asked
      if (rename === true) {
        newFile.title = pdfBlob.getName();
        Drive.Files.patch(newFile, newFile.id);
      }

      return [true, newFile.alternateLink];

    }

    // Just return the error
  } catch (err) {
    Logger.log(`There was an error saving the PDF to Drive: ${err}`)
    return [false, err];
  }

}


/**
 * Incomplete. Converts the PDF blob to XLSX using an online convertor.
 * This is not used because pdftoexcel.com recognises that the script is
 * a bot from it's user agent, and therefore rejects it.
 *
 * @param {Blob} pdfBlob The PDF Blob of the ledger.
 * @returns {String|*}
 */
function convertToXLSX(pdfBlob) {

  // Prepare the URL, headers, and body
  const url = "https://www.pdftoexcel.com/upload.instant.php";
  const headers = {
    "Connection": "keep-alive",
    "X-Requested-With": "XMLHttpRequest",
    "DNT": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4240.111 Safari/537.36",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Dest": "empty",
    "Accept-Language": "en-GB,en;q=0.9",
    "Accept": "*/*",
    "Origin": "https://www.pdftoexcel.com",
    "Referrer": "https://www.pdftoexcel.com"
  };
  const data = {"Filedata": pdfBlob};

  // Condense the above into a single object
  const params = {
    "headers": headers,
    "method": "post",
    "payload": data
  };

  // Make the POST request
  const response = UrlFetchApp.fetch(url, params);
  Logger.log(response);
  Logger.log(response.getResponseCode());
  Logger.log(response.getAllHeaders())
  return;
  const status = response.getResponseCode();
  const responseText = response.getContentText();

  // If successful
  if (status === 200) {

    Logger.log(responseText);
    return responseText;

    // Otherwise log the error and stop.
  } else {
    Logger.log(`There was a ${status} error fetching ${url}.`);
    Logger.log(responseText);
    return `Error: There was a ${status} error fetching ${url}. The server returned: ${responseText}`;
  }

}
