/**
 * Create the sidebar.
 */
function createSidebar() {
  const htmlTemplate = HtmlService.createTemplateFromFile("sidebar");
  htmlTemplate.expense365Config = getUserProperties();
  const html = htmlTemplate.evaluate().setTitle("Society Ledger Manager").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Returns the content in the HTML file.
 * No .html file extension is needed.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
  createSidebar();
}


/**
 * Runs when the spreadsheet is opened.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem("Start", "createSidebar")
    .addItem("Clear all saved user data", "clearSavedData")
    .addToUi();
}

/**
 * Deletes all the saved user data.
 */
function clearSavedData() {
  PropertiesService.getUserProperties().deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet().toast("All saved user data deleted.");
}


/**
 * Gets and returns the user properties
 */
function getUserProperties() {
  return PropertiesService.getUserProperties().getProperties();
}


/**
 * Saves the user properties specified as key-value pairs.
 */
function saveUserProperties(data) {

  // Get the current saved properties
  const userProperties = PropertiesService.getUserProperties()

  // Save them
  for (let prop in data) {
    if (Object.prototype.hasOwnProperty.call(data, prop)) {
      userProperties.setProperty(prop, data[prop]);
    }
  }
}


/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {

  const html = HtmlService.createHtmlOutputFromFile('picker.html')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a file');

}


/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}


/**
 * Gets the file ID from the URL.
 * Taken from https://stackoverflow.com/a/40324645.
 */
function getIdFrom(url) {
  let id;
  let parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);

  if (url.indexOf('?id=') >= 0) {
    id = (parts[6].split("=")[1]).replace("&usp", "");
    return id;

  } else {
    id = parts[5].split("/");
    //Using sort to get the id as it is the longest element.
    let sortArr = id.sort(function (a, b) {
      return b.length - a.length
    });
    id = sortArr[0];

    return id;
  }
}


/**
 * Processes the form submission from the sidebar.
 */
function processSidebarForm(formData) {

  Logger.log(`User clicked the button labelled ${formData.action}`);
  let result = "";

  // Clear the saved data
  if (formData.action === "Clear saved data") {
    PropertiesService.getUserProperties().deleteAllProperties();
    result = "All saved user data deleted.";

    // Runs the main script
  } else if (formData.action === "GO!") {

    // Save the preferences if requested
    if (formData.saveDetails === "on") {
      Logger.log("Saving form data.");
      saveUserProperties(formData);
    }

    // Download the PDF
    Logger.log("Downloading the PDF ledger...");
    const pdfBlob = downloadLedger(formData);
    Logger.log("PDF ledger downloaded successfully!");

    // Save the PDF if asked
    if (formData.savePDF === "on") {
      const id = getIdFrom(formData.PDFLocation);
      result = result.concat(saveToDrive(pdfBlob, id, "application/pdf"));
    }


  }
  return result;
}

