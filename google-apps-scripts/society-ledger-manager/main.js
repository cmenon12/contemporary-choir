/**
 * Create the sidebar.
 */
function createHome() {
  
  const htmlTemplate = HtmlService.createTemplateFromFile("home");
  htmlTemplate.expense365Config = getUserProperties();
  const html = htmlTemplate.evaluate().setTitle("Society Ledger Manager").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);

}


/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
  createHome();
}


/**
 * Runs when the spreadsheet is opened.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem("Start", "createHome")
      .addItem("Delete all saved user data", "deleteSavedData")
      .addToUi();
}


function deleteSavedData() { PropertiesService.getUserProperties().deleteAllProperties(); };


/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  
  var html = HtmlService.createHtmlOutputFromFile('picker.html')
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


function fetchLedger(formData) {
  
  Logger.log(formData.action);
  
  if (formData.action == "Clear saved data") {
    PropertiesService.getUserProperties().deleteAllProperties();
    return "All saved user data deleted!"; 
  }
  
  if (formData.saveDetails == "on") {
    Logger.log("Saving form data.");
    saveUserProperties(formData);
  }

  const pdfBlob = downloadLedger(formData);
}