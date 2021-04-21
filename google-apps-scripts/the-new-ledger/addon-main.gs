/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Deletes all the saved user data and clears the form.
 *
 * @param {eventObject} e The event object.
 * @returns {ActionResponse} The ActionResponse that reloads the
 * current Card.
 */
function clearSavedData(e) {

  PropertiesService.getUserProperties().deleteAllProperties();

  // Build and return an ActionResponse
  const navigation = CardService.newNavigation()
    .updateCard(buildSheetsHomePage(e));
  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .build();
}


/**
 * Gets and returns the user properties
 */
function getUserProperties() {
  return PropertiesService.getUserProperties().getProperties();
}


/**
 * Saves the user data specified as key-value pairs. These are saved to
 * the user's properties, so they're unique to them.
 *
 * @param {eventObject.formInput} data The data to be saved as key-value
 * pairs.
 */
function saveUserProperties(data) {

  // Get the current saved properties
  const userProperties = PropertiesService.getUserProperties();

  // Save them
  for (let prop in data) {
    if (Object.prototype.hasOwnProperty.call(data, prop)) {
      userProperties.setProperty(prop, data[prop]);
    }
  }

}


/**
 * Processes the form submission from the Sheets sidebar.
 *
 * @param {eventObject} e The event object.
 * @returns {ActionResponse} The response, either a link that opens in
 * a new tab or an error notification.
 */
function processSheetsSidebarForm(e) {

  saveUserProperties(e.formInput);

  // Attempt to get the sheet
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(getUserProperties().sheetURL);
    const sheet = spreadsheet.getSheetByName(getUserProperties().sheetName);
  } catch (err) {
    Logger.log(JSON.stringify(err))
    return updateWithHomepage(e);
  }

  return buildNotification("success");


}
