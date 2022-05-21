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

  // Create the navigation
  const navigation = CardService.newNavigation()
    .updateCard(buildDriveHomePage(e));

  // Build and return an ActionResponse
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
 *
 * Gets the Drive file ID from the URL.
 * Taken from https://stackoverflow.com/a/40324645.
 *
 * @param {String} url The URL to get the file ID from.
 * @returns {String} The file ID.
 */
function getIdFrom(url) {
  let id;
  let parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);

  if (url.indexOf('?id=') >= 0) {
    id = (parts[6].split("=")[1]).replace("&usp", "");
    return id;

  } else {
    id = parts[5].split("/");
    // Using sort to get the id as it is the longest element.
    let sortArr = id.sort(function (a, b) {
      return b.length - a.length
    });
    id = sortArr[0];

    return id;
  }
}


/**
 * Processes the form submission from the Drive sidebar.
 * This will download the ledger, save it as requested, and then return
 * the homepage with a result message.
 *
 * @param {eventObject} e The event object.
 * @returns {Card} The homepage card
 */
function processDriveSidebarForm(e) {

  saveUserProperties(e.formInput);

  const currentFile = e.drive.activeCursorItem;

  // Download the ledger
  let result = downloadLedger(e.formInput);
  let pdfBlob;
  if (result[0] === true) {
    pdfBlob = result[1];
  } else {
    // If there was an error then stop here
    PropertiesService.getUserProperties().setProperty("resultMessage", result[1]);
    return buildDriveHomePage(e);
  }

  // Upload it to Drive
  let folder = false;
  let rename = false;
  if (currentFile.mimeType === "application/vnd.google-apps.folder") {
    folder = true;
  } else if (e.formInput.saveMethod === "rename") {
    rename = true;
  }
  result = saveToDrive(pdfBlob, currentFile.id, folder, rename);

  // Return the result
  PropertiesService.getUserProperties().setProperty("resultMessage", result[1]);
  return buildDriveHomePage(e);

}


/**
 * Callback function for a button action. Instructs Drive to display a
 * permissions dialog to the user, requesting `drive.file` scope for a
 * specific item on behalf of this add-on.
 * Also saves the form data.
 *
 * @param {Object} e The parameters object that contains the item's
 *   Drive ID.
 * @return {DriveItemsSelectedActionResponse}
 */
function onRequestFileScopeButtonClicked(e) {

  saveUserProperties(e.formInput)

  const idToRequest = e.parameters.id;
  return CardService.newDriveItemsSelectedActionResponseBuilder()
    .requestFileScope(idToRequest)
    .build();
}
