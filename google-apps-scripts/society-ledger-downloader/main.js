/**
 * Deletes all the saved user data.
 */
function clearSavedData() {
  PropertiesService.getUserProperties().deleteAllProperties();
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
 * Processes the form submission from the Drive sidebar.
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
    return buildErrorNotification(result[1]);
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

  // Open the newly-uploaded ledger
  if (result[0] === true) {
    const openLink = CardService.newOpenLink()
      .setUrl(result[1])
      .setOpenAs(CardService.OpenAs.FULL_SIZE)
      .setOnClose(CardService.OnClose.NOTHING)
    return CardService.newActionResponseBuilder().setOpenLink(openLink).build();

  } else {
    // Otherwise just notify of the error
    return buildErrorNotification(result[1]);
  }

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
