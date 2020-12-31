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
 * Processes the form submission from the sidebar.
 */
function processSidebarForm(e) {
  
  Logger.log(e)
  
  saveUserProperties(e.formInput)
}


/**
 * Callback function for a button action. Instructs Drive to display a
 * permissions dialog to the user, requesting `drive.file` scope for a
 * specific item on behalf of this add-on.
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
