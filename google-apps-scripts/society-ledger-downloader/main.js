function buildCommonHomePage() {

}


function buildDriveHomePage() {

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
