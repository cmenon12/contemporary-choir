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

  // Ensure that empty items are saved as empty
  actions = {}
  if (!data.hasOwnProperty("copyNewSheetName") && data.hasOwnProperty("copy")) {
    actions.copyNewSheetName = ""
  }

  // Save them
  for (let prop in actions) {
    if (Object.prototype.hasOwnProperty.call(actions, prop)) {
      userProperties.setProperty(prop, actions[prop]);
    }
  }

}


/**
 * Returns an array of CSS colour names.
 *
 * @returns {String[]} an array of colour names.
 */
function getColourNames() {
  return ["aliceblue", "antiquewhite", "aqua", "aquamarine", "azure", "beige", "bisque", "black", "blanchedalmond", "blue", "blueviolet", "brown", "burlywood", "cadetblue", "chartreuse", "chocolate", "coral", "cornflowerblue", "cornsilk", "crimson", "cyan", "darkblue", "darkcyan", "darkgoldenrod", "darkgray", "darkgreen", "darkgrey", "darkkhaki", "darkmagenta", "darkolivegreen", "darkorange", "darkorchid", "darkred", "darksalmon", "darkseagreen", "darkslateblue", "darkslategray", "darkslategrey", "darkturquoise", "darkviolet", "deeppink", "deepskyblue", "dimgray", "dimgrey", "dodgerblue", "firebrick", "floralwhite", "forestgreen", "fuchsia", "gainsboro", "ghostwhite", "gold", "goldenrod", "gray", "green", "greenyellow", "grey", "honeydew", "hotpink", "indianred", "indigo", "ivory", "khaki", "lavender", "lavenderblush", "lawngreen", "lemonchiffon", "lightblue", "lightcoral", "lightcyan", "lightgoldenrodyellow", "lightgray", "lightgreen", "lightgrey", "lightpink", "lightsalmon", "lightseagreen", "lightskyblue", "lightslategray", "lightslategrey", "lightsteelblue", "lightyellow", "lime", "limegreen", "linen", "magenta", "maroon", "mediumaquamarine", "mediumblue", "mediumorchid", "mediumpurple", "mediumseagreen", "mediumslateblue", "mediumspringgreen", "mediumturquoise", "mediumvioletred", "midnightblue", "mintcream", "mistyrose", "moccasin", "navajowhite", "navy", "oldlace", "olive", "olivedrab", "orange", "orangered", "orchid", "palegoldenrod", "palegreen", "paleturquoise", "palevioletred", "papayawhip", "peachpuff", "peru", "pink", "plum", "powderblue", "purple", "red", "rosybrown", "royalblue", "saddlebrown", "salmon", "sandybrown", "seagreen", "seashell", "sienna", "silver", "skyblue", "slateblue", "slategray", "slategrey", "snow", "springgreen", "steelblue", "tan", "teal", "thistle", "tomato", "turquoise", "violet", "wheat", "white", "whitesmoke", "yellow", "yellowgreen"];
}


/**
 * Validate the supplied saved preferences for the actions.
 *
 * @param {Object} data the preferences
 * @returns {boolean|String} true for valid, otherwise string of errors.
 */
function validatePreferences(data) {

  let errors = "";

  // Attempt to get the spreadsheet
  let compareSpreadsheet;
  let compareSheet;
  try {
    compareSpreadsheet = SpreadsheetApp.openByUrl(data.sheetURL);

    // Attempt to get the sheet within the spreadsheet
    try {
      compareSheetId = compareSpreadsheet.getSheetByName(data.sheetName).getSheetId();
    } catch (err) {
      Logger.log(err.stack);
      errors = errors.concat("<br>That sheet name isn't valid. Make sure you validate the URL.")
    }

  } catch (err) {
    Logger.log(err.stack);
    errors = errors.concat("<br>That spreadsheet URL isn't valid.")
  }

  // Check that the colour is valid
  if (data.highlightNewRowColour === undefined || ((data.highlightNewRowColour.match(/^#(?:[0-9a-fA-F]{3}){1,2}$/) === null ||
    data.highlightNewRowColour.match(/^#(?:[0-9a-fA-F]{3}){1,2}$/).length !== 0) &&
    !getColourNames().includes(data.highlightNewRowColour))) {
    errors = errors.concat("<br>You must specify a valid colour.")
  }

  if (errors === "") {
    Logger.log("No errors found.")
    return true;
  } else {
    Logger.log(`Errors were found: ${JSON.stringify(errors)}`)
    return errors;
  }

}


/**
 * Processes the form submission from the Sheets sidebar.
 *
 * @param {eventObject} e The event object.
 * @returns {ActionResponse} The response, either the homepage or a notification
 */
function processSheetsSidebarForm(e) {

  let errors;
  let sheetURL;
  if ("sheetURL" in e.formInput) {
    sheetURL = e.formInput.sheetURL;
    errors = validatePreferences(e.formInput);
  } else {
    sheetURL = getUserProperties().sheetURL;
    let url = {"sheetURL": getUserProperties().sheetURL}
    data = Object.assign(url, e.formInput)
    errors = validatePreferences(data);
  }

  if (errors !== true) {

    // Format the errors in red
    errors = `<font color="#ff0000"><b>Uh oh!</b>${errors}</font>`;

  } else {

    // Only save if they're valid
    saveUserProperties(e.formInput);

    // Create a success message
    const spreadsheet = SpreadsheetApp.openByUrl(sheetURL);
    const sheetUrl = `${spreadsheet.getUrl()}#gid=${spreadsheet.getSheetByName(e.formInput.sheetName).getSheetId()}`
    errors = `<font color="#008000"><b>Success!</b><br>All values have been validated and saved. You can now run the actions via the 'Scripts' menu.</font><br><a href="${sheetUrl}">Open the spreadsheet.</a>`;

  }

  // Build the homepage
  const header = CardService.newCardHeader()
    .setTitle("Process your Society Ledger")
    .setImageUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/icon.png");
  const card = CardService.newCardBuilder()
    .setHeader(header);

  card.addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(errors)))
    .addSection(getSelectSheetSection())
    .addSection(getActionsSection())
    .addSection(createButtonsSection())
    .addSection(createDisclaimerSection());

  const navigation = CardService.newNavigation()
    .updateCard(card.build());
  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .build()

}
