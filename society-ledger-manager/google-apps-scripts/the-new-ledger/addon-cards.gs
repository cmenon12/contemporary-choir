/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

/**
 * Builds the homepage Card specifically for Sheets.
 *
 * @param {eventObject} e The event object
 * @returns {Card} The homepage Card
 */
function buildSheetsHomePage(e) {

  // Start to create the Card
  const header = CardService.newCardHeader()
    .setTitle("Process your Society Ledger")
    .setImageUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/icon.png");
  const card = CardService.newCardBuilder()
    .setHeader(header);

  // Let the user know if they're already valid
  if (validatePreferences(getUserProperties()) === true) {
    const spreadsheet = SpreadsheetApp.openByUrl(getUserProperties().sheetURL);
    const sheetUrl = `${spreadsheet.getUrl()}#gid=${spreadsheet.getSheetByName(getUserProperties().sheetName).getSheetId()}`
    const message = `<font color="#008000"><b>Success!</b><br>All values have been validated and saved. You can now run the actions via the 'Scripts' menu.</font><br><a href="${sheetUrl}">Open the spreadsheet.</a>`;
    card.addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(message)));
  }

  return card
    .addSection(getSelectSheetSection())
    .addSection(getActionsSection())
    .addSection(createButtonsSection())
    .addSection(createDisclaimerSection())
    .build();

}


/**
 * Produces and displays a notification.
 *
 * @param {String} message The message to notify.
 * @returns {ActionResponse} The notification to display.
 */
function buildNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(message))
    .build();
}


/**
 * Validate the sheet URL.
 *
 * @param {eventObject} e The event object
 * @returns {Card} The homepage Card
 */
function validateSheetURL(e) {

  saveUserProperties({"sheetURL": e.formInput.sheetURL});

  const navigation = CardService.newNavigation()
    .updateCard(buildSheetsHomePage(e));
  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .build();

}


/**
 * Build and return the section with the form for the actions.
 *
 * @returns {CardSection} A section with the form for the actions.
 */
function getActionsSection() {

  const section = CardService.newCardSection()
  section.addWidget(CardService.newTextParagraph().setText("<b>PREFERENCES</b>"));

  // Formatting options
  const formatSheetName = CardService.newTextInput()
    .setFieldName("formatSheetName")
    .setTitle("The new name of this sheet")
    .setHint("Leave blank to keep the existing name. Use {{datetime}} for the date and time of the ledger.")
    .setValue(getUserProperties().formatSheetName === undefined ? "" : getUserProperties().formatSheetName);
  section.addWidget(formatSheetName);

  // Highlighting options
  const highlightNewRowColour = CardService.newTextInput()
    .setFieldName("highlightNewRowColour")
    .setTitle("The colour to highlight new entries")
    .setHint("A color code in CSS notation (such as '#E62073' or 'red')")
    .setSuggestions(CardService.newSuggestions().addSuggestions(getColourNames()))
    .setValue(getUserProperties().highlightNewRowColour === undefined ? "" : getUserProperties().highlightNewRowColour);
  section.addWidget(highlightNewRowColour);

  // Copying options
  const copyNewSheetName = CardService.newTextInput()
    .setFieldName("copyNewSheetName")
    .setTitle("The name of the copied sheet (optional)")
    .setHint("Leave blank to overwrite the sheet above. Use {{datetime}} for the date and time of the ledger.")
    .setValue(getUserProperties().copyNewSheetName === undefined ? "" : getUserProperties().copyNewSheetName);
  section.addWidget(copyNewSheetName);

  return section;

}


/**
 * Build and return the section with the sheet form.
 * Ensures that any undefined form inputs are replaced with empty strings.
 * Also attempts to get the name and URL of the saved sheet.
 *
 * @returns {CardSection} A section with the sheet form.
 */
function getSelectSheetSection() {

  // Create the URL text input
  let url = CardService.newTextInput()
    .setFieldName("sheetURL")
    .setTitle("Spreadsheet URL");

  // Fill the URL text input with the saved data
  if (getUserProperties().sheetURL === undefined) {
    url = url.setValue("");
  } else {
    url = url.setValue(getUserProperties().sheetURL);
  }

  // Create the sheet name selection input
  let name = CardService.newSelectionInput()
    .setFieldName("sheetName")
    .setTitle("Sheet Name")
    .setType(CardService.SelectionInputType.DROPDOWN);

  // Start to create the spreadsheet identifier
  const selected = CardService.newDecoratedText();

  // Create the validate URL button
  const button = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("VALIDATE URL")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("validateSheetURL")));

  // If the URL is present then attempt to get the spreadsheet name
  let validURL = false;
  if (getUserProperties().sheetURL !== "" &&
    getUserProperties().sheetURL !== undefined) {
    try {
      const spreadsheet = SpreadsheetApp.openByUrl(getUserProperties()
        .sheetURL);
      selected.setText(spreadsheet.getName());
      selected.setStartIcon(CardService.newIconImage()
        .setIconUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/outline_task_black_48dp.png"));

      // Add the sheet names to the selection dropdown
      let sheetName;
      for (let i = 0; i < spreadsheet.getNumSheets(); i += 1) {
        sheetName = spreadsheet.getSheets()[i].getName();
        if (sheetName === getUserProperties().sheetName) {
          name.addItem(sheetName, sheetName, true);
        } else {
          name.addItem(sheetName, sheetName, false);
        }
      }

      validURL = true;

      // Spreadsheet not found
    } catch (err) {
      selected.setText(`Spreadsheet not found`);
      selected.setStartIcon(CardService.newIconImage()
        .setIconUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/outline_error_outline_black_48dp.png"));
      url.setHint("Invalid.");
    }

    // If no URL is available
  } else {
    selected.setText(`No spreadsheet specified`);
    selected.setStartIcon(CardService.newIconImage()
      .setIconUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/outline_error_outline_black_48dp.png"));
  }

  // Build the section
  // Only add the URL text input and button if the URL was invalid
  // Only add the name selection input if the URL was valid
  const section = CardService.newCardSection()
  section.addWidget(CardService.newTextParagraph().setText("<b>LEDGER TO COMPARE</b>"))
  if (validURL === false) {
    section.addWidget(url)
      .addWidget(button)
  }
  section.addWidget(selected)
  if (validURL === true) {
    section.addWidget(name)
  }
  return section;

}


/**
 * Creates the buttons.
 * @returns {Section} the buttons section.
 */
function createButtonsSection() {

  return CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText("These details will be saved for you, and for your use only."))
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText("SAVE")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("processSheetsSidebarForm"))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .addButton(CardService.newTextButton()
        .setText("Clear saved data")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("clearSavedData"))));
}


/**
 * Creates the section with the standard disclaimer & attribution.
 * @returns {CardSection} The disclaimer.
 */
function createDisclaimerSection() {
  const disclaimerText = '<font color="#bdbdbd">Icon made by <a href="https://www.flaticon.com/authors/photo3idea-studio">photo3idea_studio</a> from <a href="https://www.flaticon.com/">www.flaticon.com</a>.</font>';
  return CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText(disclaimerText));
}