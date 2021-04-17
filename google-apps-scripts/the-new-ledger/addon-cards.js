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

  return card
    .addSection(getOriginalSheetSection())
    .addSection(getActionsSection())
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
 * Save the formInput and update the current card with the homepage.
 *
 * @param {eventObject} e The event object
 * @returns {Card} The homepage Card
 */
function updateWithHomepage(e) {

  saveUserProperties(e.formInput);
  const navigation = CardService.newNavigation()
    .updateCard(buildSheetsHomePage(e));
  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .build();

}


/**
 * Build and return the section with the selection input for the actions
 *
 * @returns {CardSection} A section with the selection input for the actions.
 */
function getActionsSection() {

  const selection = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("actions")
    .addItem("Format neatly", "format", getUserProperties().actionsFormat === "on")
    .addItem("Highlight new entries", "highlight", getUserProperties().actionsHighlight === "on")
    .addItem("Copy to the ledger", "copy", getUserProperties().actionsCopy === "on")

  const buttons = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("RUN")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("processSheetsSidebarForm"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
    .addButton(createClearSavedDataButton());

  return CardService.newCardSection()
    .addWidget(selection)
    .addWidget(buttons);

}


/**
 * Build and return the section with the original sheet form.
 * Ensures that any undefined form inputs are replaced with empty strings.
 * Also attempts to get the name and URL of the saved sheet.
 *
 * @returns {CardSection} A section with the original sheet form.
 */
function getOriginalSheetSection() {

  // Create the URL text input
  let url = CardService.newTextInput()
    .setFieldName("originalSheetURL")
    .setTitle("Spreadsheet URL")
    .setOnChangeAction(CardService.newAction()
      .setFunctionName("updateWithHomepage"));

  // Fill the URL text input with the saved data
  if (getUserProperties().originalSheetURL === undefined) {
    url = url.setValue("");
  } else {
    url = url.setValue(getUserProperties().originalSheetURL);
  }

  // Create the sheet name selection input
  let name = CardService.newSelectionInput()
    .setFieldName("originalSheetName")
    .setTitle("Sheet Name")
    .setType(CardService.SelectionInputType.DROPDOWN)

  // Start to create the spreadsheet identifier
  const selected = CardService.newDecoratedText();

  // If the URL is present then attempt to get the spreadsheet name
  let validURL = false;
  if (getUserProperties().originalSheetURL !== "" &&
    getUserProperties().originalSheetURL !== undefined) {
    try {
      const spreadsheet = SpreadsheetApp.openByUrl(getUserProperties()
        .originalSheetURL);
      selected.setText(spreadsheet.getName());
      selected.setStartIcon(CardService.newIconImage()
        .setIconUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/outline_task_black_48dp.png"));

      // Add the sheet names
      let sheetName;
      for (let i = 0; i < spreadsheet.getNumSheets(); i += 1) {
        sheetName = spreadsheet.getSheets()[i].getName();
        if (sheetName === getUserProperties().originalSheetName) {
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
  // Only add the name selection input if the URL was valid
  const section = CardService.newCardSection()
    .addWidget(url)
    .addWidget(selected)
  if (validURL === true) {
    section.addWidget(name)
  }
  return section.addWidget(CardService.newTextParagraph()
    .setText("These details will be saved for you, and for your use only."));

}


/**
 * Creates a text button for Cards that clears the saved data.
 * @returns {TextButton} The button.
 */
function createClearSavedDataButton() {
  return CardService.newTextButton()
    .setText("Clear saved data")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("clearSavedData"))
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