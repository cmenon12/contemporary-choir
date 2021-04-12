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

  return card.addSection(getOriginalSheetSection()).build();

}


/**
 * Build and return the section with the login form.
 * This will ensure that any undefined (or unsaved) form inputs are replaced with empty strings.
 *
 * @returns {CardSection} A section with the login form.
 */
function getOriginalSheetSection() {

  let url = CardService.newTextInput()
    .setFieldName("originalSheetURL")
    .setTitle("Spreadsheet URL");

  if (getUserProperties().originalSheetURL === undefined) {
    url = url.setValue("");
  } else {
    url = url.setValue(getUserProperties().originalSheetURL);
  }

  let name = CardService.newTextInput()
    .setFieldName("originalSheetName")
    .setTitle("Sheet Name");

  if (getUserProperties().originalSheetName === undefined) {
    name = name.setValue("");
  } else {
    url = url.setValue(getUserProperties().originalSheetName);
  }

  return CardService.newCardSection()
  .addWidget(CardService.newTextParagraph()
    .setText("<b>The Google Sheet to compare against.</b>"))
  .addWidget(url)
  .addWidget(name)
  .addWidget(CardService.newTextParagraph()
    .setText("These details will be saved for you, and for your use only."));

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