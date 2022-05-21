/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Builds the homepage Card specifically for Drive.
 * If the user has selected a folder they'll be prompted to save the
 * ledger to it. If they've selected a PDF they'll be prompted to update
 * it. Otherwise they'll be asked to select a folder or PDF.
 *
 * @param {eventObject} e The event object
 * @returns {Card} The homepage Card
 */
function buildDriveHomePage(e) {

  // Start to create the Card
  const header = CardService.newCardHeader()
    .setTitle("Download your Society Ledger")
    .setImageUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/icon.png");
  const card = CardService.newCardBuilder()
    .setHeader(header);

  // If we have a previous result then display and clear it
  if (getUserProperties().resultMessage !== undefined) {
    const result = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(getUserProperties().resultMessage));
    card.addSection(result);
    PropertiesService.getUserProperties().deleteProperty("resultMessage");
  }

  // If they haven't selected anything yet
  // Or they've selected something that isn't a PDF nor a folder
  // Just ask them to select something
  if (Object.keys(e.drive).length === 0 ||
    (e.drive.activeCursorItem.mimeType !== "application/vnd.google-apps.folder" &&
      e.drive.activeCursorItem.mimeType !== "application/pdf")) {

    const instructions = CardService.newTextParagraph()
      .setText("Please select a folder to save the ledger to, or a PDF file to update.");
    card.addSection(CardService.newCardSection()
      .addWidget(instructions))
      .addSection(createDisclaimerSection());

    return card.build();

  }

  const currentFile = e.drive.activeCursorItem;

  let buttons = CardService.newButtonSet();

  // If we already have access to it
  if (currentFile.addonHasFileScopePermission) {
    buttons
      .addButton(CardService.newTextButton()
        .setText("DOWNLOAD")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("processDriveSidebarForm"))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .addButton(createClearSavedDataButton());

    // Otherwise we'll need to request access
  } else {
    buttons
      .addButton(CardService.newTextButton()
        .setText("AUTHORISE ACCESS")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("onRequestFileScopeButtonClicked")
          .setParameters({id: currentFile.id}))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .addButton(createClearSavedDataButton());
  }

  // If they've selected a folder
  if (currentFile.mimeType === "application/vnd.google-apps.folder") {

    // Tell the user that it'll be saved to the folder
    const folder = CardService.newKeyValue()
      .setContent(currentFile.title)
      .setIconUrl(currentFile.iconUrl)
      .setTopLabel("The ledger will be saved to this folder.");

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(folder)
      .addWidget(buttons);

    card.addSection(getLoginSection())
      .addSection(form)
      .addSection(createDisclaimerSection());

  } else if (currentFile.mimeType === "application/pdf") {

    // Tell the user that it'll be saved to the file
    const file = CardService.newKeyValue()
      .setContent(currentFile.title)
      .setIconUrl(currentFile.iconUrl)
      .setTopLabel("The ledger will be saved to this file.");

    const saveMethod = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setFieldName("saveMethod")
      .addItem("Rename this existing PDF", "rename", getUserProperties().saveMethod === "rename")
      .addItem("Keep the current filename", "keep", getUserProperties().saveMethod === "keep");

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(file)
      .addWidget(saveMethod)
      .addWidget(buttons);

    card.addSection(getLoginSection())
      .addSection(form)
      .addSection(createDisclaimerSection());

  }

  return card.build();

}


/**
 * Build and return the section with the login form.
 * This will ensure that any undefined (or unsaved) form inputs are replaced with empty strings.
 *
 * @returns {CardSection} A section with the login form.
 */
function getLoginSection() {

  const section = CardService.newCardSection();

  const groupId = CardService.newTextInput()
    .setFieldName("groupId")
    .setTitle("Society Group ID")
    .setHint("You can find this from your society admin URL (e.g. 12345 in exeterguild.org/organisation/admin/12345).")
    .setValue("");

  if (getUserProperties().groupId === undefined) {
    groupId.setValue("");
  } else {
    groupId.setValue(getUserProperties().groupId);
  }


  section.addWidget(groupId);

  // If we don't have all the login details
  if (getUserProperties().email === undefined ||
    getUserProperties().password === undefined) {

    const email = CardService.newTextInput()
      .setFieldName("email")
      .setTitle("eXpense365 Email")
      .setValue("");

    const password = CardService.newTextInput()
      .setFieldName("password")
      .setTitle("eXpense365 Password")
      .setValue("");

    section.addWidget(email).addWidget(password);

  } else {
    const message = CardService.newTextParagraph().setText("To change your login details, clear your saved data.");
    section.addWidget(message);
  }

  const saveDetails = CardService.newTextParagraph()
    .setText("Your details will be saved for you, and for your use only.");

  section.addWidget(saveDetails);

  return section;


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