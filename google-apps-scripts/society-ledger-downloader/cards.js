/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Build the homepage common to all non-Drive add-ons.
 * This has been disabled in the manifest.
 */
function buildCommonHomePage(e) {

  const header = CardService.newCardHeader()
    .setTitle("Download your Society Ledger")
    .setImageUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/icon.png");
  const card = CardService.newCardBuilder()
    .setHeader(header);
  const disclaimerText = '<font color="#bdbdbd">Icon made by <a href="https://www.flaticon.com/authors/photo3idea-studio">photo3idea_studio</a> from <a href="https://www.flaticon.com/">www.flaticon.com</a>.</font>';
  const disclaimer = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText(disclaimerText));

  const saveMethod = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("How do you want to save the PDF ledger?")
    .setFieldName("saveMethod")
    .addItem("Save to this folder", "folder", getUserProperties().saveMethod === "folder")
    .addItem("Rename this existing PDF", "rename", getUserProperties().saveMethod === "rename")
    .addItem("Keep the current filename", "keep", getUserProperties().saveMethod === "keep");

  const url = CardService.newTextInput()
    .setFieldName("url")
    .setTitle("URL of folder or PDF file in Drive")
    .setValue("");

  const buttons = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("DOWNLOAD")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("processSidebarForm"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
    .addButton(CardService.newTextButton()
      .setText("Clear saved data")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("clearSavedData")));

  const form = CardService.newCardSection()
    .addWidget(saveMethod)
    .addWidget(url)
    .addWidget(buttons);

  card.addSection(getLoginSection())
    .addSection(form)
    .addSection(disclaimer);

  return card.build();

}


/**
 * Build the homepage for Drive.
 */
function buildDriveHomePage(e) {

  // Start to create the Card
  const header = CardService.newCardHeader()
    .setTitle("Download your Society Ledger")
    .setImageUrl("https://raw.githubusercontent.com/cmenon12/contemporary-choir/main/assets/icon.png");
  const card = CardService.newCardBuilder()
    .setHeader(header);
  const disclaimerText = '<font color="#bdbdbd">Icon made by <a href="https://www.flaticon.com/authors/photo3idea-studio">photo3idea_studio</a> from <a href="https://www.flaticon.com/">www.flaticon.com</a>.</font>';
  const disclaimer = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText(disclaimerText));

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
      .addSection(disclaimer);

    return card.build();

  }

  const currentFile = e.drive.activeCursorItem;

  // If they've selected a folder
  if (currentFile.mimeType === "application/vnd.google-apps.folder") {

    // Tell the user that it'll be saved to the folder
    const folder = CardService.newKeyValue()
      .setContent(currentFile.title)
      .setIconUrl(currentFile.iconUrl)
      .setTopLabel("The ledger will be saved to this folder.");

    let buttons = CardService.newButtonSet();

    // If we already have access to it
    if (currentFile.addonHasFileScopePermission) {
      buttons
        .addButton(CardService.newTextButton()
          .setText("DOWNLOAD")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("processDriveSidebarForm"))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));

      // Otherwise we'll need to request access
    } else {
      buttons
        .addButton(CardService.newTextButton()
          .setText("AUTHORISE ACCESS")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("onRequestFileScopeButtonClicked")
            .setParameters({id: currentFile.id}))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));
    }

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(folder)
      .addWidget(buttons);
    card.addSection(getLoginSection()).addSection(form).addSection(disclaimer);

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

    let buttons = CardService.newButtonSet()

    // If we already have access to it
    if (currentFile.addonHasFileScopePermission) {
      buttons
        .addButton(CardService.newTextButton()
          .setText("DOWNLOAD")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("processDriveSidebarForm"))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));

      // Otherwise we'll need to request access
    } else {
      buttons
        .addButton(CardService.newTextButton()
          .setText("AUTHORISE ACCESS")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("onRequestFileScopeButtonClicked")
            .setParameters({id: currentFile.id}))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));
    }

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(file)
      .addWidget(saveMethod)
      .addWidget(buttons);

    card.addSection(getLoginSection()).addSection(form).addSection(disclaimer);

  }

  return card.build();

}


/**
 * Build and return the section with the login details.
 */
function getLoginSection() {

  let groupId = CardService.newTextInput()
    .setFieldName("groupId")
    .setTitle("Society Group ID")
    .setHint("You can find this from your society admin URL (e.g. 12345 in exeterguild.org/organisation/admin/12345).");

  if (getUserProperties().groupId === undefined) {
    groupId = groupId.setValue("");
  } else {
    groupId = groupId.setValue(getUserProperties().groupId);
  }

  let email = CardService.newTextInput()
    .setFieldName("email")
    .setTitle("eXpense365 Email");

  if (getUserProperties().email === undefined) {
    email = email.setValue("");
  } else {
    email = email.setValue(getUserProperties().email);
  }

  let password = CardService.newTextInput()
    .setFieldName("password")
    .setTitle("eXpense365 Password");

  if (getUserProperties().password === undefined) {
    password = password.setValue("");
  } else {
    password = password.setValue(getUserProperties().password);
  }

  const saveDetails = CardService.newTextParagraph()
    .setText("Your details will be saved for you, and for your use only.");

  const section = CardService.newCardSection()
    .addWidget(groupId)
    .addWidget(email)
    .addWidget(password)
    .addWidget(saveDetails);

  return section

}


function buildErrorNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(message))
    .build();
}