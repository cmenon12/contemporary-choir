/**
 * Build the homepage common to all non-Drive add-ons.
 */
function buildCommonHomePage(e) {

  Logger.log(e)

  const header = CardService.newCardHeader().setTitle("Download your society ledger");
  const card = CardService.newCardBuilder().setHeader(header)

  const saveMethod = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("How do you want to save the PDF ledger?")
    .setFieldName("saveMethod")
    .addItem("Save to this folder", "folder", getUserProperties().saveMethodFolder === "true")
    .addItem("Replace this existing PDF", "replace", getUserProperties().saveMethodReplace === "true")
    .addItem("Add a new revision to this existing PDF", "update", getUserProperties().saveMethodUpdate === "true");

  const url = CardService.newTextInput()
    .setFieldName("url")
    .setTitle("URL of folder or PDF file in Drive")
    .setValue("");

  const buttons = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText("DOWNLOAD")
      .setOnClickAction(CardService.newAction().setFunctionName("processSidebarForm")))
    .addButton(CardService.newTextButton()
      .setText("Clear saved data")
      .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")))

  const form = CardService.newCardSection()
    .addWidget(saveMethod)
    .addWidget(url)
    .addWidget(buttons)

  card.addSection(getLoginSection()).addSection(form);

  return card.build();

}


/**
 * Build the homepage for Drive.
 */
function buildDriveHomePage(e) {

  Logger.log(e)

  // Start to create the Card
  const header = CardService.newCardHeader().setTitle("Download your society ledger");
  const card = CardService.newCardBuilder().setHeader(header)

  // If they haven't selected anything yet
  // Or they've selected something that isn't a PDF nor a folder
  // Just ask them to select something
  if (Object.keys(e.drive).length === 0 ||
    (e.drive.activeCursorItem.mimeType !== "application/vnd.google-apps.folder" &&
      e.drive.activeCursorItem.mimeType !== "application/pdf")) {

    const instructions = CardService.newTextParagraph()
      .setText("Please select a folder to save the ledger to, or a PDF file to update.");
    card.addSection(CardService.newCardSection()
      .addWidget(instructions));
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

    let buttons;

    // If we already have access to it
    if (currentFile.addonHasFileScopePermission) {
      buttons = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("DOWNLOAD")
          .setOnClickAction(CardService.newAction().setFunctionName("processSidebarForm")))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")))

      // Otherwise we'll need to request access
    } else {
      buttons = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("AUTHORISE ACCESS")
          .setOnClickAction(CardService.newAction().setFunctionName("onRequestFileScopeButtonClicked").setParameters({id: currentFile.id})))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")))
    }

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(folder)
      .addWidget(buttons)

    card.addSection(getLoginSection()).addSection(form);

  } else if (currentFile.mimeType === "application/pdf") {

    // Tell the user that it'll be saved to the folder
    const file = CardService.newKeyValue()
      .setContent(currentFile.title)
      .setIconUrl(currentFile.iconUrl)
      .setTopLabel("The ledger will be saved to this file.");

    const saveMethod = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setFieldName("saveMethod")
      .addItem("Replace this existing PDF", "replace", getUserProperties().saveMethodReplace === "true")
      .addItem("Add a new revision to this existing PDF", "update", getUserProperties().saveMethodUpdate === "true");

    let buttons;

    // If we already have access to it
    if (currentFile.addonHasFileScopePermission) {
      buttons = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("DOWNLOAD")
          .setOnClickAction(CardService.newAction().setFunctionName("processSidebarForm")))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));

      // Otherwise we'll need to request access
    } else {
      buttons = CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText("AUTHORISE ACCESS")
          .setOnClickAction(CardService.newAction().setFunctionName("onRequestFileScopeButtonClicked").setParameters({id: currentFile.id})))
        .addButton(CardService.newTextButton()
          .setText("Clear saved data")
          .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")));
    }

    // Finish creating the form
    const form = CardService.newCardSection()
      .addWidget(file)
      .addWidget(saveMethod)
      .addWidget(buttons)

    card.addSection(getLoginSection()).addSection(form);

  }

  return card.build();

}


/**
 * Build and returnthe section with the login details.
 */
function getLoginSection() {

  const groupId = CardService.newTextInput()
    .setFieldName("groupId")
    .setTitle("Group ID")
    .setHint("You can find this from your society admin URL (e.g. 12345 in exeterguild.org/organisation/admin/12345).")
    .setValue(getUserProperties().groupId);

  const email = CardService.newTextInput()
    .setFieldName("email")
    .setTitle("Email")
    .setValue(getUserProperties().email);

  const password = CardService.newTextInput()
    .setFieldName("password")
    .setTitle("Password")
    .setValue(getUserProperties().password);

  const saveDetails = CardService.newTextParagraph().setText("Your details will be saved for you, and for your use only.");

  const section = CardService.newCardSection()
    .addWidget(groupId)
    .addWidget(email)
    .addWidget(password)
    .addWidget(saveDetails)

  return section

}