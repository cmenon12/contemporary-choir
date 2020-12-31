/**
 * Build the homepage common to all non-Drive add-ons.
 */
function buildCommonHomePage(e) {
  
  const header = CardService.newCardHeader().setTitle("Download your society ledger");
  
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
  
  const saveMethod = CardService.newSelectionInput()
  .setType(CardService.SelectionInputType.RADIO_BUTTON)
  .setTitle("How do you want to save the PDF ledger?")
  .setFieldName("saveMethod")
  .addItem("Save to this folder", "folder", getUserProperties().saveMethodFolder == "true")
  .addItem("Replace this existing PDF", "replace", getUserProperties().saveMethodReplace == "true")
  .addItem("Add a new revision to this existing PDF", "update", getUserProperties().saveMethodUpdate == "true");
  
  const url = CardService.newTextInput()
  .setFieldName("url")
  .setTitle("URL of folder or PDF file in Drive")
  .setValue("");
  
  const saveDetails = CardService.newKeyValue()
  .setBottomLabel("This will be saved for you individually.")
    .setContent("Save for next time")
    .setSwitch(CardService.newSwitch()
               .setFieldName("saveDetails")
               .setValue("true")
               .setSelected(getUserProperties().saveDetails == "true"));
  
  const buttons = CardService.newButtonSet()
  .addButton(CardService.newTextButton()
             .setText("DOWNLOAD")
             .setOnClickAction(CardService.newAction().setFunctionName("processSidebarForm")))
  .addButton(CardService.newTextButton()
             .setText("Clear saved data")
             .setOnClickAction(CardService.newAction().setFunctionName("clearSavedData")))
  
  const form = CardService.newCardSection().addWidget(groupId).addWidget(email).addWidget(password)
  .addWidget(saveMethod).addWidget(url).addWidget(saveDetails).addWidget(buttons)
  
  const card = CardService.newCardBuilder().setHeader(header).addSection(form);
  
  return card.build();

}


function buildDriveHomePage(e) {

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

  Logger.log(formData);
}
