var spreadsheet = SpreadsheetApp.getActive();
var sheet = spreadsheet.getActiveSheet();

/**
 * Does everything at once: the formatting and the comparison.
 * This uses the URL at the named range called DefaultUrl
 * The user is alerted when it is completed (so that
 * they know it hasn't stalled).
 */
function processWithDefaultUrl() {
  formatNeatly();

  // Locate the named range
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var url;
  for (var i=0; i<namedRanges.length; i++) {
    if (namedRanges[i].getName() == "DefaultUrl") {
      url = namedRanges[i].getRange().getValue();
    }
  }

  compareLedgers(url);
  copyToLedger(url);
  var ui = SpreadsheetApp.getUi();
  ui.alert("Complete!",
           "The process completed successfully. " +
           "You can view the logs here: " +
           "https://script.google.com/home/projects/1ycYC3ziL1kmxVl-RlqJhO-4UXZXjHw0Nljnal1IIfkWNDG-MPooyufqr/executions",
           ui.ButtonSet.OK);
}

/**
 * Just tests if the String value is a date in the form (DD/MM/YYYY).
 * Note that D M and Y could be any digit, not necessarily valid ones.
 */
function isADate(value) {
  return (String(value).match("\\d\\d/\\d\\d/\\d\\d\\d\\d") != null);
}

/**
 * Determines whether or not the row number in the current sheet
 * appears in the oldSheet.
 * Returns true if it's new, or false if it's old.
  */
function compareWithOld(row, oldSheetValues, newSheet) {

  // If newSheet has been supplied then use it, otherwise use the default
  var rowValues;
  if (newSheet == undefined) {
    rowValues = sheet.getSheetValues(row, 1, 1, 4)[0];
  } else {
    rowValues = newSheet.getSheetValues(row, 1, 1, 4)[0];
  }
  var newRow = true;
  for (i=0; i<oldSheetValues.length; i+=1) {
    if (rowValues.toString() == oldSheetValues[i].toString()) {
      newRow = false;
      break;
    }
  }
  return newRow;
}

/**
 * Prompts the user to enter the URL of the Google Sheet with the Original
 * sheet to compare with.
 * They can choose not to enter a URL and use the default one.
 * This will return the URL or false.
 */
function showPrompt() {

  // Create the prompt and save the result
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      "Do you want to use the default URL?",
      "If No then enter a new one, otherwise the default will be used. " +
      "Click Cancel to abort.",
      ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response
  var button = result.getSelectedButton();
  var url = result.getResponseText();

  // If the user wants to use a different URL
  if (button == ui.Button.NO) {
    return url;

  // If the user wants to use the default URL
  } else if (button == ui.Button.YES) {
    url = spreadsheet.getNamedRanges()[0].getRange().getValue();  // There's only one named range.
    return url
  } else {
    return false;
  }
}

/**
 * Compares the sheet with the sheet from another spreadsheet and highlights
 * differences in this sheet.
 * The sheet in the other spreadsheet must be called Original.
 * The sheet in this sheet must be the newer version. The Original sheet is
 * not modified.
 * This will highlight all the rows, and unhighlight them as it progresses.
 * Changed rows will be highlighted in red.
 * Note that any differences in whitespace will be recognised as a difference.
*/
function compareLedgers(url) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  var originalSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log("Script has opened spreadsheet " + url);
  var originalSheet = originalSpreadsheet.getSheetByName("Original");
  var originalSheetValues = originalSheet.getSheetValues(1, 1, originalSheet.getLastRow(), 4);

  var passedHeader = false;
  var cell;
  var cellValue;
  for (row = 1; row<=sheet.getLastRow(); row+=1) {
    cell = sheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't pased the first header row then skip it
    if (passedHeader == false) {
      if (cellValue == "Date") {
        passedHeader = true;
        sheet.getRange(row+1, 1, sheet.getLastRow()-row-1).setBackground("green");
      }


    // Compare it with the original/old sheet
    // Comparing all rows allows us to identify changes in the totals too
    } else {
      var isNew = compareWithOld(row, originalSheetValues);

      // If it is a new row then colour it
      if (isNew) {
        sheet.getRange(row, 1, 1, 4).setBackground("red");
        Logger.log("Row " + row + " is a new row!");

      // Otherwise just reset it
      } else {
        sheet.getRange(row, 1).setBackground("white");
      }
    }
  }
  Logger.log("Finished comparing sheets!")
}


/**
 * Copies the sheet to another spreadsheet file.
 * This is normally the one with the ledger in it.
 * The script then deletes the old Original,
 * and renames the new Original and sets protections on it.
 */
function copyToLedger(url) {

  // Gets the ledger spreadsheet
  var ledgerSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log("Script has opened spreadsheet " + url);

  // Copies to the ledger spreadsheet
  var newSheet = sheet.copyTo(ledgerSpreadsheet);

  // Remove protections from the old Original sheet and delete it
  var oldOriginalSheet = ledgerSpreadsheet.getSheetByName("Original")
  oldOriginalSheet.protect().remove();
  ledgerSpreadsheet.deleteSheet(oldOriginalSheet);

  // Rename the new Original sheet and protect it
  newSheet.setName("Original");
  newSheet.protect().setWarningOnly(true);

  Logger.log("Finished copying the sheet to the ledger spreadsheet.")


}

/**
 * Prompt the user for a URL and then run compareLedgers(url).
 */
function compareLedgersGetUrl() {
  var url = showPrompt()
  if (url == false) {
    return;
  }
  compareLedgers(url)
}

/**
 * Prompt the user for a URL and then run copyToLedger(url).
 */
function copyToLedgerGetUrl() {
  var url = showPrompt()
  if (url == false) {
    return;
  }
  copyToLedger(url)
}

/**
 * Creates the menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts")
      .addItem("Format neatly", "formatNeatly")
      .addItem("Highlight new entries", "compareLedgersGetUrl")
      .addItem("Copy to the ledger", "copyToLedgerGetUrl")
      .addItem("Process entirely with the default URL", "processWithDefaultUrl")
      .addToUi();
}


/**
 * Runs formatNeatlyWithSheet() with the active sheet.
 */
function formatNeatly() {
  formatNeatlyWithSheet(sheet)
}


/**
 * Formats the ledger neatly.
 * This can be used with any sheet (not necessarily the active one)
 * This function renames the sheet, resizes the columns,
 * removes unnecesary headers, and removes excess columns & rows.
 */
function formatNeatlyWithSheet(thisSheet) {

  // Convert the first column to text
  thisSheet.getRange(1, 1, thisSheet.getLastRow()).setNumberFormat("@")

  // Change the tab colour to be white
  thisSheet.setTabColor("white")

  // Rename the sheet
  var dateString = "\\d\\d\\/\\d\\d\\/\\d\\d\\d\\d \\d\\d:\\d\\d:\\d\\d"
  var finder = thisSheet.createTextFinder(dateString).matchEntireCell(true).useRegularExpression(true)
  var foundRange = finder.findNext()
  foundRange.setNumberFormat("@")
  var datetime = foundRange.getValue()
  thisSheet.setName(datetime)

  // Resize the columns
  finder = thisSheet.createTextFinder("Please note recent transactions may not be included.")
  foundRange = finder.findNext()
  foundRange.setValue("")
  thisSheet.autoResizeColumns(1, 4)

  // Get all occurrences of UNIV01 (except for the first one)
  finder = thisSheet.createTextFinder("UNIV01")
  var matches = finder.findAll()
  matches.shift()
  matches.reverse() // start at the bottom to avoid changing future ranges

  // Remove each of these headers
  // If the totals are on a new page with no entries, then we should delete one less row
  for (i = 0; i<matches.length; i+=1) {
    row = matches[i].getRow()-2
    if (matches[i].offset(3,0).getValue() != "") {
      thisSheet.deleteRows(row,5)
      Logger.log("Deleted five rows starting at row " + row)
    } else {
      thisSheet.deleteRows(row, 6)
      Logger.log("Deleted six rows starting at row " + row)
    }
  }

  // Remove all the excess rows & columns.
  try {
    thisSheet.deleteRows(thisSheet.getLastRow()+5, thisSheet.getMaxRows()-thisSheet.getLastRow()-5)
    thisSheet.deleteColumns(thisSheet.getLastColumn()+1, thisSheet.getMaxColumns()-thisSheet.getLastColumn()-1)
  } catch (error){
    Logger.log(error.message)
  }

  return true;
}