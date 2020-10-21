/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */

const spreadsheet = SpreadsheetApp.getActive();
const sheet = spreadsheet.getActiveSheet();

/**
 * Does everything at once: the formatting and the comparison.
 * This uses the URL at the named range called DefaultUrl
 * The user is alerted when it is completed (so that
 * they know it hasn't stalled).
 */
function processWithDefaultUrl() {
  formatNeatly();

  // Locate the named range
  const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  let url;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() == "DefaultUrl") {
      url = namedRanges[i].getRange().getValue();
    }
  }

  compareLedgers(url);
  copyToLedger(url);
  const ui = SpreadsheetApp.getUi();
  ui.alert("Complete!",
      "The process completed successfully. " +
      "You can view the logs here: " +
      "https://script.google.com/home/projects/1ycYC3ziL1kmxVl-RlqJhO-4UXZXjHw0Nljnal1IIfkWNDG-MPooyufqr/executions",
      ui.ButtonSet.OK);
}

/**
 * Open a URL in a new tab.
 * This is from StackOverflow
 * https://stackoverflow.com/questions/10744760/google-apps-script-to-open-a-url
 */
function openUrl(url) {
  const html = HtmlService.createHtmlOutput('<html lang="en-GB"><script>'
      + 'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
      + 'let a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";'
      + 'if(document.createEvent){'
      + '  let event=document.createEvent("MouseEvents");'
      + '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
      + '  event.initEvent("click",true,true); a.dispatchEvent(event);'
      + '}else{ a.click() }'
      + 'close();'
      + '</script>'
      // Offer URL as clickable link in case above code fails.
      + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. '
      + '<a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
      + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
      + '</html>')
      .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
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
function compareWithOld(row, oldSheetValues, newSheet = undefined) {

  // If newSheet has been supplied then use it, otherwise use the default
  let rowValues;
  if (newSheet == undefined) {
    rowValues = sheet.getSheetValues(row, 1, 1, 4)[0];
  } else {
    rowValues = newSheet.getSheetValues(row, 1, 1, 4)[0];
  }
  let newRow = true;
  for (let i = 0; i < oldSheetValues.length; i += 1) {
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
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
      "Do you want to use the default URL?",
      "If No then enter a new one, otherwise the default will be used. " +
      "Click Cancel to abort.",
      ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response
  const button = result.getSelectedButton();
  let url = result.getResponseText();

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
 * This will highlight all the rows, and un-highlight them as it progresses.
 * Changed rows will be highlighted in red.
 * Note that any differences in whitespace will be recognised as a difference.
 */
function compareLedgers(url) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  const originalSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log("Script has opened spreadsheet " + url);
  const originalSheet = originalSpreadsheet.getSheetByName("Original");
  const originalSheetValues = originalSheet.getSheetValues(1, 1, originalSheet.getLastRow(), 4);

  let passedHeader = false;
  let cell;
  let cellValue;
  for (let row = 1; row <= sheet.getLastRow(); row += 1) {
    cell = sheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't passed the first header row then skip it
    if (passedHeader == false) {
      if (cellValue == "Date") {
        passedHeader = true;
        sheet.getRange(row + 1, 1, sheet.getLastRow() - row - 1).setBackground("green");
      }


      // Compare it with the original/old sheet
      // Comparing all rows allows us to identify changes in the totals too
    } else {
      const isNew = compareWithOld(row, originalSheetValues);

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
  const ledgerSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log("Script has opened spreadsheet " + url);

  // Copies to the ledger spreadsheet
  const newSheet = sheet.copyTo(ledgerSpreadsheet);

  // Remove protections from the old Original sheet and delete it
  const oldOriginalSheet = ledgerSpreadsheet.getSheetByName("Original");
  oldOriginalSheet.protect().remove();
  ledgerSpreadsheet.deleteSheet(oldOriginalSheet);

  // Rename the new Original sheet and protect it
  newSheet.setName("Original");
  newSheet.protect().setWarningOnly(true);

  Logger.log("Finished copying the sheet to the ledger spreadsheet.")

  // Ask the user if they want to open the new sheet
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
      "Do you want to open the ledger?",
      "This will open the ledger in a new tab.",
      ui.ButtonSet.YES_NO);

  // Open it if the user want to
  if (result == ui.Button.YES) {
    const sheetUrl = url + "#gid=" + newSheet.getSheetId();
    openUrl(sheetUrl);
  }
}

/**
 * Prompt the user for a URL and then run compareLedgers(url).
 */
function compareLedgersGetUrl() {
  const url = showPrompt();
  if (url == false) {
    return;
  }
  compareLedgers(url)
}

/**
 * Prompt the user for a URL and then run copyToLedger(url).
 */
function copyToLedgerGetUrl() {
  const url = showPrompt();
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
 * removes unnecessary headers, and removes excess columns & rows.
 */
function formatNeatlyWithSheet(thisSheet) {

  // Convert the first column to text
  thisSheet.getRange(1, 1, thisSheet.getLastRow()).setNumberFormat("@")

  // Change the tab colour to be white
  thisSheet.setTabColor("white")

  // Rename the sheet
  const dateString = "\\d\\d\\/\\d\\d\\/\\d\\d\\d\\d \\d\\d:\\d\\d:\\d\\d";
  let finder = thisSheet.createTextFinder(dateString).matchEntireCell(true).useRegularExpression(true);
  let foundRange = finder.findNext();
  foundRange.setNumberFormat("@")
  const datetime = foundRange.getValue();
  thisSheet.setName(datetime)

  // Resize the columns
  finder = thisSheet.createTextFinder("Please note recent transactions may not be included.")
  foundRange = finder.findNext()
  foundRange.setValue("")
  thisSheet.autoResizeColumns(1, 4)

  // Get all occurrences of UNIV01 (except for the first one)
  finder = thisSheet.createTextFinder("UNIV01")
  const matches = finder.findAll();
  matches.shift()
  matches.reverse() // start at the bottom to avoid changing future ranges

  // Remove each of these headers
  // If the totals are on a new page with no entries, then we should delete one less row
  for (let i = 0; i < matches.length; i += 1) {
    let row = matches[i].getRow() - 2
    if (matches[i].offset(3, 0).getValue() != "") {
      thisSheet.deleteRows(row, 5)
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