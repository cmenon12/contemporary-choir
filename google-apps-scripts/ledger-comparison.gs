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
 * Just tests if the String value is a date in the form (DD/MM/YYYY).
 * Note that D M and Y could be any digit, not necessarily valid ones.
 *
 * @param {String} value the string to check.
 * @return {boolean} true if it's a date, otherwise false.
 */
function isADate(value) {
  return (String(value).match("\\d\\d/\\d\\d/\\d\\d\\d\\d") != null);
}


/**
 * Get the range with the name from the spreadsheet.
 *
 * @param {String} name the name of the range.
 * @param {Spreadsheet} spreadsheet the spreadsheet to search.
 * @return {Range|undefined} the named range if found, otherwise undefined.
 */
function getNamedRange(name, spreadsheet) {

  const namedRanges = spreadsheet.getNamedRanges();
  let range;
  for (let i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === name) {
      range = namedRanges[i].getRange();
      break;
    }
  }
  return range;
}


/**
 * Open a URL in a new tab using the openUrl.html HTML template.
 *
 * @param {String} url the URL to open.
 */
function openUrl(url) {

  // Create the HTML template and incorporate the URL
  const html = HtmlService.createTemplateFromFile("openUrl");
  html.url = url;

  // Create and show the dialog
  const dialog = html.evaluate().setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(dialog, "Opening ...");

}


/**
 * Prompts the user to enter the URL of the Google Sheet with the Original
 * sheet to compare with.
 * They can choose not to enter a URL and use the default one.
 *
 * @return {String|Boolean} the entered URL or false.
 */
function showPrompt() {

  // Create the prompt and save the result
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "Do you want to use the default URL?",
    "If No then enter a new one, otherwise the default will be used. " +
    "Click Cancel to abort.", ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response
  const button = result.getSelectedButton();
  let url = result.getResponseText();

  // If the user wants to use a different URL
  if (button === ui.Button.NO) {
    return url;

    // If the user wants to use the default URL
  } else if (button === ui.Button.YES) {
    url = spreadsheet.getNamedRanges()[0].getRange().getValue();  // There's only one named range.
    return url;
  } else {
    return false;
  }
}


/**
 * Formats the ledger neatly.
 * This function renames the sheet, resizes the columns,
 * removes unnecessary headers, and removes excess columns & rows.
 *
 * @param {Sheet} thisSheet the sheet to format.
 */
function formatNeatlyWithSheet(thisSheet) {

  // Don't re-run if it's already been done
  if (thisSheet.getRange("A1").getValue() === "Income/Expense Statement") {
    return;
  }

  // Convert the first column to text
  thisSheet.getRange(1, 1, thisSheet.getLastRow()).setNumberFormat("@");

  // Change the tab colour to be white
  thisSheet.setTabColor("white");

  // Rename the sheet
  datetime = `${thisSheet.getRange("C1").getValue()}${thisSheet.getRange("D1").getValue()}`;
  thisSheet.setName(datetime);

  // Delete the 'Please note' text at the end
  finder = thisSheet.getRange("A:A").createTextFinder("Please note");
  foundRange = finder.findNext();
  thisSheet.deleteRows(foundRange.getRow(), 1);

  // Remove all the excess rows & columns
  try {
    thisSheet.deleteRows(thisSheet.getLastRow() + 5,
      thisSheet.getMaxRows() - thisSheet.getLastRow() - 5);
    thisSheet.deleteColumns(thisSheet.getLastColumn() + 1,
      thisSheet.getMaxColumns() - thisSheet.getLastColumn() - 1);
  } catch (error) {
    Logger.log(`There was an error removing the excess rows & columns: ${error.message}`);
  }

  // Set the font size for all 11, but important parts to 12
  thisSheet.getRange(`1:${thisSheet.getMaxRows()}`).setFontSize(11);
  thisSheet.getRange(`${thisSheet.getLastRow() - 4}:${thisSheet.getLastRow()}`).setFontSize(12);
  thisSheet.getRange("1:4").setFontSize(12);

  // Replace all in-cell newlines with spaces
  finder = thisSheet.createTextFinder("\n").useRegularExpression(true).replaceAllWith(" ");

  // Resize the columns
  thisSheet.autoResizeColumns(1, 4);

  // Get all occurrences of UNIV01 (except for the first one)
  finder = thisSheet.createTextFinder("UNIV01");
  let matches = finder.findAll();
  matches.shift();
  matches.reverse(); // start at the bottom to avoid changing future ranges

  // Remove each of these headers
  // If the cost code row isn't present,
  // then we should delete one less row
  for (let i = 0; i < matches.length; i += 1) {
    let row = matches[i].getRow() - 1;
    if (matches[i].offset(3, 0).getValue() === "") {
      thisSheet.deleteRows(row, 5);
      Logger.log(`Deleted five rows starting at row ${row}`);
    } else {
      thisSheet.deleteRows(row, 4);
      Logger.log(`Deleted four rows starting at row ${row}`);
    }
  }

  // For each cell in the first column
  // If it's not a date then make it bold
  // If it's also not empty nor "Date" then
  // append the value in the next column and merge the two cells
  // This fixes words being split up
  let value;
  const boldOn = SpreadsheetApp.newTextStyle().setBold(true).build();
  for (let i = 1; i <= thisSheet.getLastRow(); i += 1) {
    value = thisSheet.getRange(`A${i}`).getValue();

    if (isADate(value) === false) {
      thisSheet.getRange(`${i}:${i}`).setTextStyle(boldOn);

      if (value !== "" && value !== "Date") {
        thisSheet.getRange(`A${i}`).setValue(value + thisSheet.getRange(`B${i}`).getValue());
        thisSheet.getRange(`B${i}`).setValue("");
        thisSheet.getRange(`A${i}:B${i}`).merge();
      }
    }
  }

  // Smarten up the top rows
  thisSheet.getRange("A3").setValue(`${thisSheet.getRange("C1").getValue()}${thisSheet.getRange("D1").getValue()}`);
  thisSheet.insertRowAfter(3);
  thisSheet.getRange("C1:D3").clear({contentsOnly: true});

  // Align all text vertically in the center
  thisSheet.getRange(`1:${thisSheet.getMaxRows()}`).setVerticalAlignment("center");

  // Set the amounts to be recognised as numbers
  thisSheet.getRange(`C2:D${thisSheet.getMaxRows()}`).setNumberFormat("#,##0.00");

  // Freeze the first five rows
  thisSheet.setFrozenRows(5);

  // Insert a blank row after each cost code's totals (except the grand total)
  finder = thisSheet.createTextFinder("Total Balance - ");
  matches = finder.findAll();
  matches.pop();
  matches.reverse();  // start at the bottom to avoid changing future ranges
  for (let i = 0; i < matches.length; i += 1) {
    thisSheet.insertRowAfter(matches[i].getRow());
  }

  // Finish - this is a flag to indicate that formatting is done
  thisSheet.getRange("A1").setValue("Income/Expense Statement");
  SpreadsheetApp.flush();

}


/**
 * Determines whether or not the row number in the current sheet
 * appears in the oldSheet.
 *
 * @param {Number} row the row number in the new sheet.
 * @param {[][]} oldSheetValues the sheet to search, from getSheetValues().
 * @param {Sheet} newSheet the sheet with the numbered row.
 * @return {Boolean} true if present, otherwise false.
 */
function compareWithOld(row, oldSheetValues, newSheet) {

  // If newSheet has been supplied then use it, otherwise use the default
  let rowValues;
  if (newSheet === undefined) {
    rowValues = sheet.getSheetValues(row, 1, 1, 4)[0];
  } else {
    rowValues = newSheet.getSheetValues(row, 1, 1, 4)[0];
  }
  let newRow = true;
  for (let i = 0; i < oldSheetValues.length; i += 1) {
    if (rowValues.toString() === oldSheetValues[i].toString()) {
      newRow = false;
      break;
    }
  }
  return newRow;
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
 *
 * @param {String} url the URL of the spreadsheet with a sheet named Original.
 */
function compareLedgers(url) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  const originalSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log(`Script has opened spreadsheet ${url}`);
  const originalSheet = originalSpreadsheet.getSheetByName("Original");
  const originalSheetValues = originalSheet.getSheetValues(1, 1, originalSheet.getLastRow(), 4);

  let passedHeader = false;
  let cell;
  let cellValue;
  for (let row = 1; row <= sheet.getLastRow(); row += 1) {
    cell = sheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't passed the first header row then skip it
    if (passedHeader === false) {
      if (cellValue === "Date") {
        passedHeader = true;
        sheet.getRange(row + 1, 1, sheet.getLastRow() - row - 1).setBackground("green");
      }


      // Compare it with the original/old sheet
      // Comparing all rows allows us to identify changes in the totals too
    } else {
      const isNew = compareWithOld(row, originalSheetValues, sheet);

      // If it is a new row then colour it
      if (isNew) {
        sheet.getRange(row, 1, 1, 4).setBackground("red");
        Logger.log(`Row ${row} is a new row!`);

        // Otherwise just reset it
      } else {
        sheet.getRange(row, 1).setBackground("white");
      }
    }
  }
  Logger.log("Finished comparing sheets!");
}


/**
 * Copies the sheet to another spreadsheet file.
 * This is normally the one with the ledger in it.
 * The script then deletes the old Original, and renames the new
 * Original and sets protections on it.
 *
 * @param url the URL of the spreadsheet with the sheet to replace.
 */
function copyToLedger(url) {

  // Gets the ledger spreadsheet
  const ledgerSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log(`Script has opened spreadsheet ${url}`);

  // Copies to the ledger spreadsheet
  const newSheet = sheet.copyTo(ledgerSpreadsheet);

  // Remove protections from the old Original sheet and delete it
  const oldOriginalSheet = ledgerSpreadsheet.getSheetByName("Original");
  oldOriginalSheet.protect().remove();
  ledgerSpreadsheet.deleteSheet(oldOriginalSheet);

  // Rename the new Original sheet and protect it
  newSheet.setName("Original");
  newSheet.protect().setWarningOnly(true);

  Logger.log("Finished copying the sheet to the ledger spreadsheet.");

  // Ask the user if they want to open the new sheet
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    "Do you want to open the ledger?",
    "This will open the ledger in a new tab.",
    ui.ButtonSet.YES_NO);

  // Open it if the user want to
  if (result === ui.Button.YES) {
    const sheetUrl = `${url}#gid=${newSheet.getSheetId()}`;
    openUrl(sheetUrl);
  }
}


/**
 * Prompt the user for a URL and then run compareLedgers(url).
 */
function compareLedgersGetUrl() {
  const url = showPrompt();
  if (url === false) {
    return;
  }
  compareLedgers(url);
}


/**
 * Runs formatNeatlyWithSheet() with the active sheet.
 */
function formatNeatly() {
  formatNeatlyWithSheet(sheet);
}


/**
 * Prompt the user for a URL and then run copyToLedger(url).
 */
function copyToLedgerGetUrl() {
  const url = showPrompt();
  if (url === false) {
    return;
  }
  copyToLedger(url);
}


/**
 * Does everything at once: the formatting and the comparison.
 * This uses the URL at the named range called DefaultUrl
 * The user is alerted when it is completed (so that
 * they know it hasn't stalled).
 */
function processWithDefaultUrl() {

  formatNeatly();
  const url = getNamedRange("DefaultUrl",
    SpreadsheetApp.getActiveSpreadsheet()).getValue()

  compareLedgers(url);
  copyToLedger(url);
  const ui = SpreadsheetApp.getUi();
  ui.alert("Complete!",
    "The process completed successfully. You can view the logs here: " +
    "https://script.google.com/home/projects/1ycYC3ziL1kmxVl-RlqJhO-4UXZXjHw0Nljnal1IIfkWNDG-MPooyufqr/executions",
    ui.ButtonSet.OK);
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
