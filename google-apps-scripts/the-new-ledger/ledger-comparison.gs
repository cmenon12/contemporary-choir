/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2020 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


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
 * Formats the ledger neatly.
 *
 * @param {Sheet} thisSheet the sheet to format.
 * @param {String} sheetName the name of the sheet, with {{datetime}}
 * as the date and time.
 *
 */
function formatNeatly(thisSheet, sheetName) {

  // Don't re-run if it's already been done
  if (thisSheet.getRange("A1").getValue() === "Income/Expense Statement") {
    return;
  }

  // Convert the first column to text
  thisSheet.getRange(1, 1, thisSheet.getLastRow()).setNumberFormat("@");

  // Change the tab colour to be white
  thisSheet.setTabColor("white");

  // Rename the sheet
  if (sheetName !== "") {
    const datetime = `${thisSheet.getRange("C1").getValue()}${thisSheet.getRange("D1").getValue()}`;
    thisSheet.setName(sheetName.replace(/{{datetime}}/g, datetime));
  }

  // Delete the 'Please note' text at the end
  let finder = thisSheet.getRange("A:A").createTextFinder("Please note");
  const foundRange = finder.findNext();
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
  thisSheet.getRange(`${thisSheet.getLastRow() - 3}:${thisSheet.getLastRow()}`).setFontSize(12);
  thisSheet.getRange("1:4").setFontSize(12);

  // Replace all in-cell newlines with spaces
  finder = thisSheet.createTextFinder("\n")
    .useRegularExpression(true)
    .replaceAllWith(" ");

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

  // Get the values in this row
  const rowValues = newSheet.getSheetValues(row, 1, 1, 4)[0];
  let newRow = true;

  // Check for this row in the old sheet's values
  for (let i = 0; i < oldSheetValues.length; i += 1) {
    if (rowValues.toString() === oldSheetValues[i].toString()) {
      newRow = false;
      break;
    }
  }
  return newRow;
}


/**
 * Compares the new sheet with the old sheet, and highlights anything
 * different or new in the new sheet.
 * Note that any differences in whitespace will be recognised as a difference.
 * newLedger must have the cost codes saved to it.
 *
 * @param {Sheet} newSheet the new sheet with the new changes.
 * @param {Sheet} oldSheet the old sheet to compare against.
 * @param {boolean} colourCountdown whether to show the comparison working downwards.
 * @param {String} newRowColour the colour of the new row.
 * @param {Ledger|null} newLedger the ledger object to save the entries to, optional.
 * @returns {Ledger|null} the updated ledger object if one was supplied.
 */
function compareLedgers(newSheet, oldSheet, colourCountdown, newRowColour, newLedger = null) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  const oldSheetValues = oldSheet.getSheetValues(1, 1, oldSheet.getLastRow(), 4);

  let passedHeader = false;
  let cell;
  let cellValue;
  let costCodeRows;
  if (newLedger !== null) {
    costCodeRows = newLedger.getCostCodeRows();
  }

  for (let row = 1; row <= newSheet.getLastRow(); row += 1) {
    cell = newSheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't passed the first header row then skip it
    if (passedHeader === false) {

      // Header row is indicated by Date
      if (cellValue === "Date") {
        passedHeader = true;

        // Colour all the rows after this
        if (colourCountdown === true) {

          // Don't use the same colour as to highlight
          let bgColour = "green";
          if (newRowColour === "green" || newRowColour === "#008000") {
            bgColour = "lightblue";
          }
          newSheet.getRange(row + 1, 1, newSheet.getLastRow() - row - 1).setBackground("green");
        }
      }


    } else {

      // Compare it with the original/old sheet
      const isNew = compareWithOld(row, oldSheetValues, newSheet);

      // If it is a new row then colour it
      if (isNew) {
        Logger.log(`Row ${row} is a new row!`);
        newSheet.getRange(row, 1, 1, 4).setBackground(newRowColour);

        // If we are using the ledger object and it's not a date
        if (newLedger !== null && isADate(newSheet.getRange(row, 1).getValue())) {

          // Identify the relevant cost code and save the entry
          for (let i = 0; i < costCodeRows.length; i++) {
            if (row < costCodeRows[i][1]) {
              newLedger.addEntry(costCodeRows[i][0],
                newSheet.getRange(row, 1).getValue(),
                newSheet.getRange(row, 2).getValue(),
                Number(newSheet.getRange(row, 3).getValue()),
                Number(newSheet.getRange(row, 4).getValue()));
              break;
            }
          }
        }

        // Otherwise just reset it
      } else if (colourCountdown === true) {
        newSheet.getRange(row, 1).setBackground("white");
      }
    }
  }
  Logger.log("Finished comparing sheets!");
  if (newLedger !== null) {
    newLedger.log();
    return newLedger;
  }
}


/**
 * Copies the sheet to another spreadsheet file.
 * Either overwrites an existing sheet (but retains protections), or copies it
 * and sets a custom name.
 *
 * @param {Sheet} thisSheet the sheet to copy.
 * @param {Spreadsheet} destSpreadsheet the destination spreadsheet to copy to.
 * @param {Sheet} destSpreadsheet the destination sheet to overwrite, or null.
 * @param {String|Sheet} newSheet the name of the new sheet or the sheet to
 * overwrite. If the former then with {{datetime}} as the date and time.
 */
function copyToLedger(thisSheet, destSpreadsheet, newSheet) {

  // Test if newSheet is a sheet or a string
  let isNewSheetASheet;
  try {
    newSheet.getName();
    isNewSheetASheet = true;
  } catch (err) {
    isNewSheetASheet = false;
  }

  // If they gave a sheet then overwrite it
  if (isNewSheetASheet) {

    // Get the name and protections
    const newSheetProtections = newSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    const newSheetName = newSheet.getName();

    // Copy it, and set the same protections
    const copiedSheet = thisSheet.copyTo(destSpreadsheet);
    newSheetProtections.setRange(copiedSheet.getRange(`1:${copiedSheet.getMaxRows()}`));

    // Unprotect and delete the old sheet
    newSheet.protect().remove();
    destSpreadsheet.deleteSheet(newSheet);

    // Set the name
    copiedSheet.setName(newSheetName);


  } else {

    // Just copy it over and set the name
    const copiedSheet = thisSheet.copyTo(destSpreadsheet);
    const datetime = `${thisSheet.getRange("C1").getValue()}${thisSheet.getRange("D1").getValue()}`;
    copiedSheet.setName(newSheet.replace(/{{datetime}}/g, datetime));
  }
}

