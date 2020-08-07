/**
 * This function checks if there are any new entries in the sheet,
 * and if so returns a list of them all, tagged with their cost code.
 */
function checkForNewTotals(sheetName) {

  // Get the spreadsheet & sheet
  var spreadsheet = SpreadsheetApp.getActive();
  var newSheet = spreadsheet.getSheetByName(sheetName)
  Logger.log("Looking at the sheet called " + newSheet.getName())

  // Find the income & expenditure for each cost code
  var newCostCodeTotals = getCostCodeTotals(newSheet)

  // Format the sheet neatly, and rename it to reflect that this is automated
  formatNeatlyWithSheet(newSheet)
  newSheet.setName(newSheet.getName() + " auto")

  // Locate the named range with the URL to the old sheet
  var namedRanges = spreadsheet.getNamedRanges();
  var url;
  for (var i=0; i<namedRanges.length; i++) {
    if (namedRanges[i].getName() == "DefaultUrl") {
      url = namedRanges[i].getRange().getValue();
    }
  }

  // Find the total income and expense in the Original (the one that we are comparing against)
  var oldSpreadsheet = SpreadsheetApp.openByUrl(url);
  Logger.log("Script has opened spreadsheet " + url);
  var oldSheet = oldSpreadsheet.getSheetByName("Original");
  var oldCostCodeTotals = getCostCodeTotals(oldSheet);

  // If there's no difference then stop and delete the sheet
  if (newCostCodeTotals[newCostCodeTotals.length - 1][1] == oldCostCodeTotals[oldCostCodeTotals.length - 1][1] &&
      newCostCodeTotals[newCostCodeTotals.length - 1][2] == oldCostCodeTotals[oldCostCodeTotals.length - 1][2]) {
    Logger.log("There is no difference in the total income or expenditure.")
    spreadsheet.deleteSheet(newSheet)
    return "False";
  }

  // If there is a difference then make comparisons and return the changes
  Logger.log("There is a difference in the total income and/or expenditure!")
  var changes = compareLedgersWithCostCodes(newSheet, oldSheet, newCostCodeTotals)
  changes.unshift(newSheet.getSheetId())
  changes.push(newCostCodeTotals)
  Logger.log(changes)
  return changes;
}


/**
 * This function is used to search for changes in the newSheet compared
 * with the oldSheet (not vice-versa). It will categorise them by cost
 * code and return them.
 */
function compareLedgersWithCostCodes(newSheet, oldSheet, costCodes) {

  // Fetch the old sheet and it's values
  // This saves making multiple requests, which is slow
  var oldSheetValues = oldSheet.getSheetValues(1, 1, oldSheet.getLastRow(), 4);

  var passedHeader = false;
  var cell;
  var cellValue;
  var changes = [];
  for (var row = 1; row<=newSheet.getLastRow(); row+=1) {
    cell = newSheet.getRange(row, 1);
    cellValue = String(cell.getValue());

    // If we still haven't pased the first header row then skip it
    if (passedHeader == false) {
      if (cellValue == "Date") {
        passedHeader = true;
      }


      // Compare it with the original/old sheet
    // Comparing all rows allows us to identify changes in the totals too
    } else {
      var isNew = compareWithOld(row, oldSheetValues, newSheet);

      // If it is a new row and has a date then save it with its cost code
      if (isNew && isADate(newSheet.getRange(row, 1).getValue())) {
        Logger.log("Row " + row + " is a new row!");
        newSheet.getRange(row, 1, 1, 4).setBackground("#FFA500");

        // Identify the relevant cost code and save it
        for (var i = 0; i < costCodes.length; i++) {
          if (row < costCodes[i][4]) {
            changes.push([costCodes[i][0],
              newSheet.getRange(row, 1).getValue(),
              newSheet.getRange(row, 2).getValue(),
              newSheet.getRange(row, 3).getValue(),
              newSheet.getRange(row, 4).getValue()])
            break;
          }
        }
      }
    }
  }
  Logger.log("Finished comparing sheets!")
  return changes;
}

/**
 * This function retrieves the total income, expenditure, and balance for
 * each cost code, as well as the grand total for the entire ledger.
 */
function getCostCodeTotals(sheet) {

  var costCodeTotals = [];
  var costCode;

  // Search for the total for each cost code (and the grand total)
  var finder = sheet.createTextFinder("Totals for ").matchEntireCell(false)
  var foundRanges = finder.findAll()
  for (var i=0; i< foundRanges.length; i++) {

    // Get the name of the cost code
    costCode = String(foundRanges[i].getValue()).replace("Totals for ", "")

    // Append the name, total income, total expenditure, balance, and row number
    costCodeTotals.push([costCode, foundRanges[i].offset(0, 1).getValue(),
      foundRanges[i].offset(0, 2).getValue(),
      foundRanges[i].offset(1, 2).getValue(),
      foundRanges[i].getRow()])
  }
  Logger.log(costCodeTotals)
  return costCodeTotals

}
