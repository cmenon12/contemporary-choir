/**
 * Get the event categories in an array.
 *
 * @returns {Array.<string>} the updated ledger object if one was supplied.
 */
function getCategories() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let range = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2);
  return range.getValues()
}


/**
 * Get the events on a date, returning them in an object.
 * If there are any duplicates for each category, then only
 * the first will be kept (and the rest deleted).
 * Events with invalid categories are deleted.
 */
function getEventsOnDate(dateObj) {

  let categories = getCategories()
  let calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  let events = calendar.getEventsForDay(dateObj)
  let eventsObj = {}

  for (let i = 0; i < events.length; i++) {
    let category = events[i].getTag("category")
    if (eventsObj[category] === undefined && categories.includes(category)) {
      eventsObj[category] = events[i];
    } else {
      events[i].deleteEvent();
    }
  }

  return eventsObj;
}


function checkSheet() {

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  // Iterate over each date (row)
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    let dateRange = sheet.getRange(i, 2, 1, sheet.getLastColumn() - 1);
    let eventsObj = getEventsOnDate(dateRange.getValue());

    // iterate over each category
    // add/remove/update as needed


  }
}