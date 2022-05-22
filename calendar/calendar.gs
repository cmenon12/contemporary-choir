/**
 * Get the event categories in an array.
 *
 * @returns {Array.<string>} an array of the categories
 */
function getCategories() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let range = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2);
  return range.getValues()[0]
}


/**
 * Get the events on a date, returning them in an object.
 * If there are any duplicates for each category, then only
 * the first will be kept (and the rest deleted).
 * Events with invalid categories are deleted.
 */
function getEventsOnDate(dateObj) {

  const categories = getCategories()
  let calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  let events = calendar.getEventsForDay(dateObj)
  let eventsObj = {}

  // For each event on this date
  for (let i = 0; i < events.length; i++) {
    let category = events[i].getTag("category")

    // If the category is valid and we don't have one for that category
    if (categories.includes(category) && eventsObj[category] === undefined) {
      eventsObj[category] = events[i];
    } else {
      events[i].deleteEvent();
    }
  }

  return eventsObj;
}


/**
 * @param {number} rowNum
 * @param {number} colNum
 * @param {string} category
 */
function createEventFromCell(rowNum, colNum, category) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID);

  let date = sheet.getRange(rowNum, 2).getValue();
  let eventRange = sheet.getRange(rowNum, colNum);

  // Create the event and set the category
  let event = calendar.createAllDayEvent(eventRange.getValue(), date)
  event.setTag("category", category)

  return event

}


/**
 * @param {CalendarApp.CalendarEvent} calEvent
 * @param {number} sheetEventRow
 * @param {number} sheetEventCol
 */
function compareEvents(calEvent, sheetEventRow, sheetEventCol) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let sheetEventRange = sheet.getRange(sheetEventRow, sheetEventCol);

  // False if the titles are different
  if (calEvent.getTitle() !== sheetEventRange.getValue()) return false;

  return true;
}


function checkSheet() {

  const categories = getCategories()
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  // Iterate over each date (row)
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    let dateRange = sheet.getRange(i, 2, 1, sheet.getLastColumn() - 1);
    let sheetEvents = dateRange.getValues()[0]
    let calEvents = getEventsOnDate(dateRange.getValue());

    // Iterate over each category
    for (let j = 0; j < categories.length; j++) {
      let currentCategory = categories[j]

      // If there is no calendar event but there is a sheets event
      if (calEvents[currentCategory] === undefined && sheetEvents[j + 1] !== "") {
        // Create a new event from the cell
        let event = createEventFromCell(i, j + 3, currentCategory)
        console.log(`Created ${currentCategory} event on ${event.getAllDayStartDate()}.`)

        // If there is a calendar event and a sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] !== "") {
        // Compare them; delete and replace if different
        if (compareEvents(calEvents[currentCategory], i, j + 3) === false) {
          calEvents[currentCategory].deleteEvent()
          let event = createEventFromCell(i, j + 3, currentCategory)
          console.log(`Replaced ${currentCategory} event on ${event.getAllDayStartDate()}.`)
        } else {
          console.log(`Unchanged ${currentCategory} event on ${calEvents[currentCategory].getAllDayStartDate()}.`)
        }

        // If there is a calendar event but no sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] === "") {
        // Delete the event
        console.log(`Deleted ${currentCategory} event on ${calEvents[currentCategory].getAllDayStartDate()}.`)
        calEvents[currentCategory].deleteEvent()
      }
    }


  }
}