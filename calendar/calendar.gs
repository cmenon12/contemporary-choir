/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2022 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Get the event categories and colours in an object.
 * The keys are the category names; the values are the colour numbers.
 * https://developers.google.com/apps-script/reference/calendar/event-color
 *
 * @returns {Object} an object of the categories
 */
function getCategories() {
  let categories = {}
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let range = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2);
  for (let i = 0; i < range.getValues()[0].length; i++) {
    categories[range.getValues()[0][i]] = String(sheet.getRange(2, i + 3).getValue());
  }
  return categories
}


/**
 * Get the events on a date, returning them in an object.
 * The keys are the category names; the values are the events
 *
 * If there are any duplicates for each category, then only
 * the first will be kept (and the rest deleted).
 * Events with invalid categories are deleted.
 *
 * @param {Array.<CalendarApp.CalendarEvent>} allEvents all the fetched events
 * @param {Date} dateObj the date to get events on
 * @returns {Object} an object of the events by category
 */
function getEventsOnDate(allEvents, dateObj) {

  const categories_names = Object.keys(getCategories())
  let eventsObj = {}

  // For each event
  for (let i = 0; i < allEvents.length; i++) {

    // If it's an all day event on the right day
    if (allEvents[i].isAllDayEvent() && allEvents[i].getAllDayStartDate().getTime() === dateObj.getTime()) {

      let category = allEvents[i].getTag("category")

      // If the category is valid and we don't have one for that category
      if (categories_names.includes(category) && eventsObj[category] === undefined) {
        eventsObj[category] = allEvents[i];
      } else {
        allEvents[i].deleteEvent();
      }

    // Delete it if it's not an all day event
    } else if (!allEvents[i].isAllDayEvent()) {
      allEvents[i].deleteEvent();
    }
  }

  return eventsObj;
}


/**
 * Generate the event description for the given cell.
 *
 * @param {SpreadsheetApp.Range} eventRange the cell to use
 * @returns {string} the generated description
 */
function generateDescription(eventRange) {
  let description = `View the event in the ${SpreadsheetApp.getActiveSpreadsheet().getName()} spreadsheet at the link above (in cell ${eventRange.getA1Notation()}).`

  // Add any URLs
  let urls = []
  const rtfRuns = eventRange.getRichTextValue().getRuns()
  for (let i = 0; i < rtfRuns.length; i++) {
    if (rtfRuns[i].getLinkUrl() !== null) {
      urls.push(rtfRuns[i].getLinkUrl())
    }
  }
  if (urls.length == 1) {
    description = description + `\n\nLINK: ${urls[0]}`
  } else if (urls.length > 1) {
    description = description + `\n\nLINKS:`
    for (let i = 0; i < urls.length; i++) {
      description = description + `\n${urls[i]}`
    }
  }

  // Add the note if it's present
  if (eventRange.getNote() != "") {
    description = description + `\n\nNOTE: ${eventRange.getNote()}`
  }

  return description

}


/**
 *
 * Create an event from the cell in Google Calendar.
 *
 * @param {number} rowNum the row number of the cell
 * @param {number} colNum the column number of the cell
 * @param {string} category the category of the event
 * @param {string} colourStr the colour number of the event (disabled)
 * @returns {CalendarApp.CalendarEvent} the created event
 */
function createEventFromCell(rowNum, colNum, category, colourStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID);

  let date = sheet.getRange(rowNum, 2).getValue();
  let eventRange = sheet.getRange(rowNum, colNum);

  // Create the event
  let event = calendar.createAllDayEvent(`${eventRange.getValue()} [${category.toLowerCase()}]`, date)
  event.setTag("category", category)
  event.setLocation(getSecrets().CELL_LINK_PREFIX + eventRange.getA1Notation().replace(/:/g, ''))
  event.setDescription(generateDescription(eventRange))
  // event.setColor(colourStr)

  return event

}


/**
 * Compare the sheet and calendar event.
 *
 * This compares by title and description.
 *
 * @param {CalendarApp.CalendarEvent} calEvent the calendar event
 * @param {number} sheetEventRow the row number of the cell
 * @param {number} sheetEventCol the column number of the cell
 * @returns {boolean} true if equal otherwise false
 */
function compareEvents(calEvent, sheetEventRow, sheetEventCol, category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let sheetEventRange = sheet.getRange(sheetEventRow, sheetEventCol);

  // False if the titles are different
  if (calEvent.getTitle() !== `${sheetEventRange.getValue()} [${category.toLowerCase()}]`) return false;

  // False if the descriptions are different
  if (calEvent.getDescription() !== generateDescription(sheetEventRange)) return false;

  return true;
}


/**
 * Main function, updates Google Calendar with any changes.
 */
function checkSheet(startRow = 3) {

  const categories = getCategories()
  const categories_names = Object.keys(getCategories())
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  // Prefetch data
  const allDatesRange = sheet.getRange(startRow, 2, sheet.getLastRow()-startRow+1, sheet.getLastColumn() - 1).getValues()
  const startDate = allDatesRange[0][0]
  let endDate = allDatesRange[allDatesRange.length - 1][0]
  endDate.setDate(endDate.getDate()+1)
  console.log(`Fetching events from ${startDate} to ${endDate}.`)
  const calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  const allCalEvents = calendar.getEvents(startDate, endDate)

  // Iterate over each row
  for (let i = 0; i < allDatesRange.length; i++) {
    let sheetEvents = allDatesRange[i]
    let calEvents = getEventsOnDate(allCalEvents, allDatesRange[i][0]);

    // Iterate over each category
    for (let j = 0; j < categories_names.length; j++) {
      let currentCategory = categories_names[j]

      // If there is no calendar event but there is a sheets event
      if (calEvents[currentCategory] === undefined && sheetEvents[j + 1] !== "") {
        // Create a new event from the cell
        let event = createEventFromCell(i+startRow, j + 3, currentCategory, categories[currentCategory])
        console.log(`Created "${event.getTitle()}" event on ${event.getAllDayStartDate().toLocaleDateString("en-GB")}.`)

        // If there is a calendar event and a sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] !== "") {
        // Compare them; delete and replace if different
        if (compareEvents(calEvents[currentCategory], i+startRow, j + 3, currentCategory) === false) {
          calEvents[currentCategory].deleteEvent()
          let event = createEventFromCell(i+startRow, j + 3, currentCategory, categories[currentCategory])
          console.log(`Replaced "${event.getTitle()}" event on ${event.getAllDayStartDate().toLocaleDateString("en-GB")}.`)
        } else {
          console.log(`Unchanged "${calEvents[currentCategory].getTitle()}" event on ${calEvents[currentCategory].getAllDayStartDate().toLocaleDateString("en-GB")}.`)
        }

        // If there is a calendar event but no sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] === "") {
        // Delete the event
        console.log(`Deleted "${calEvents[currentCategory].getTitle()}" event on ${calEvents[currentCategory].getAllDayStartDate().toLocaleDateString("en-GB")}.`)
        calEvents[currentCategory].deleteEvent()
      }
    }


  }
}


/**
 * Hides rows in the past up to (and excluding) the most recent Monday.
 */
function hidePastRows() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  // Hide rows in the past
  sheet.showRows(1, sheet.getLastRow());
  const today = new Date();
  const first = today.getDate() - today.getDay() + 1;
  const monday = new Date(today.setDate(first));
  for (let i = 3; i < sheet.getLastRow() - 1; i++) {
    if (sheet.getRange(i, 2).getValue().getTime() >= monday.getTime()) {
      sheet.hideRows(2, i - 3)
      break;
    }
  }

}


/**
 * Run checkSheet() with all rows.
 */
function checkSheetAll() {
  checkSheet()
}


/**
 * Run checkSheet() with the rows from today onwards only.
 * We include the 6 before as these might not be hidden.
 */
function checkSheetNoPast() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  const now = new Date();
  for (let i = 3; i < sheet.getLastRow() - 1; i++) {
    if (sheet.getRange(i, 2).getValue().getTime() >= now.getTime()) {
      if (i-7 <= 3) {
        checkSheet(3)
      } else {
        checkSheet(i-7)
      }
      break;
    }
  }
}


/**
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Update Google Calendar (all)", "checkSheetAll");
  menu.addItem("Update Google Calendar (today onwards only)", "checkSheetNoPast");
  menu.addSeparator()
  menu.addItem("Hide past rows", "hidePastRows")
  menu.addToUi();
}