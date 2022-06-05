/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2022 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


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
 * @param {Array.<String>} categories the categories
 * @returns {Object} an object of the events by category
 */
function getEventsOnDate(allEvents, dateObj, categories) {

  let eventsObj = {}

  // For each event
  for (let i = 0; i < allEvents.length; i++) {

    // If it's an all day event on the right day
    if (allEvents[i].isAllDayEvent() && allEvents[i].getAllDayStartDate().getTime() === dateObj.getTime()) {

      const category = allEvents[i].getTag("category")

      // If the category is valid and we don't have one for that category
      if (categories.includes(category) && eventsObj[category] === undefined) {
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
 * @param {String} cellNotation the A1 notation of the cell
 * @param {SpreadsheetApp.RichTextValue} cellRtf the RTF of the cell
 * @param {String} cellNote the cell note
 * @returns {String} the generated description
 */
function generateDescription(cellNotation, cellRtf, cellNote) {
  let description = `View the event in the ${SpreadsheetApp.getActiveSpreadsheet().getName()} spreadsheet at the link above (in cell ${cellNotation}).`

  // Add any URLs
  let urls = []
  const rtfRuns = cellRtf.getRuns()
  for (let i = 0; i < rtfRuns.length; i++) {
    if (rtfRuns[i].getLinkUrl() !== null) {
      urls.push(rtfRuns[i].getLinkUrl())
    }
  }
  if (urls.length === 1) {
    description = description + `\n\nLINK: ${urls[0]}`
  } else if (urls.length > 1) {
    description = description + `\n\nLINKS:`
    for (let i = 0; i < urls.length; i++) {
      description = description + `\n${urls[i]}`
    }
  }

  // Add the note if it's present
  if (cellNote !== "") {
    description = description + `\n\nNOTE: ${cellNote}`
  }

  return description

}


/**
 *
 * Create an event from the cell in Google Calendar.
 *
 * @param {Date} eventDate the row number of the cell
 * @param {String} cellValue the value of the cell
 * @param {String} cellNotation the A1 notation of the cell
 * @param {SpreadsheetApp.RichTextValue} cellRtf the RTF of the cell
 * @param {String} cellNote the cell note
 * @param {String} category the category of the event
 * @returns {CalendarApp.CalendarEvent} the created event
 */
function createEventFromCell(eventDate, cellValue, cellNotation, cellRtf, cellNote, category) {
  const calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID);

  // Create the event
  const event = calendar.createAllDayEvent(`${cellValue} [${category.toLowerCase()}]`, eventDate)
  event.setTag("category", category)
  event.setLocation(getSecrets().CELL_LINK_PREFIX + cellNotation.replace(/:/g, ''))
  event.setDescription(generateDescription(cellNotation, cellRtf, cellNote))

  return event

}


/**
 * Compare the sheet and calendar event.
 *
 * This compares by title and description.
 *
 * @param {CalendarApp.CalendarEvent} calEvent the calendar event
 * @param {String} cellValue the value of the cell
 * @param {String} cellNotation the A1 notation of the cell
 * @param {SpreadsheetApp.RichTextValue} cellRtf the RTF of the cell
 * @param {String} cellNote the cell note
 * @param {String} category the category of the event
 * @returns {boolean} true if equal otherwise false
 */
function compareEvents(calEvent, cellValue, cellNotation, cellRtf, cellNote, category) {
  // False if the titles are different
  if (calEvent.getTitle() !== `${cellValue} [${category.toLowerCase()}]`) return false;

  // False if the descriptions are different, otherwise true
  return calEvent.getDescription() === generateDescription(cellNotation, cellRtf, cellNote);


}


/**
 * Get the A1 notation (e.g. B24) of the cell.
 *
 * @param rowNum {number} the row number, from 1
 * @param colNum {number} the column number, from 1
 * @returns {String} the A1 notation of the cell
 */
function getA1Notation(rowNum, colNum){
  return `${String.fromCharCode("A".charCodeAt(0)+colNum-1)}${rowNum}`
}


/**
 * Main function, updates Google Calendar with any changes.
 */
function checkSheet(startRow = 2) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  const categories = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2).getValues()[0]

  // Prefetch data
  const allDatesRange = sheet.getRange(startRow, 2, sheet.getLastRow()-startRow+1, sheet.getLastColumn() - 1)
  const allDatesValues = allDatesRange.getValues()
  const allDatesRtfs = allDatesRange.getRichTextValues()
  const allDatesNotes = allDatesRange.getNotes()
  const startDate = allDatesValues[0][0]
  let endDate = allDatesValues[allDatesValues.length - 1][0]
  endDate.setDate(endDate.getDate()+1)
  console.log(`Fetching events from ${startDate} to ${endDate}.`)
  const calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  const allCalEvents = calendar.getEvents(startDate, endDate)

  // Iterate over each row
  for (let i = 0; i < allDatesValues.length; i++) {
    const sheetEvents = allDatesValues[i]
    const calEvents = getEventsOnDate(allCalEvents, allDatesValues[i][0], categories);

    // Iterate over each category
    for (let j = 0; j < categories.length; j++) {
      const currentCategory = categories[j]

      // If there is no calendar event but there is a sheets event
      if (calEvents[currentCategory] === undefined && sheetEvents[j + 1] !== "") {
        // Create a new event from the cell
        const event = createEventFromCell(allDatesValues[i][0], allDatesValues[i][j+1],getA1Notation(i+startRow, j+3), allDatesRtfs[i][j+1], allDatesNotes[i][j+1], currentCategory)
        console.log(`Created "${event.getTitle()}" event on ${event.getAllDayStartDate().toLocaleDateString("en-GB")}.`)

        // If there is a calendar event and a sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] !== "") {
        // Compare them; delete and replace if different
        if (compareEvents(calEvents[currentCategory], allDatesValues[i][j+1],getA1Notation(i+startRow, j+3), allDatesRtfs[i][j+1], allDatesNotes[i][j+1], currentCategory) === false) {
          calEvents[currentCategory].deleteEvent()
          const event = createEventFromCell(allDatesValues[i][0], allDatesValues[i][j+1],getA1Notation(i+startRow, j+3), allDatesRtfs[i][j+1], allDatesNotes[i][j+1], currentCategory)
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