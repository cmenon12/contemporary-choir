/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2022 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Add days to a date.
 *
 * Credit to https://stackoverflow.com/a/563442/
 *
 * @param {Number} days the number of days to add
 * @returns {Date} the modified date
 */
Date.prototype.addDays = function (days) {
  let date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
}


/**
 * Get an object of the multi-day events.
 *
 * @param {SpreadsheetApp.Range} allDatesRange the calendar range
 * @param {Array.<Array<any>>} allDatesValues the values of the calendar range
 * @param {Array.<String>} categories the categories
 * @param {number} startRow the row number to start on
 * @returns {Object} an object of the multi-day events
 */
function getMergedEvents(allDatesRange, allDatesValues, categories, startRow) {

  // Prep the result variable
  let result = {}
  for (c of categories) {
    result[c] = {}
  }

  // Get the merged ranges
  const mergedRanges = allDatesRange.getMergedRanges()

  // Iterate over each one
  for (let i = 0; i < mergedRanges.length; i++) {
    const currentCategory = categories[mergedRanges[i].getColumn() - 3]
    const firstRowNum = mergedRanges[i].getRow()
    const lastRowNum = mergedRanges[i].getLastRow()
    if (firstRowNum >= startRow) {
      const startDate = allDatesValues[firstRowNum - startRow][0]
      const endDate = allDatesValues[lastRowNum - startRow][0].addDays(1)
      result[currentCategory][startDate] = endDate;
    } else {
      console.log(`Skipped the multi-day event starting on row ${firstRowNum}`)
    }
  }

  console.log(`Multi-day events are ${JSON.stringify(result)}.`)
  return result

}


/**
 * Get the events on a date, returning them in an object.
 * The keys are the category names; the values are the arrays of events
 *
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

      // If the category is valid, and we don't have one for that category
      if (categories.includes(category) && eventsObj[category] === undefined) {
        eventsObj[category] = [allEvents[i]];
      } else if (categories.includes(category)) {
        eventsObj[category].push(allEvents[i])
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
    description = description + `\n\n<b>Link: </b>${urls[0]}`
  } else if (urls.length > 1) {
    description = description + `\n\n<b>Links</b><ul>`
    for (let i = 0; i < urls.length; i++) {
      description = description + `<li>${urls[i]}</li>`
    }
    description = description + `</li>`
  }

  // Add the note if it's present
  if (cellNote !== "") {
    description = description + `\n\n<b>Note: </b>${cellNote}`
  }

  return description

}


/**
 *
 * Create an event from the cell in Google Calendar.
 *
 * This uses the Calendar Advanced Service, so that it can
 * appear as 'free' (using the transparency).
 *
 * @param {Date} eventDate the start date
 * @param {Date} endDate the end date
 * @param {String} cellValue the value of the cell
 * @param {String} cellNotation the A1 notation of the cell
 * @param {SpreadsheetApp.RichTextValue} cellRtf the RTF of the cell
 * @param {String} cellNote the cell note
 * @param {String} category the category of the event
 * @returns {Calendar.Events} the created event
 */
function createEventFromCell(eventDate, endDate, cellValue, cellNotation, cellRtf, cellNote, category) {

  // Define the new event
  const eventData = {
    start: {
      date: eventDate.toLocaleDateString('en-CA')
    },
    end: {
      date: endDate.toLocaleDateString('en-CA')
    },
    description: generateDescription(cellNotation, cellRtf, cellNote),
    extendedProperties: {
      shared: {
        "category": category
      }
    },
    location: getSecrets().CELL_LINK_PREFIX + cellNotation.replace(/:/g, ''),
    source: {
      title: getSecrets().SOURCE_TITLE,
      url: getSecrets().CELL_LINK_PREFIX + cellNotation.replace(/:/g, '')
    },
    summary: `${cellValue} [${category.toLowerCase()}]`,
    transparency: "transparent",
  }

  // Create the event
  const calEvent = Calendar.Events.insert(eventData, getSecrets().CALENDAR_ID);

  return calEvent

}


/**
 * Compare the sheet and calendar event.
 *
 * This compares by title, end date, and description.
 *
 * @param {CalendarApp.CalendarEvent} calEvent the calendar event
 * @param {String} cellValue the value of the cell
 * @param {String} cellNotation the A1 notation of the cell
 * @param {SpreadsheetApp.RichTextValue} cellRtf the RTF of the cell
 * @param {String} cellNote the cell note
 * @param {Date} cellEndDate the end date based on cell merges
 * @param {String} category the category of the event
 * @returns {boolean} true if equal otherwise false
 */
function compareEvents(calEvent, cellValue, cellNotation, cellRtf, cellNote, cellEndDate, category) {
  // False if the titles are different
  if (calEvent.getTitle() !== `${cellValue} [${category.toLowerCase()}]`) return false;

  // False if the end dates are different (ignore times)
  if (calEvent.getEndTime().toLocaleDateString("en-CA") !== cellEndDate.toLocaleDateString("en-CA")) return false;

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
function getA1Notation(rowNum, colNum) {
  return `${String.fromCharCode("A".charCodeAt(0) + colNum - 1)}${rowNum}`
}


/**
 * Main function, updates Google Calendar with any changes.
 *
 * @param {number} startRow the row number to start on
 */
function checkSheet(startRow = 2) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  const categories = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2).getValues()[0]

  // Prefetch data
  const allDatesRange = sheet.getRange(startRow, 2, sheet.getLastRow() - startRow + 1, sheet.getLastColumn() - 1)
  const allDatesValues = allDatesRange.getValues()
  const allDatesRtfs = allDatesRange.getRichTextValues()
  const allDatesNotes = allDatesRange.getNotes()
  const startDate = allDatesValues[0][0]
  let endDate = allDatesValues[allDatesValues.length - 1][0]
  endDate.setDate(endDate.getDate() + 1)
  console.log(`Fetching events from ${startDate} to ${endDate}.`)
  const calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  const allCalEvents = calendar.getEvents(startDate, endDate)
  const mergedEvents = getMergedEvents(allDatesRange, allDatesValues, categories, startRow)

  // Iterate over each row
  for (let i = 0; i < allDatesValues.length; i++) {
    const sheetEvents = allDatesValues[i]
    const calEvents = getEventsOnDate(allCalEvents, sheetEvents[0], categories);

    // Iterate over each category
    for (let j = 0; j < categories.length; j++) {
      const currentCategory = categories[j]

      // Set the end date based on if it's merged
      let endDate;
      if (mergedEvents[currentCategory][sheetEvents[0]] !== undefined) {
        endDate = mergedEvents[currentCategory][sheetEvents[0]]
      } else {
        endDate = sheetEvents[0].addDays(1)
      }

      // If there is no calendar event but there is one or more sheets events
      if (calEvents[currentCategory] === undefined && sheetEvents[j + 1] !== "") {

        // Split up the cell value
        const sheetEventsSplit = sheetEvents[j + 1].split(getSecrets().MULTI_EVENT_SPLITTER)

        // Create new events from the cell
        for (let s = 0; s < sheetEventsSplit.length; s++) {
          const event = createEventFromCell(sheetEvents[0], endDate, sheetEventsSplit[s].trim(), getA1Notation(i + startRow, j + 3), allDatesRtfs[i][j + 1], allDatesNotes[i][j + 1], currentCategory)
          console.log(`Created "${event.summary}" event on ${event.start.date}.`)
        }

        // If there is at least one calendar and sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] !== "") {

        // Keep track of which calendar events to delete
        let calToDelete = new Array(calEvents[currentCategory].length).fill(true)

        // Iterate over the sheets events
        const sheetEventsSplit = sheetEvents[j + 1].split(getSecrets().MULTI_EVENT_SPLITTER)
        for (let s = 0; s < sheetEventsSplit.length; s++) {

          // Search through the calendar events to try & find one that's the same
          let foundIt = false
          for (let c = 0; c < calEvents[currentCategory].length; c++) {
            if (compareEvents(calEvents[currentCategory][c], sheetEventsSplit[s].trim(), getA1Notation(i + startRow, j + 3), allDatesRtfs[i][j + 1], allDatesNotes[i][j + 1], endDate, currentCategory) === true) {
              console.log(`Unchanged "${calEvents[currentCategory][c].getTitle()}" event on ${calEvents[currentCategory][c].getAllDayStartDate().toLocaleDateString("en-CA")}.`)
              foundIt = true
              calToDelete[c] = false
              break
            }
          }

          // If we didn't find it after iterating then create it
          if (foundIt === false) {
            const event = createEventFromCell(sheetEvents[0], endDate, sheetEventsSplit[s].trim(), getA1Notation(i + startRow, j + 3), allDatesRtfs[i][j + 1], allDatesNotes[i][j + 1], currentCategory)
            console.log(`Created "${event.summary}" event on ${event.start.date}.`)
          }

        }

        // Delete the calendar events we didn't find
        for (let c = 0; c < calEvents[currentCategory].length; c++) {
          if (calToDelete[c] === true) {
            console.log(`Deleted "${calEvents[currentCategory][c].getTitle()}" event on ${calEvents[currentCategory][c].getAllDayStartDate().toLocaleDateString("en-CA")}.`)
            calEvents[currentCategory][c].deleteEvent()
          }
        }

        // If there is one or more calendar events but no sheets events
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] === "") {

        // Delete all the events
        for (let c = 0; c < calEvents[currentCategory].length; c++) {
          console.log(`Deleted "${calEvents[currentCategory][c].getTitle()}" event on ${calEvents[currentCategory][c].getAllDayStartDate().toLocaleDateString("en-CA")}.`)
          calEvents[currentCategory][c].deleteEvent()
        }
      }
    }


  }
}


/**
 * Hides rows in the past up to (and excluding) the most recent Monday.
 */
function hidePastRows() {

  // Get all sheets and a list of which we should process
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const validSheets = getSecrets().VALID_SHEETS;

  for (let i = 0; i < sheets.length; i++) {

    // Only process some sheets
    if (!validSheets.includes(sheets[i].getName())) continue;

    // Hide rows in the past
    const sheet = sheets[i];
    sheet.showRows(1, sheet.getLastRow());
    const today = new Date();
    let daysToRemove;
    today.getDay() === 0 ? (daysToRemove = 7) : (daysToRemove = today.getDay())   // Sunday fix
    const first = today.getDate() - daysToRemove + 1;
    const monday = new Date(today.setDate(first));
    for (let i = 3; i < sheet.getLastRow() - 1; i++) {
      if (sheet.getRange(i, 2).getValue().getTime() >= monday.getTime() && i > 3) {
        sheet.hideRows(2, i - 3)
        break;
      }
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
      if (i - 7 <= 3) {
        checkSheet(3)
      } else {
        checkSheet(i - 7)
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
