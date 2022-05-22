/**
 * Get the event categories in an array.
 *
 * @returns {Object} an object of the categories
 */
function getCategories() {
  let categories = {}
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let range = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 2);
  for (let i = 0; i < range.getValues()[0].length; i++) {
    categories[range.getValues()[0][i]] = String(sheet.getRange(2, i + 3).getValue());
  }
  return categories
}


/**
 * Get the events on a date, returning them in an object.
 * If there are any duplicates for each category, then only
 * the first will be kept (and the rest deleted).
 * Events with invalid categories are deleted.
 */
function getEventsOnDate(dateObj) {

  const categories_names = Object.keys(getCategories())
  let calendar = CalendarApp.getCalendarById(getSecrets().CALENDAR_ID)
  let events = calendar.getEventsForDay(dateObj)
  let eventsObj = {}

  // For each event on this date
  for (let i = 0; i < events.length; i++) {
    let category = events[i].getTag("category")

    // If the category is valid and we don't have one for that category
    if (categories_names.includes(category) && eventsObj[category] === undefined) {
      eventsObj[category] = events[i];
    } else {
      events[i].deleteEvent();
    }
  }

  return eventsObj;
}


/**
 * @param {SpreadsheetApp.Range} eventRange
 */
function generateDescription(eventRange) {
  let description = `View the event in the ${SpreadsheetApp.getActiveSpreadsheet().getName()} spreadsheet at the link above (in cell ${eventRange.getA1Notation()}).`

  // Add any URLs
  let urls = []
  const rtfRuns = eventRange.getRichTextValue().getRuns()
  for (let i = 0; i < rtfRuns.length; i++) {
    if (rtfRuns[i].getLinkUrl() !== null) {
      urls.push(rtfRuns[i].getLinkUrl())
      description = description + `\n\n${rtfRuns[i].getLinkUrl()}`
    }
  }
  if (urls.length >= 1) {
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
 * @param {number} rowNum
 * @param {number} colNum
 * @param {string} category
 * @param {string} colourStr
 */
function createEventFromCell(rowNum, colNum, category, colourStr) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
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
 * @param {CalendarApp.CalendarEvent} calEvent
 * @param {number} sheetEventRow
 * @param {number} sheetEventCol
 */
function compareEvents(calEvent, sheetEventRow, sheetEventCol, category) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  let sheetEventRange = sheet.getRange(sheetEventRow, sheetEventCol);

  // False if the titles are different
  if (calEvent.getTitle() !== `${sheetEventRange.getValue()} [${category.toLowerCase()}]`) return false;

  // False if the descriptions are different
  if (calEvent.getDescription() !== generateDescription(sheetEventRange)) return false;

  return true;
}


function checkSheet() {

  const categories = getCategories()
  const categories_names = Object.keys(getCategories())
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");

  // Iterate over each date (row)
  for (let i = 3; i <= sheet.getLastRow(); i++) {
    let dateRange = sheet.getRange(i, 2, 1, sheet.getLastColumn() - 1);
    let sheetEvents = dateRange.getValues()[0]
    let calEvents = getEventsOnDate(dateRange.getValue());

    // Iterate over each category
    for (let j = 0; j < categories_names.length; j++) {
      let currentCategory = categories_names[j]

      // If there is no calendar event but there is a sheets event
      if (calEvents[currentCategory] === undefined && sheetEvents[j + 1] !== "") {
        // Create a new event from the cell
        let event = createEventFromCell(i, j + 3, currentCategory, categories[currentCategory])
        console.log(`Created ${currentCategory} event on ${event.getAllDayStartDate()}.`)

        // If there is a calendar event and a sheets event
      } else if (calEvents[currentCategory] !== undefined && sheetEvents[j + 1] !== "") {
        // Compare them; delete and replace if different
        if (compareEvents(calEvents[currentCategory], i, j + 3, currentCategory) === false) {
          calEvents[currentCategory].deleteEvent()
          let event = createEventFromCell(i, j + 3, currentCategory, categories[currentCategory])
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