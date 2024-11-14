/*** @OnlyCurrentDoc */

// Function to format a date as "YYYY-MM-DD" for consistent display
function formatDate(date) {
  var year = date.getFullYear(); // Extract the full year (e.g., 2024)
  var month = String(date.getMonth() + 1).padStart(2, '0'); // Month is zero-based, so add 1; pad with 0 if single digit
  var day = String(date.getDate()).padStart(2, '0'); // Extract the day of the month; pad with 0 if single digit
  return `${year}-${month}-${day}`; // Format as "YYYY-MM-DD"
}

// Function to calculate the duration in minutes given start and end times
function calculateDurationInMinutes(startTime, endTime) {
  var start = new Date(startTime); // Convert start time to Date object
  var end = new Date(endTime); // Convert end time to Date object
  var differenceInMs = end - start; // Calculate the difference in milliseconds
  return differenceInMs / (1000 * 60); // Convert milliseconds to minutes
}

// Function to determine the start date for event retrieval based on the sheet's last row
function getStartDate(sheet, defaultDate) {
  var lastRow = sheet.getLastRow(); // Get the last row with data in the sheet
  
  if (lastRow > 1) { // Check if there is existing data (assumes headers in the first row)
    var lastDate = sheet.getRange(lastRow, 3).getValue(); // Get the latest "Start Time" from column 3 in the last row
    var startDate = new Date(lastDate); // Convert last date to a Date object
    startDate.setDate(startDate.getDate() + 1); // Set start date to one day after the last date
    startDate.setHours(0, 0, 0, 0); // Set time to 12:00 AM (start of the day)
    return startDate; // Return the calculated start date
  } else {
    return new Date(defaultDate); // If no data, use the specified default date
  }
}

// Function to write the details of a single event to the sheet at a specified row
function writeEventDetailsToSheet(sheet, row, eventDetails) {
  // Set values in the sheet for each event detail in the specified columns
  sheet.getRange(row, 1).setValue(eventDetails.dateOfEvent); // Column 1: Date of Event
  sheet.getRange(row, 2).setValue(eventDetails.title); // Column 2: Event Title
  sheet.getRange(row, 3).setValue(eventDetails.start_time); // Column 3: Start Time
  sheet.getRange(row, 4).setValue(eventDetails.end_time); // Column 4: End Time
  sheet.getRange(row, 5).setValue(eventDetails.duration); // Column 5: Duration in minutes
  sheet.getRange(row, 6).setValue(eventDetails.colorOfEvent); // Column 6: Event Color
  sheet.getRange(row, 7).setValue(eventDetails.description); // Column 7: Event Description
}

// Main function to retrieve events from Google Calendar and log them into the spreadsheet
function getEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Access the active spreadsheet
  var sheet = ss.getSheetByName("work_tracker"); // Select the sheet by name "work_tracker"

  // Determine the start date for retrieving events based on existing sheet data
  var startDate = getStartDate(sheet, "November 10, 2024 00:00:00"); // Default start date if no data in sheet
  
  // Set the end date to the current day at 11:59 PM for retrieving today's events
  var endDate = new Date();
  endDate.setHours(23, 59, 59, 999); // Set end time to the end of the day

  var cal = CalendarApp.getCalendarById("youremail@yourdomain.com"); // Get Google Calendar by its ID
  var events = cal.getEvents(startDate, endDate); // Fetch events within the specified date range

  var lastRow = sheet.getLastRow(); // Get the last row in the sheet to start appending events
  var firstUpdateRow = lastRow + 1; // Row where new event data will begin

  // Loop through each event to extract details and log them in the sheet
  events.forEach((event, i) => {
    // Create an object to store details of the current event
    var eventDetails = {
      title: event.getTitle(), // Event title
      start_time: event.getStartTime(), // Event start time
      end_time: event.getEndTime(), // Event end time
      location: event.getLocation(), // Event location
      description: event.getDescription(), // Event description
      colorOfEvent: event.getColor(), // Color ID of the event
      duration: calculateDurationInMinutes(event.getStartTime(), event.getEndTime()), // Duration in minutes
      dateOfEvent: formatDate(event.getStartTime()) // Formatted event date as "YYYY-MM-DD"
    };
    
    // Write the current event's details to the sheet in the appropriate row
    writeEventDetailsToSheet(sheet, firstUpdateRow + i, eventDetails);
  });

  Logger.log("Events have been added to the Spreadsheet"); // Log completion message
}
