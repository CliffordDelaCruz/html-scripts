function doGet(e) {
  var name = e.parameter.name; // Retrieve name from URL parameter
  var results = searchAttendance(name); // Call search function

  return ContentService.createTextOutput(JSON.stringify(results))
    .setMimeType(ContentService.MimeType.JSON);
}

function searchAttendance(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toLowerCase().includes(name.toLowerCase())) { // Partial match
      results.push({ id: i, name: data[i][1], date: data[i][2], status: data[i][3] });
    }
  }

  return results; // Send search results back to frontend
}

function updateAttendance(id, date) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  sheet.getRange(id + 1, 3).setValue(date); // Update date
  sheet.getRange(id + 1, 4).setValue("Present"); // Mark as attended
}

function addNewPerson(name, date, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Person_Master");
  sheet.appendRow([new Date(), name, date, status]);
  return "New person added!";
}
