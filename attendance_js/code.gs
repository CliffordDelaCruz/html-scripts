/**
 * Serves the HTML web interface using a templated HTML file so that
 * included files (such as CSS) are properly processed.
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate().setTitle('Attendance Tracker');
}

/**
 * Scans Column A (from row 2 onward) to find the highest numeric ID,
 * then returns the next available ID.
 */
function getNextId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1; // No data rows yet.
  
  var idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxId = 0;
  for (var i = 0; i < idRange.length; i++) {
    var num = parseInt(idRange[i][0], 10);
    if (!isNaN(num) && num > maxId) {
      maxId = num;
    }
  }
  return maxId + 1;
}

/**
 * Inserts a new record into the "Person_Master" sheet.
 * New logic:
 *   1. The Last_Service_Attended_Date copies the date field.
 *   2. Automatically creates an entry in Attendance_Table.
 *   3. Formats the Last_Service_Attended_Date (column 14) the same as Date (column 3).
 */
function addNewPerson(name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var newId = getNextId(sheet);
  // Append a new row with 14 columns:
  // [ID, Name, Date, Status, Cell Group, Ministry, Baptized, Disciple 1, Disciple 2, Disciple 3,
  //  Bible Literacy, Financial Literacy, Teaching Class, Last_Service_Attended_Date]
  sheet.appendRow([newId, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass, date]);
  
  var lastRow = sheet.getLastRow();
  // Format the ID column (col 1), Date column (col 3) and Last_Service_Attended_Date column (col 14)
  sheet.getRange(lastRow, 1).setNumberFormat("0");
  sheet.getRange(lastRow, 3).setNumberFormat("@");
  sheet.getRange(lastRow, 14).setNumberFormat("@");
  
  // Automatically create an attendance record in Attendance_Table.
  markAttendance(newId, name, date);
  
  return "New person added with ID: " + newId;
}

/**
 * Searches the "Person_Master" sheet for records matching the given name.
 * It performs a case‑insensitive, partial match in Column B.
 * Returns all 14 columns.
 */
function searchAttendance(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Use 14 columns (including the new Last_Service_Attended_Date)
  var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  var results = [];
  var searchTerm = name.toString().toLowerCase().trim();
  for (var i = 0; i < data.length; i++) {
    var recName = (data[i][1] || "").toString().toLowerCase().trim();
    if (recName.indexOf(searchTerm) !== -1) {
      results.push(data[i]);
    }
  }
  return results;
}

/**
 * Updates an existing record in the "Person_Master" sheet.
 * In this updated function the Last_Service_Attended_Date (column 14) remains unchanged.
 * The new values from the update form are applied to columns 1-13.
 */
function updatePersonRecord(id, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  // Get existing data for 14 columns (to preserve Last_Service_Attended_Date)
  var range = sheet.getRange(2, 1, lastRow - 1, 14);
  var data = range.getValues();
  var targetRow = -1;
  var existingLastServiceDate = ""; // default
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      targetRow = i + 2; // Account for header row.
      existingLastServiceDate = data[i][13]; // Preserve existing Last_Service_Attended_Date
      break;
    }
  }
  if (targetRow == -1) return "Error: Record with ID " + id + " not found.";
  
  // Update columns 1-13 with new values; keep Last_Service_Attended_Date unchanged.
  sheet.getRange(targetRow, 1, 1, 13).setValues([[id, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass]]);
  // Restore the existing Last_Service_Attended_Date in column 14.
  sheet.getRange(targetRow, 14).setValue(existingLastServiceDate);
  
  // Set formatting for the ID, Date and Last_Service_Attended_Date columns.
  sheet.getRange(targetRow, 1).setNumberFormat("0");
  sheet.getRange(targetRow, 3).setNumberFormat("@");
  sheet.getRange(targetRow, 14).setNumberFormat("@");
  
  return "Record updated successfully for ID: " + id;
}

/**
 * Marks attendance by appending a record into the "Attendance_Table" sheet.
 * Also updates the Last_Service_Attended_Date in Person_Master.
 * Ensures that both the ID and the Date in Attendance_Table are formatted properly.
 */
function markAttendance(id, name, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance_Table");
  if (!attSheet) return "Error: 'Attendance_Table' sheet not found.";
  
  var year = date ? date.substring(0, 4) : "";
  attSheet.appendRow([id, name, year, date]);
  
  // Get the last row of Attendance_Table and set the formatting:
  // Format the ID column (column 1) the same as Person_Master and format the Date column (column 4)
  var attLastRow = attSheet.getLastRow();
  attSheet.getRange(attLastRow, 1).setNumberFormat("0");
  attSheet.getRange(attLastRow, 4).setNumberFormat("@");
  
  // Update the corresponding Person_Master record's Last_Service_Attended_Date.
  var pmSheet = ss.getSheetByName("Person_Master");
  if (pmSheet) {
    var lastRowPM = pmSheet.getLastRow();
    var data = pmSheet.getRange(2, 1, lastRowPM - 1, 14).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == id) {
        // Update column 14 (Last_Service_Attended_Date) and set its format
        pmSheet.getRange(i + 2, 14).setValue(date);
        pmSheet.getRange(i + 2, 14).setNumberFormat("@");
        break;
      }
    }
  }
  return "Attendance recorded successfully!";
}

/**
 * Retrieves a list of cell group names from the "Cell_group_master" sheet.
 */
function getCellGroupList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cell_group_master");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(function(row) { return row[0]; });
}

/**
 * Retrieves a list of ministry names from the "Ministry_master" sheet.
 */
function getMinistryList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Ministry_master");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(function(row) { return row[0]; });
}

/**
 * Generates a report based on the specified criteria from Attendance_Table.
 * For "year": filters by the given year.
 * For "range": filters by a date range.
 * Returns an Excel download URL if records are found.
 */
function generateReport(reportType, year, startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance_Table");
  if (!sheet) return "";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "";
  
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var filtered = [];
  
  if (reportType === "year") {
    var reportYear = parseInt(year, 10);
    for (var i = 0; i < data.length; i++) {
      if (parseInt(data[i][2], 10) === reportYear) {
        filtered.push(data[i]);
      }
    }
  } else if (reportType === "range") {
    var start = new Date(startDate);
    var end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    for (var i = 0; i < data.length; i++) {
      var recordDate = new Date(data[i][3]);
      if (recordDate >= start && recordDate <= end) {
        filtered.push(data[i]);
      }
    }
  } else {
    return "";
  }
  
  if (filtered.length === 0) return "";
  
  var reportSpreadsheet = SpreadsheetApp.create("Attendance Report - " + new Date().toLocaleString());
  var reportSheet = reportSpreadsheet.getActiveSheet();
  
  var header = sheet.getRange(1, 1, 1, 4).getValues()[0];
  reportSheet.appendRow(header);
  
  for (var j = 0; j < filtered.length; j++) {
    reportSheet.appendRow(filtered[j]);
  }
  
  var fileId = reportSpreadsheet.getId();
  var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  return downloadUrl;
}

/**
 * Extracts person data from the "Person_Master" sheet based on the provided filter.
 * Filter types: "status", "cellgroup", "ministry", or "date".
 * Returns a download URL for an Excel file or an error message.
 * The extraction now includes the Last_Service_Attended_Date column.
 */
function extractPersonData(filterType, filterValue, startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "No records found.";
  
  // Get all columns including Last_Service_Attended_Date.
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var filtered = [];
  
  if (filterType === "status") {
    data.forEach(function(row) {
      if (row[3] === filterValue) {
        filtered.push(row);
      }
    });
  } else if (filterType === "cellgroup") {
    data.forEach(function(row) {
      if (row[4] === filterValue) {
        filtered.push(row);
      }
    });
  } else if (filterType === "ministry") {
    data.forEach(function(row) {
      if (row[5] === filterValue) {
        filtered.push(row);
      }
    });
  } else if (filterType === "date") {
    var sDate = new Date(startDate);
    var eDate = new Date(endDate);
    eDate.setHours(23, 59, 59, 999);
    data.forEach(function(row) {
      var rowDate = new Date(row[2]);
      if (!isNaN(rowDate) && rowDate >= sDate && rowDate <= eDate) {
        filtered.push(row);
      }
    });
  } else {
    return "Invalid filter type.";
  }
  
  if (filtered.length === 0) return "No records match the selected criteria.";
  
  var reportSpreadsheet = SpreadsheetApp.create("Person Data Extraction - " + new Date().toLocaleString());
  var reportSheet = reportSpreadsheet.getActiveSheet();
  
  reportSheet.appendRow(header);
  filtered.forEach(function(row) {
    reportSheet.appendRow(row);
  });
  
  var fileId = reportSpreadsheet.getId();
  var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  return downloadUrl;
}
