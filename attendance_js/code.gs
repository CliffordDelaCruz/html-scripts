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
 *  - Inserts 15 columns: the mobile number is stored in column 15.
 *  - Copies the date field to the Last_Service_Attended_Date (column 14).
 *  - Automatically marks attendance.
 */
function addNewPerson(name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass, mobile_number) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var newId = getNextId(sheet);
  // Append a new row with 15 columns:
  // [ID, Name, Date, Status, Cell Group, Ministry, Baptized, Disciple 1, Disciple 2, Disciple 3,
  //  Bible Literacy, Financial Literacy, Teaching Class, Last_Service_Attended_Date, Mobile_number]
  sheet.appendRow([newId, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass, date, mobile_number]);
  
  var lastRow = sheet.getLastRow();
  // Format columns as needed
  sheet.getRange(lastRow, 1).setNumberFormat("0");
  sheet.getRange(lastRow, 3).setNumberFormat("@");
  sheet.getRange(lastRow, 14).setNumberFormat("@");
  
  // Automatically create an attendance record.
  markAttendance(newId, name, date);
  
  return "New person added with ID: " + newId;
}

/**
 * Searches the "Person_Master" sheet for records matching the given name.
 * Performs a case‑insensitive, partial match in Column B.
 * Returns all 15 columns.
 */
function searchAttendance(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Now includes 15 columns (the last is Mobile_number)
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
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
 * Updates columns 1–13 and mobile number (column 15) while
 * preserving the existing Last_Service_Attended_Date (column 14).
 * Enforces PDPA: if status is not "One-time visitor" or "For follow-up",
 * the mobile number is cleared.
 */
function updatePersonRecord(id, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass, mobile_number) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var targetRow = -1;
  var existingLastServiceDate = "";
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      targetRow = i + 2; // Adjust for header row.
      existingLastServiceDate = data[i][13];
      break;
    }
  }
  if (targetRow === -1) return "Error: Record with ID " + id + " not found.";
  
  // PDPA Rule: Only retain mobile_number if status is "One-time visitor" or "For follow-up"
  if (status !== "One-time visitor" && status !== "For follow-up") {
    mobile_number = "";
  }
  
  // Update columns 1–13 (basic info)
  sheet.getRange(targetRow, 1, 1, 13).setValues([[id, name, date, status, cellGroup, ministry, baptized, discipleship1, discipleship2, discipleship3, bibleLiteracy, financialLiteracy, teachingClass]]);
  // Reinstate the existing value for Last_Service_Attended_Date (column 14)
  sheet.getRange(targetRow, 14).setValue(existingLastServiceDate);
  // Update Mobile_number (column 15)
  sheet.getRange(targetRow, 15).setValue(mobile_number);
  
  sheet.getRange(targetRow, 1).setNumberFormat("0");
  sheet.getRange(targetRow, 3).setNumberFormat("@");
  sheet.getRange(targetRow, 14).setNumberFormat("@");
  
  return "Record updated successfully for ID: " + id;
}

/**
 * Marks attendance by appending a record into the "Attendance_Table" sheet.
 * Also updates the Last_Service_Attended_Date in Person_Master.
 * BEFORE inserting a new attendance record, the function checks if the selected date
 * matches the current Last_Service_Attended_Date. If so, it returns a message and does not insert.
 */
function markAttendance(id, name, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance_Table");
  if (!attSheet) return "Error: 'Attendance_Table' sheet not found.";
  
  // Get the Person_Master record for the given id.
  var pmSheet = ss.getSheetByName("Person_Master");
  if (!pmSheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRowPM = pmSheet.getLastRow();
  var pmData = pmSheet.getRange(2, 1, lastRowPM - 1, 15).getValues();
  var recordFound = false;
  for (var i = 0; i < pmData.length; i++) {
    if (pmData[i][0] == id) {
      recordFound = true;
      // Check if attendance has already been marked for the selected date.
      if (pmData[i][13] && pmData[i][13].toString() === date) {
        return "Attendance has already been marked for " + date;
      }
      // Otherwise, update Last_Service_Attended_Date to the new date.
      pmSheet.getRange(i + 2, 14).setValue(date);
      pmSheet.getRange(i + 2, 14).setNumberFormat("@");
      break;
    }
  }
  if (!recordFound) return "Error: Record not found in Person_Master.";
  
  var year = date ? date.substring(0, 4) : "";
  attSheet.appendRow([id, name, year, date]);
  
  var attLastRow = attSheet.getLastRow();
  attSheet.getRange(attLastRow, 1).setNumberFormat("0");
  attSheet.getRange(attLastRow, 4).setNumberFormat("@");
  
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
 * For reportType "year": filters by the given year.
 * For reportType "range": filters by a date range.
 * Returns an XLSX export URL.
 */
function generateReport(reportType, year, startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance_Table");
  if (!sheet) return "Error: 'Attendance_Table' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "No records found.";
  
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var filtered = [];
  
  if (reportType === "year") {
    filtered = data.filter(function(row) {
      return parseInt(row[2], 10) === parseInt(year, 10);
    });
  } else if (reportType === "range") {
    var start = new Date(startDate);
    var end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    filtered = data.filter(function(row) {
      var recordDate = new Date(row[3]);
      return recordDate >= start && recordDate <= end;
    });
  }
  
  if (filtered.length === 0) return "No records match the selected criteria.";
  
  var reportSpreadsheet = SpreadsheetApp.create("Attendance Report - " + new Date().toLocaleString());
  var reportSheet = reportSpreadsheet.getActiveSheet();
  
  // Append header row.
  var header = sheet.getRange(1, 1, 1, 4).getValues()[0];
  reportSheet.appendRow(header);
  
  filtered.forEach(function(row) {
    reportSheet.appendRow(row);
  });
  
  var file = DriveApp.getFileById(reportSpreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var fileId = reportSpreadsheet.getId();
  var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  return downloadUrl;
}

/**
 * Extracts person data from the "Person_Master" sheet based on the provided filter.
 * Filter types include: "status", "cellgroup", "ministry", or "date".
 * Returns an XLSX download URL including all columns.
 */
function extractPersonData(filterType, filterValue, startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "No records found.";
  
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var filtered = [];
  
  if (filterType === "status") {
    data.forEach(function(row) {
      if (row[3] === filterValue) filtered.push(row);
    });
  } else if (filterType === "cellgroup") {
    data.forEach(function(row) {
      if (row[4] === filterValue) filtered.push(row);
    });
  } else if (filterType === "ministry") {
    data.forEach(function(row) {
      if (row[5] === filterValue) filtered.push(row);
    });
  } else if (filterType === "date") {
    var sDate = new Date(startDate);
    var eDate = new Date(endDate);
    eDate.setHours(23,59,59,999);
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
  
  var file = DriveApp.getFileById(reportSpreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var fileId = reportSpreadsheet.getId();
  var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  return downloadUrl;
}
