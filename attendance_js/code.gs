/**
 * Serves the HTML web interface using a templated HTML file.
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
 *  - Leaves Last_Service_Attended_Date blank on insert.
 *  - Calls markAttendance() to write both the attendance row and update col 14.
 */
function addNewPerson(
    name, date, status, cellGroup, ministry,
    baptized, discipleship1, discipleship2, discipleship3,
    bibleLiteracy, financialLiteracy, teachingClass, mobile_number
) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var newId = getNextId(sheet);
  
  // Append a new row with 15 columns:
  // [ID, Name, Date, Status, CellGroup, Ministry, Baptized,
  //  Disciple1, Disciple2, Disciple3, BibleLit, FinLit, TeachClass,
  //  Last_Service_Attended_Date (blank), Mobile_number]
  sheet.appendRow([
    newId,
    name,
    date,
    status,
    cellGroup,
    ministry,
    baptized,
    discipleship1,
    discipleship2,
    discipleship3,
    bibleLiteracy,
    financialLiteracy,
    teachingClass,
    "",               // ← leave Last_Service_Attended_Date blank
    mobile_number
  ]);
  
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat("0");
  sheet.getRange(lastRow, 3).setNumberFormat("@");
  sheet.getRange(lastRow, 14).setNumberFormat("@");
  
  // Now create an attendance record AND update Last_Service_Attended_Date
  var markResponse = markAttendance(newId, name, date);
  return "New person added with ID: " + newId + " — " + markResponse;
}

/**
 * Searches the "Person_Master" sheet for records matching the given name.
 * Case-insensitive, partial match in Column B. Returns full 15-col rows.
 */
function searchAttendance(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
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
 * Updates an existing record in "Person_Master". Preserves col 14,
 * enforces PDPA on mobile number, then writes back.
 */
function updatePersonRecord(
    id, name, date, status, cellGroup, ministry,
    baptized, discipleship1, discipleship2, discipleship3,
    bibleLiteracy, financialLiteracy, teachingClass, mobile_number
) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var targetRow = -1;
  var existingLastServiceDate = "";
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      targetRow = i + 2;
      existingLastServiceDate = data[i][13];
      break;
    }
  }
  if (targetRow === -1) return "Error: Record with ID " + id + " not found.";
  
  // PDPA: clear mobile if not visitor/follow-up
  if (status !== "One-time visitor" && status !== "For follow-up") {
    mobile_number = "";
  }
  
  // Write cols 1–13
  sheet.getRange(targetRow, 1, 1, 13).setValues([[
    id, name, date, status, cellGroup,
    ministry, baptized, discipleship1, discipleship2, discipleship3,
    bibleLiteracy, financialLiteracy, teachingClass
  ]]);
  
  // Reinstate existing Last_Service_Attended_Date
  sheet.getRange(targetRow, 14).setValue(existingLastServiceDate);
  sheet.getRange(targetRow, 15).setValue(mobile_number);
  
  sheet.getRange(targetRow, 1).setNumberFormat("0");
  sheet.getRange(targetRow, 3).setNumberFormat("@");
  sheet.getRange(targetRow, 14).setNumberFormat("@");
  
  return "Record updated successfully for ID: " + id;
}

/**
 * Marks attendance by appending to "Attendance_Table" and updating
 * Last_Service_Attended_Date in Person_Master.
 * Prevents duplicate marks for the same date.
 */
function markAttendance(id, name, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance_Table");
  var pmSheet  = ss.getSheetByName("Person_Master");
  if (!attSheet) return "Error: 'Attendance_Table' sheet not found.";
  if (!pmSheet)  return "Error: 'Person_Master' sheet not found.";
  
  // Look up Person_Master for this ID
  var pmData = pmSheet.getRange(2, 1, pmSheet.getLastRow() - 1, 15).getValues();
  var recordFound = false;
  for (var i = 0; i < pmData.length; i++) {
    if (pmData[i][0] == id) {
      recordFound = true;
      // If already marked today, abort
      if (pmData[i][13] && pmData[i][13].toString() === date) {
        return "Attendance already marked for " + date;
      }
      // Update Last_Service_Attended_Date (col 14)
      pmSheet.getRange(i + 2, 14).setValue(date);
      pmSheet.getRange(i + 2, 14).setNumberFormat("@");
      break;
    }
  }
  if (!recordFound) return "Error: Record not found in Person_Master.";
  
  // Append to Attendance_Table: [ID, Name, Year, Date]
  var year = date ? date.substring(0, 4) : "";
  attSheet.appendRow([id, name, year, date]);
  var attLastRow = attSheet.getLastRow();
  attSheet.getRange(attLastRow, 1).setNumberFormat("0");
  attSheet.getRange(attLastRow, 4).setNumberFormat("@");
  
  return "Attendance recorded successfully for " + date;
}

/**
 * Returns a sorted, de-duplicated list of YYYY-MM-DD attendance dates
 * for the given person ID.
 */
function getAttendanceDates(id) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance_Table");
  if (!attSheet) return [];
  var data = attSheet.getRange(2,1, attSheet.getLastRow()-1, 4).getValues();
  var dates = data
    .filter(r => r[0] == id && r[3])
    .map(r => new Date(r[3]).toISOString().slice(0,10))
    // unique
    .filter((v, i, a) => a.indexOf(v) === i)
    // newest first
    .sort((a,b) => b.localeCompare(a));
  return dates;
}

/**
 * Deletes attendance record(s) for id & date.
 * If un‐marked date===today, blanks out Last_Service_Attended_Date.
 * Otherwise, back‐fills it to the next‐most‐recent date.
 */
function unmarkAttendance(id, date) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance_Table");
  var pmSheet  = ss.getSheetByName("Person_Master");
  if (!attSheet || !pmSheet) return "Error: Required sheet missing.";

  // 1) delete all matching rows in Attendance_Table
  var attData = attSheet.getRange(2,1, attSheet.getLastRow()-1, 4).getValues();
  var rowsToRemove = [];
  attData.forEach(function(r,i) {
    if (r[0] == id && new Date(r[3]).toISOString().slice(0,10) === date) {
      rowsToRemove.push(i+2); // +2 for header offset
    }
  });
  if (!rowsToRemove.length) {
    return "No attendance record for ID " + id + " on " + date;
  }
  // delete from bottom up
  rowsToRemove.reverse().forEach(function(rowNum) {
    attSheet.deleteRow(rowNum);
  });

  // 2) recompute Last_Service_Attended_Date
  var remaining = getAttendanceDates(id); // newest first
  var newLast;
  var today = new Date().toISOString().slice(0,10);
  if (date === today) {
    // user removed today → blank it
    newLast = "";
  } else {
    // either leave as-is (if not today) or back-fill
    newLast = remaining.length ? remaining[0] : "";
  }

  // write back to Person_Master
  var pmData = pmSheet.getRange(2,1, pmSheet.getLastRow()-1, 15).getValues();
  for (var i=0; i<pmData.length; i++) {
    if (pmData[i][0] == id) {
      pmSheet.getRange(i+2, 14).setValue(newLast);
      pmSheet.getRange(i+2, 14).setNumberFormat("@");
      break;
    }
  }
  return "Attendance unmarked for " + date;
}


/**
 * Retrieves list of cell groups from "Cell_group_master".
 */
function getCellGroupList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cell_group_master");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1)
              .getValues()
              .map(function(r){ return r[0]; });
}

/**
 * Retrieves list of ministries from "Ministry_master".
 */
function getMinistryList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Ministry_master");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1)
              .getValues()
              .map(function(r){ return r[0]; });
}

/**
 * Generates an XLSX attendance report, filtered by year or date range.
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
    filtered = data.filter(function(r) {
      return parseInt(r[2], 10) === parseInt(year, 10);
    });
  } else if (reportType === "range") {
    var s = new Date(startDate), e = new Date(endDate);
    e.setHours(23,59,59,999);
    filtered = data.filter(function(r) {
      var d = new Date(r[3]);
      return d >= s && d <= e;
    });
  }
  
  if (filtered.length === 0) return "No records match the selected criteria.";
  
  // Build & share a new spreadsheet
  var reportSS = SpreadsheetApp.create("Attendance Report - " + new Date().toLocaleString());
  var reportSh = reportSS.getActiveSheet();
  reportSh.appendRow(sheet.getRange(1, 1, 1, 4).getValues()[0]);
  filtered.forEach(function(r){ reportSh.appendRow(r); });
  
  var file = DriveApp.getFileById(reportSS.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://docs.google.com/spreadsheets/d/" 
         + reportSS.getId() + "/export?format=xlsx";
}

/**
 * Extracts person data from Person_Master by various filters.
 * Returns an XLSX download URL.
 */
function extractPersonData(filterType, filterValue, startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Person_Master");
  if (!sheet) return "Error: 'Person_Master' sheet not found.";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "No records found.";
  
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data   = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
                    .getValues();
  var filtered = [];
  
  if (filterType === "status") {
    filtered = data.filter(function(r){ return r[3] === filterValue; });
  } else if (filterType === "cellgroup") {
    filtered = data.filter(function(r){ return r[4] === filterValue; });
  } else if (filterType === "ministry") {
    filtered = data.filter(function(r){ return r[5] === filterValue; });
  } else if (filterType === "date") {
    var s = new Date(startDate), e = new Date(endDate);
    e.setHours(23,59,59,999);
    filtered = data.filter(function(r){
      var d = new Date(r[2]);
      return !isNaN(d) && d >= s && d <= e;
    });
  } else {
    return "Invalid filter type.";
  }
  
  if (filtered.length === 0) return "No records match the selected criteria.";
  
  var reportSS = SpreadsheetApp.create("Person Data Extraction - " + new Date().toLocaleString());
  var reportSh = reportSS.getActiveSheet();
  reportSh.appendRow(header);
  filtered.forEach(function(r){ reportSh.appendRow(r); });
  
  var file = DriveApp.getFileById(reportSS.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://docs.google.com/spreadsheets/d/" 
         + reportSS.getId() + "/export?format=xlsx";
}
