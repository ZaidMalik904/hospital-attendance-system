/**
 * Hospital Attendance & Leave Management System
 * Backend Controller Script (Code.gs)
 */

// Timezone configured for consistency
const TIMEZONE = "Asia/Kolkata";

/**
 * Serves the HTML Web App UI
 */
function doGet(e) {
  // Support API GET requests
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      if (typeof this[action] === 'function') {
        var res = this[action].apply(this, args);
        return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: res }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  var errorMsg = null;
  // Setup sheets if they don't exist before serving
  try {
    setupSheets();
  } catch (err) {
    Logger.log("Setup sheets error: " + err.message);
    errorMsg = err.message;
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hospital Staff Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles HTTP POST requests from external web apps (e.g., Vercel)
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var action = data.action;
    var args = data.args || [];

    if (!action) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No action parameter provided' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (typeof this[action] === 'function') {
      var result = this[action].apply(this, args);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Function ' + action + ' not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper to include CSS/JS sub-templates in Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Trigger function which checks database status when spreadsheet is opened
 */
function onOpen(e) {
  setupSheets();
}

/* =========================================================================
   DATABASE INITIALIZATION & SEEDING (Google Sheets)
   ========================================================================= */

/**
 * Helper to fetch a sheet case-insensitively and trim spaces.
 * Automatically renames the sheet if a case/space mismatch is found.
 */
function getSheetByNameSafe(ss, name) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  var sheets = ss.getSheets();
  var target = name.trim().toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim().toLowerCase();
    if (sName === target) {
      try {
        sheets[i].setName(name);
      } catch (e) {
        Logger.log("Failed to rename sheet: " + e.message);
      }
      return sheets[i];
    }
  }
  return null;
}

/**
 * Auto-creates sheets/tabs and appends headers and dummy testing data if empty.
 */
function setupSheets() {
  var ss = getDbSpreadsheet();
  if (!ss) {
    throw new Error("Active spreadsheet not found. Link this script to a Google Sheet container.");
  }

  // Delete Settings and Activities sheets if they exist in the spreadsheet
  var sheetsToDelete = ["Settings", "Activity", "Activities"];
  sheetsToDelete.forEach(function(sheetName) {
    var sh = getSheetByNameSafe(ss, sheetName);
    if (sh) {
      try {
        ss.deleteSheet(sh);
      } catch(e) {
        Logger.log("Failed to delete sheet " + sheetName + ": " + e.message);
      }
    }
  });

  // One-time cleanup and setup of custom accounts as requested
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("CLEANUP_DONE_V7") !== "true") {
    // A. Re-create Hospitals sheet
    var hospSheet = getSheetByNameSafe(ss, "Hospitals");
    if (hospSheet) ss.deleteSheet(hospSheet);
    hospSheet = ss.insertSheet("Hospitals");
    hospSheet.appendRow(["HospitalID", "Name", "Latitude", "Longitude", "Range"]);
    hospSheet.appendRow(["HOSP001", "St. Jude Hospital", 28.6139, 77.2090, 100]);
    hospSheet.appendRow(["HOSP002", "City General Hospital", 28.6250, 77.2150, 150]);
    hospSheet.appendRow(["HOSP003", "Metro Clinic", 28.6010, 77.1980, 200]);
    hospSheet.setFrozenRows(1);

    // B. Reset Employees to contain precisely the two user-specified rows with Hospital ID
    var empSheet = getSheetByNameSafe(ss, "Employees");
    if (empSheet) {
      empSheet.clear();
      empSheet.appendRow(["EmployeeID", "Name", "Email", "Password", "Role", "Department", "Designation", "JoiningDate", "Status", "HospitalID"]);
      empSheet.appendRow(["EMP001", "Admin User", "mohdzaid4994@gmail.com", "qwerty", "Admin", "Administration", "Hospital Manager", "'01-01-2026", "Active", "HOSP001", "09:00 AM - 05:00 PM (General Shift)"]);
      empSheet.appendRow(["EMP002", "vishu malik", "vishumalik12@gmail.com", "vishu@#2026", "Employee", "cardiology", "consultant", "'08-07-2026", "Active", "HOSP001", "09:00 AM - 05:00 PM (General Shift)"]);
      empSheet.setFrozenRows(1);
    }
    
    // C. Clear Attendance
    var attSheet = getSheetByNameSafe(ss, "Attendance");
    if (attSheet) {
      attSheet.clear();
      attSheet.appendRow(["AttendanceID", "EmployeeID", "Name", "Date", "CheckInTime", "CheckOutTime", "Status", "Remarks", "SelfieURL", "Location"]);
      attSheet.setFrozenRows(1);
    }
    
    // D. Clear Leaves
    var leaveSheet = getSheetByNameSafe(ss, "Leaves");
    if (leaveSheet) {
      leaveSheet.clear();
      leaveSheet.appendRow(["LeaveID", "EmployeeID", "Name", "LeaveType", "FromDate", "ToDate", "NumberOfDays", "Reason", "Status", "AppliedOn", "ActionBy", "ActionOn"]);
      leaveSheet.setFrozenRows(1);
    }
    
    props.setProperty("CLEANUP_DONE_V7", "true");
  }

  // 1. Employees Sheet
  var empSheet = getSheetByNameSafe(ss, "Employees");
  if (!empSheet) {
    empSheet = ss.insertSheet("Employees");
    empSheet.appendRow(["EmployeeID", "Name", "Email", "Password", "Role", "Department", "Designation", "JoiningDate", "Status", "HospitalID"]);
    empSheet.setFrozenRows(1);
  } else {
    var lastCol = empSheet.getLastColumn();
    if (lastCol > 0) {
      var headers = empSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf("HospitalID") === -1) {
        empSheet.getRange(1, headers.length + 1).setValue("HospitalID");
        lastCol = empSheet.getLastColumn();
        headers = empSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }
      if (headers.indexOf("ShiftTime") === -1) {
        empSheet.getRange(1, headers.length + 1).setValue("ShiftTime");
      }
    } else {
      empSheet.appendRow(["EmployeeID", "Name", "Email", "Password", "Role", "Department", "Designation", "JoiningDate", "Status", "HospitalID"]);
    }
  }

  // 2. Hospitals Sheet
  var hospSheet = getSheetByNameSafe(ss, "Hospitals");
  if (!hospSheet) {
    hospSheet = ss.insertSheet("Hospitals");
    hospSheet.appendRow(["HospitalID", "Name", "Latitude", "Longitude", "Range"]);
    hospSheet.setFrozenRows(1);
  }

  // 3. Attendance Sheet
  var attSheet = getSheetByNameSafe(ss, "Attendance");
  if (!attSheet) {
    attSheet = ss.insertSheet("Attendance");
    attSheet.appendRow(["AttendanceID", "EmployeeID", "Name", "Date", "CheckInTime", "CheckOutTime", "Status", "Remarks", "SelfieURL", "Location"]);
    attSheet.setFrozenRows(1);
  } else {
    // Ensure the SelfieURL and Location columns are appended if they are missing from existing sheet
    var lastCol = attSheet.getLastColumn();
    if (lastCol > 0) {
      var headers = attSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf("SelfieURL") === -1) {
        attSheet.getRange(1, headers.length + 1).setValue("SelfieURL");
        headers.push("SelfieURL");
      }
      // Re-fetch last column to append Location safely
      lastCol = attSheet.getLastColumn();
      var currentHeaders = attSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (currentHeaders.indexOf("Location") === -1) {
        attSheet.getRange(1, currentHeaders.length + 1).setValue("Location");
      }
    } else {
      attSheet.appendRow(["AttendanceID", "EmployeeID", "Name", "Date", "CheckInTime", "CheckOutTime", "Status", "Remarks", "SelfieURL", "Location"]);
    }
  }

  // 4. Leaves Sheet
  var leaveSheet = getSheetByNameSafe(ss, "Leaves");
  if (!leaveSheet) {
    leaveSheet = ss.insertSheet("Leaves");
    leaveSheet.appendRow(["LeaveID", "EmployeeID", "Name", "LeaveType", "FromDate", "ToDate", "NumberOfDays", "Reason", "Status", "AppliedOn", "ActionBy", "ActionOn"]);
    leaveSheet.setFrozenRows(1);
  }

  // Seeding dummy data only if sheet headers are the only rows
  seedDummyData(empSheet, attSheet, leaveSheet, null, null, hospSheet);
}

/**
 * Seeds starting dummy records for clinical and administrative staff
 */
function seedDummyData(empSheet, attSheet, leaveSheet, settingsSheet, actSheet, hospSheet) {
  // A. Settings Seeding
  if (settingsSheet && settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(["09:00 AM", "09:15 AM", "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday"]);
  }

  // B. Hospitals Seeding
  if (hospSheet && hospSheet.getLastRow() <= 1) {
    hospSheet.appendRow(["HOSP001", "St. Jude Hospital", 28.6139, 77.2090, 100]);
    hospSheet.appendRow(["HOSP002", "City General Hospital", 28.6250, 77.2150, 150]);
    hospSheet.appendRow(["HOSP003", "Metro Clinic", 28.6010, 77.1980, 200]);
  }

  // C. Employees Seeding (with plain text passwords)
  if (empSheet.getLastRow() <= 1) {
    var dummyEmployees = [
      ["EMP001", "Admin User", "mohdzaid4994@gmail.com", "qwerty", "Admin", "Administration", "Hospital Manager", "'01-01-2026", "Active", "HOSP001", "09:00 AM - 05:00 PM (General Shift)"],
      ["EMP002", "vishu malik", "vishumalik12@gmail.com", "vishu@#2026", "Employee", "cardiology", "consultant", "'08-07-2026", "Active", "HOSP001", "09:00 AM - 05:00 PM (General Shift)"]
    ];
    dummyEmployees.forEach(row => empSheet.appendRow(row));
  }

  // Leaves, Attendance, and Activities are left empty for manual logging by staff.
}

/* =========================================================================
   SECURITY & AUTHENTICATION
   ========================================================================= */

/**
 * Standard SHA-256 Hashing helper
 */
function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hexString = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byteValue = rawHash[i];
    if (byteValue < 0) byteValue += 256;
    var byteString = byteValue.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    hexString += byteString;
  }
  return hexString;
}

/**
 * Authenticates user credentials against Employees Database
 */
function authenticateUser(email, password) {
  var sheet = getDbSpreadsheet().getSheetByName("Employees");
  if (!sheet) return { success: false, message: "Database system error. Employees sheet missing." };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empEmail = row[2];
    var empPass = row[3];
    var empStatus = row[8];

    if (empEmail.toLowerCase() === email.toLowerCase()) {
      if (empPass.toString() === password.toString()) {
        if (empStatus === "Active") {
          var assignedHospId = row[9] || "";
          var assignedHospName = "St. Jude Hospital"; // default fallback
          
          if (assignedHospId) {
            var hospSheet = getDbSpreadsheet().getSheetByName("Hospitals");
            if (hospSheet) {
              var hospData = hospSheet.getDataRange().getValues();
              for (var k = 1; k < hospData.length; k++) {
                if (hospData[k][0] === assignedHospId) {
                  assignedHospName = hospData[k][1];
                  break;
                }
              }
            }
          }

          return {
            success: true,
            user: {
              employeeId: row[0],
              name: row[1],
              email: row[2],
              role: row[4],
              department: row[5],
              designation: row[6],
              hospitalId: assignedHospId,
              hospitalName: assignedHospName,
              shiftTime: row[10] || ""
            }
          };
        } else {
          return { success: false, message: "Account is currently deactivated. Please contact Admin." };
        }
      } else {
        return { success: false, message: "Incorrect password." };
      }
    }
  }
  return { success: false, message: "Account email not registered." };
}

/* =========================================================================
   EMPLOYEE ACTIONS (Attendance Logging & Leaves)
   ========================================================================= */

/**
 * Check-In current employee
 */
function markCheckIn(employeeId, selfieBase64, latitude, longitude, department) {
  var lock = LockService.getScriptLock();
  try {
    // Attempt lock for 10 seconds to handle high concurrency updates
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    var attSheet = ss.getSheetByName("Attendance");
    var settingsSheet = ss.getSheetByName("Settings");
    
    var todayStr = formatDateStr(new Date());

    // 1. Fetch Employee Name, Hospital, and Shift assignment
    var empName = "";
    var assignedHospitalId = "";
    var assignedShiftTime = "";
    var empData = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeId) {
        empName = empData[i][1];
        assignedHospitalId = empData[i][9]; // Column 10 (index 9)
        break;
      }
    }
    
    if (!empName) return { success: false, message: "Employee not found." };
    
    if (!assignedHospitalId) {
      return { success: false, message: "Check-in blocked! You have not been assigned to any hospital by the Administrator." };
    }

    // 2. Validate against active open shift (allow multiple shifts on same day once previous shift is duty out)
    var attData = attSheet.getDataRange().getValues();
    for (var j = attData.length - 1; j >= 1; j--) {
      if (attData[j][1] === employeeId && (!attData[j][5] || attData[j][5] === "")) {
        return { success: false, message: "Duplicate attempt. You already have an active duty. Please duty out first." };
      }
    }

    // 3. Fetch Hospital coordinates and range limit from Hospitals sheet
    var hospSheet = ss.getSheetByName("Hospitals");
    if (!hospSheet) return { success: false, message: "Database system error. Hospitals database missing." };
    
    var hospData = hospSheet.getDataRange().getValues();
    var assignedIds = assignedHospitalId.toString().split(",");
    
    var inRangeHospital = null;
    var distances = [];
    
    for (var k = 1; k < hospData.length; k++) {
      var hId = hospData[k][0];
      if (assignedIds.indexOf(hId) !== -1) {
        var name = hospData[k][1];
        var lat = parseFloat(hospData[k][2]);
        var lng = parseFloat(hospData[k][3]);
        var rad = parseInt(hospData[k][4], 10);
        
        if (isNaN(lat) || isNaN(lng)) continue;
        
        if (latitude !== undefined && longitude !== undefined && latitude !== null && longitude !== null && latitude !== "" && longitude !== "") {
          var dist = getHaversineDistance(parseFloat(latitude), parseFloat(longitude), lat, lng);
          distances.push({ name: name, distance: dist, range: rad });
          if (dist <= rad) {
            inRangeHospital = { name: name, distance: dist, id: hId };
            break;
          }
        }
      }
    }
    
    // 4. Validate Location/Distance
    if (latitude === undefined || longitude === undefined || latitude === null || longitude === null || latitude === "" || longitude === "" || (parseFloat(latitude) === 0 && parseFloat(longitude) === 0)) {
      // Fetch fallback hospital coordinates
      var fallbackHosp = null;
      for (var k = 1; k < hospData.length; k++) {
        var hId = hospData[k][0];
        if (assignedIds.indexOf(hId) !== -1) {
          fallbackHosp = { name: hospData[k][1], lat: parseFloat(hospData[k][2]), lng: parseFloat(hospData[k][3]), id: hId };
          break;
        }
      }
      if (fallbackHosp && !isNaN(fallbackHosp.lat) && !isNaN(fallbackHosp.lng)) {
        latitude = fallbackHosp.lat;
        longitude = fallbackHosp.lng;
        inRangeHospital = { name: fallbackHosp.name, distance: 0, id: fallbackHosp.id };
      } else {
        return { success: false, message: "Location verification failed: Coordinates not provided. Please enable GPS/Location services." };
      }
    }
    
    if (inRangeHospital === null) {
      if (distances.length === 0) {
        return { success: false, message: "Check-in blocked! Assigned hospital location is not configured correctly." };
      }
      var msg = "Check-in blocked! You are not within the range of any assigned hospital:\n";
      distances.forEach(d => {
        msg += "• " + d.name + ": " + Math.round(d.distance) + "m away (limit: " + d.range + "m)\n";
      });
      return { success: false, message: msg };
    }
    
    var hospitalName = inRangeHospital.name;
    var distance = inRangeHospital.distance;

    // 5. Calculate Attendance Status
    var now = new Date();
    var checkInTimeStr = formatTimeStr(now);
    var status = checkInTimeStr.toUpperCase().indexOf("AM") !== -1 ? "Present (Morning Shift)" : "Present (Night Shift)";
    var remarks = "Dept: " + (department || "Not Selected") + " | Verified within " + Math.round(distance) + "m at " + hospitalName;

    // Save selfie to Google Drive folder and get URL
    var selfieUrl = "";
    if (selfieBase64) {
      selfieUrl = saveSelfieToDrive(employeeId, todayStr, selfieBase64);
    }

    // 6. Write to Attendance Sheet with unique attId for multiple shifts per day
    var attId = "ATT" + Utilities.formatDate(now, TIMEZONE, "yyyyMMddHHmmss") + "_" + employeeId;
    
    var lastCol = attSheet.getLastColumn();
    var attHeaders = attSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var locationIdx = attHeaders.indexOf("Location");
    
    var rowValues = [attId, employeeId, empName, "'" + todayStr, checkInTimeStr, "", status, remarks, selfieUrl];
    if (locationIdx !== -1) {
      var mapsFormula = '=HYPERLINK("https://www.google.com/maps?q=' + latitude + ',' + longitude + '", "' + latitude + ', ' + longitude + '")';
      rowValues.push(mapsFormula);
    }
    
    attSheet.appendRow(rowValues);

    return { success: true };
  } catch (e) {
    return { success: false, message: "Concurrency error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Clock-Out current employee
 */
function markCheckOut(employeeId, latitude, longitude) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    var attSheet = ss.getSheetByName("Attendance");
    var todayStr = formatDateStr(new Date());
    
    var attData = attSheet.getDataRange().getValues();
    var rowIndex = -1;

    // Search backwards for the latest open shift (where checkOutTime is empty)
    // This allows duty out of overnight duties that started on a previous day
    for (var i = attData.length - 1; i >= 1; i--) {
      if (attData[i][1] === employeeId && (!attData[i][5] || attData[i][5] === "")) {
        rowIndex = i + 1; // 1-indexed conversion
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "No active duty found. Please Duty In first." };
    }

    // Fetch Employee Hospital assignment
    var assignedHospitalId = "";
    var empData = empSheet.getDataRange().getValues();
    for (var j = 1; j < empData.length; j++) {
      if (empData[j][0] === employeeId) {
        assignedHospitalId = empData[j][9]; // Column 10 (index 9)
        break;
      }
    }
    
    if (!assignedHospitalId) {
      return { success: false, message: "Check-out blocked! You have not been assigned to any hospital by the Administrator." };
    }
    
    var hospSheet = ss.getSheetByName("Hospitals");
    if (!hospSheet) return { success: false, message: "Database system error. Hospitals database missing." };
    
    var hospData = hospSheet.getDataRange().getValues();
    var refLat = null;
    var refLng = null;
    var maxRadius = 100;
    var hospitalName = "";
    
    for (var k = 1; k < hospData.length; k++) {
      if (hospData[k][0] === assignedHospitalId) {
        hospitalName = hospData[k][1];
        refLat = parseFloat(hospData[k][2]);
        refLng = parseFloat(hospData[k][3]);
        maxRadius = parseInt(hospData[k][4], 10);
        break;
      }
    }
    
    if (refLat === null || refLng === null || isNaN(refLat) || isNaN(refLng)) {
      return { success: false, message: "Check-out blocked! Assigned hospital location is not configured correctly." };
    }

    // Validate Location/Distance
    if (latitude === undefined || longitude === undefined || latitude === null || longitude === null || latitude === "" || longitude === "" || (parseFloat(latitude) === 0 && parseFloat(longitude) === 0)) {
      latitude = refLat;
      longitude = refLng;
    }
    
    var distance = getHaversineDistance(parseFloat(latitude), parseFloat(longitude), refLat, refLng);
    if (distance > maxRadius) {
      return { 
        success: false, 
        message: "Check-out blocked! You are " + Math.round(distance) + " meters away from your assigned hospital (" + hospitalName + "). You must be within " + maxRadius + " meters to clock out." 
      };
    }

    // Update CheckOutTime column (Column F is 6th column)
    var checkoutTimeStr = formatTimeStr(new Date());
    attSheet.getRange(rowIndex, 6).setValue(checkoutTimeStr);

    // Append checkout coordinates to the Location cell as clickable hyperlinks
    var lastCol = attSheet.getLastColumn();
    var attHeaders = attSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var locationIdx = attHeaders.indexOf("Location");
    if (locationIdx !== -1) {
      var checkinCoords = attSheet.getRange(rowIndex, locationIdx + 1).getValue();
      var inLat = "";
      var inLng = "";
      if (checkinCoords && String(checkinCoords).indexOf(",") !== -1) {
        var parts = String(checkinCoords).split(",");
        inLat = parts[0].trim();
        inLng = parts[1].trim();
      }
      
      var finalFormula = "";
      if (inLat && inLng) {
        finalFormula = '=HYPERLINK("https://www.google.com/maps?q=' + inLat + ',' + inLng + '", "In Map") & " | " & HYPERLINK("https://www.google.com/maps?q=' + latitude + ',' + longitude + '", "Out Map")';
      } else {
        finalFormula = '=HYPERLINK("https://www.google.com/maps?q=' + latitude + ',' + longitude + '", "Out Map")';
      }
      attSheet.getRange(rowIndex, locationIdx + 1).setValue(finalFormula);
    }
    
    // Append check-out remarks verification
    var remarksRange = attSheet.getRange(rowIndex, 8); // Column H (8th)
    var currentRemarks = remarksRange.getValue();
    remarksRange.setValue((currentRemarks ? currentRemarks + "; " : "") + "Out verified within " + Math.round(distance) + "m");

    return { success: true };
  } catch (e) {
    return { success: false, message: "Error completing checkout: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns today's attendance row for the current employee (if any)
 */
function getTodayAttendance(employeeId) {
  var attSheet = getDbSpreadsheet().getSheetByName("Attendance");
  var todayStr = formatDateStr(new Date());
  var data = attSheet.getDataRange().getValues();

  // 1. Check if there is an active open shift (even from a previous day for overnight duties)
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === employeeId && (!data[i][5] || data[i][5] === "")) {
      return {
        checkInTime: data[i][4] ? (data[i][4] instanceof Date ? formatTimeStr(data[i][4]) : String(data[i][4])) : "",
        checkOutTime: "",
        status: data[i][6],
        remarks: data[i][7]
      };
    }
  }

  // 2. Otherwise, return the latest shift for today
  for (var j = data.length - 1; j >= 1; j--) {
    if (data[j][1] === employeeId && cellValueToDateStr(data[j][3]) === todayStr) {
      return {
        checkInTime: data[j][4] ? (data[j][4] instanceof Date ? formatTimeStr(data[j][4]) : String(data[j][4])) : "",
        checkOutTime: data[j][5] ? (data[j][5] instanceof Date ? formatTimeStr(data[j][5]) : String(data[j][5])) : "",
        status: data[j][6],
        remarks: data[j][7]
      };
    }
  }
  return null;
}

/**
 * Fetch monthly attendance metrics summary for the employee
 */
function getMonthlyAttendanceSummary(employeeId, month, year) {
  var attSheet = getDbSpreadsheet().getSheetByName("Attendance");
  var data = attSheet.getDataRange().getValues();
  
  var present = 0;
  var late = 0;
  var absent = 0;
  var leave = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empId = row[1];
    var dateStr = cellValueToDateStr(row[3]); // safe conversion
    var status = row[6];

    if (empId === employeeId) {
      var parts = dateStr.split('-');
      var rMonth = parseInt(parts[1], 10);
      var rYear = parseInt(parts[2], 10);

      if (rMonth === month && rYear === year) {
        if (status.indexOf("Present") !== -1 || status === "Late") present++;
        else if (status === "Absent") absent++;
        else if (status === "On Leave") leave++;
      }
    }
  }

  // Fetch today's status too to return bundled
  var todayStatus = getTodayAttendance(employeeId);

  return {
    present: present,
    late: late,
    absent: absent,
    leave: leave,
    today: todayStatus
  };
}

/**
 * Logs a new leave application
 */
function applyLeave(employeeId, leaveType, fromDate, toDate, reason) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    var leaveSheet = ss.getSheetByName("Leaves");
    
    // Find name
    var empName = "";
    var empData = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeId) {
        empName = empData[i][1];
        break;
      }
    }

    if (!empName) return { success: false, message: "Employee not found." };

    // Format dates to dd-MM-yyyy for consistency
    var fromParsed = new Date(fromDate);
    var toParsed = new Date(toDate);
    var fromStr = formatDateStr(fromParsed);
    var toStr = formatDateStr(toParsed);

    // Compute range duration
    var diffTime = Math.abs(toParsed - fromParsed);
    var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

    var leaveId = "LEV" + String(new Date().getTime()).substring(5);
    var appliedOnStr = formatDateStr(new Date()) + " " + formatTimeStr(new Date());

    leaveSheet.appendRow([
      leaveId,
      employeeId,
      empName,
      leaveType,
      "'" + fromStr,
      "'" + toStr,
      diffDays,
      reason,
      "Pending",
      "'" + appliedOnStr,
      "",
      ""
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, message: "Error submitting leave: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns leave history for a specific employee
 */
function getMyLeaves(employeeId) {
  var leaveSheet = getDbSpreadsheet().getSheetByName("Leaves");
  var data = leaveSheet.getDataRange().getValues();
  var myLeaves = [];

  for (var i = data.length - 1; i >= 1; i--) { // Reverse order to display newest first
    if (data[i][1] === employeeId) {
      myLeaves.push({
        leaveId: data[i][0],
        leaveType: data[i][3],
        fromDate: cellValueToDateStr(data[i][4]),
        toDate: cellValueToDateStr(data[i][5]),
        days: data[i][6],
        reason: data[i][7],
        status: data[i][8],
        appliedOn: data[i][9]
      });
    }
  }
  return myLeaves;
}

/* =========================================================================
   ADMIN ACTIONS (Attendance Monitor, Leaves Approval & Roster)
   ========================================================================= */

/**
 * Return attendance records for a specific date (dd-MM-yyyy)
 */
function getAttendanceByDate(dateStr, hospitalId) {
  var ss = getDbSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance");
  var empSheet = ss.getSheetByName("Employees");
  
  var attData = attSheet.getDataRange().getValues();
  var empData = empSheet.getDataRange().getValues();
  
  var results = [];
  var markedMap = {};

  var empHospMap = {};
  for (var k = 1; k < empData.length; k++) {
    empHospMap[empData[k][0]] = empData[k][9] || "";
  }

  // 1. Gather all recorded attendance for the target date
  for (var i = 1; i < attData.length; i++) {
    if (cellValueToDateStr(attData[i][3]) === dateStr) {
      var empId = attData[i][1];
      var empHosp = empHospMap[empId] || "";
      if (hospitalId && empHosp !== hospitalId) continue;
      
      results.push({
        attendanceId: attData[i][0],
        employeeId: empId,
        name: attData[i][2],
        checkInTime: attData[i][4] ? (attData[i][4] instanceof Date ? formatTimeStr(attData[i][4]) : String(attData[i][4])) : "",
        checkOutTime: attData[i][5] ? (attData[i][5] instanceof Date ? formatTimeStr(attData[i][5]) : String(attData[i][5])) : "",
        status: attData[i][6],
        remarks: attData[i][7],
        selfieUrl: attData[i][8] || ""
      });
      markedMap[empId] = true;
    }
  }

  // 2. Add Absent markers for other active employees who didn't clock in
  // Check if date is Sunday
  var dateParts = dateStr.split('-');
  var targetDate = new Date(parseInt(dateParts[2]), parseInt(dateParts[1]) - 1, parseInt(dateParts[0]));
  var isSunday = (targetDate.getDay() === 0);

  // Skip auto-absent logs on Sunday unless specified
  if (!isSunday) {
    // If targetDate is in the future, don't auto-mark absent
    var today = new Date();
    today.setHours(0,0,0,0);
    targetDate.setHours(0,0,0,0);

    if (targetDate <= today) {
      for (var j = 1; j < empData.length; j++) {
        var empId = empData[j][0];
        var name = empData[j][1];
        var role = empData[j][4];
        var status = empData[j][8];
        var empHosp = empData[j][9] || "";

        if (status === "Active" && role !== "Admin" && !markedMap[empId]) {
          if (hospitalId && empHosp !== hospitalId) continue;
          results.push({
            attendanceId: "",
            employeeId: empId,
            name: name,
            checkInTime: "",
            checkOutTime: "",
            status: "Absent",
            remarks: "No record registered"
          });
        }
      }
    }
  }

  return results;
}

/**
 * Return pending leave applications
 */
function getPendingLeaves(hospitalId) {
  var ss = getDbSpreadsheet();
  var leaveSheet = ss.getSheetByName("Leaves");
  var empSheet = ss.getSheetByName("Employees");
  
  var data = leaveSheet.getDataRange().getValues();
  var empData = empSheet.getDataRange().getValues();
  
  var empHospMap = {};
  for (var k = 1; k < empData.length; k++) {
    empHospMap[empData[k][0]] = empData[k][9] || "";
  }
  
  var pending = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === "Pending") {
      var empId = data[i][1];
      var empHosp = empHospMap[empId] || "";
      if (hospitalId && empHosp !== hospitalId) continue;
      
      pending.push({
        leaveId: data[i][0],
        employeeId: empId,
        name: data[i][2],
        leaveType: data[i][3],
        fromDate: cellValueToDateStr(data[i][4]),
        toDate: cellValueToDateStr(data[i][5]),
        days: data[i][6],
        reason: data[i][7],
        appliedOn: data[i][9]
      });
    }
  }
  return pending;
}

/**
 * Approves a leave request and auto-updates the Attendance records for dates
 */
function approveLeave(leaveId, adminId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var leaveSheet = ss.getSheetByName("Leaves");
    var attSheet = ss.getSheetByName("Attendance");
    
    var leaveData = leaveSheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < leaveData.length; i++) {
      if (leaveData[i][0] === leaveId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Leave record not found." };
    if (leaveData[rowIndex - 1][8] !== "Pending") return { success: false, message: "Leave request is already processed." };

    var employeeId = leaveData[rowIndex - 1][1];
    var empName = leaveData[rowIndex - 1][2];
    var fromStr = leaveData[rowIndex - 1][4]; // dd-MM-yyyy
    var toStr = leaveData[rowIndex - 1][5];   // dd-MM-yyyy

    var nowStr = formatDateStr(new Date()) + " " + formatTimeStr(new Date());

    // Update Status, ActionBy, ActionOn
    leaveSheet.getRange(rowIndex, 9).setValue("Approved");
    leaveSheet.getRange(rowIndex, 11).setValue(adminId);
    leaveSheet.getRange(rowIndex, 12).setValue(nowStr);

    // Auto-mark "On Leave" on the Attendance Sheet for each date in range
    var fromParts = fromStr.split('-');
    var toParts = toStr.split('-');
    
    var startDate = new Date(parseInt(fromParts[2]), parseInt(fromParts[1]) - 1, parseInt(fromParts[0]));
    var endDate = new Date(parseInt(toParts[2]), parseInt(toParts[1]) - 1, parseInt(toParts[0]));

    var attData = attSheet.getDataRange().getValues();

    // Loop through each calendar day in the range
    for (var date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
      // Skip Sunday for attendance updates
      if (date.getDay() === 0) continue;

      var currentStr = formatDateStr(date);
      var existsIndex = -1;

      for (var j = 1; j < attData.length; j++) {
        if (attData[j][1] === employeeId && attData[j][3] === currentStr) {
          existsIndex = j + 1;
          break;
        }
      }

      if (existsIndex !== -1) {
        // Update existing attendance status (late or absent is overwritten by approved leaves)
        attSheet.getRange(existsIndex, 5).setValue(""); // Clear CheckInTime
        attSheet.getRange(existsIndex, 6).setValue(""); // Clear CheckOutTime
        attSheet.getRange(existsIndex, 7).setValue("On Leave");
        attSheet.getRange(existsIndex, 8).setValue("Approved Leave " + leaveId);
      } else {
        // Create new On Leave record
        var attId = "ATT" + Utilities.formatDate(date, TIMEZONE, "yyyyMMdd") + employeeId;
        attSheet.appendRow([attId, employeeId, empName, currentStr, "", "", "On Leave", "Approved Leave " + leaveId]);
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: "Error approving leave: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rejects a leave request
 */
function rejectLeave(leaveId, adminId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var leaveSheet = ss.getSheetByName("Leaves");
    var leaveData = leaveSheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < leaveData.length; i++) {
      if (leaveData[i][0] === leaveId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Leave record not found." };
    if (leaveData[rowIndex - 1][8] !== "Pending") return { success: false, message: "Leave request is already processed." };

    var nowStr = formatDateStr(new Date()) + " " + formatTimeStr(new Date());

    // Update Status, ActionBy, ActionOn
    leaveSheet.getRange(rowIndex, 9).setValue("Rejected");
    leaveSheet.getRange(rowIndex, 11).setValue(adminId);
    leaveSheet.getRange(rowIndex, 12).setValue(nowStr);

    return { success: true };
  } catch (e) {
    return { success: false, message: "Error rejecting leave: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Register a new employee (validates email uniqueness)
 */
function addEmployee(details) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    var empData = empSheet.getDataRange().getValues();

    // Check email uniqueness
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][2].toLowerCase() === details.email.toLowerCase()) {
        return { success: false, message: "Email is already registered under another account." };
      }
    }

    // Auto-generate employee ID
    var maxIdVal = 0;
    for (var j = 1; j < empData.length; j++) {
      var idNum = parseInt(empData[j][0].replace("EMP", ""), 10);
      if (!isNaN(idNum) && idNum > maxIdVal) {
        maxIdVal = idNum;
      }
    }
    
    var nextId = "EMP" + String(maxIdVal + 1).padStart(3, '0');
    var todayStr = formatDateStr(new Date());

    empSheet.appendRow([
      nextId,
      details.name,
      details.email,
      details.password,
      details.role,
      "", // Department column kept for backwards compatibility but empty
      details.designation,
      "'" + todayStr,
      "Active",
      details.hospitalId || ""
    ]);

    return { success: true, employeeId: nextId };
  } catch (e) {
    return { success: false, message: "Error adding staff: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get roster list of all registered employees
 */
function getAllEmployees(hospitalId) {
  var empSheet = getDbSpreadsheet().getSheetByName("Employees");
  var data = empSheet.getDataRange().getValues();
  var list = [];

  for (var i = 1; i < data.length; i++) {
    var empHospitalId = data[i][9] || "";
    if (hospitalId && empHospitalId !== hospitalId) continue;
    list.push({
      employeeId: data[i][0],
      name: data[i][1],
      email: data[i][2],
      role: data[i][4],
      department: data[i][5],
      designation: data[i][6],
      joiningDate: cellValueToDateStr(data[i][7]),
      status: data[i][8],
      hospitalId: empHospitalId
    });
  }
  return list;
}

/**
 * Toggles status (Active/Inactive) of employees
 */
function updateEmployeeStatus(employeeId, status) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    var data = empSheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Employee not found." };

    empSheet.getRange(rowIndex, 9).setValue(status);
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error updating status: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Compiles monthly stats report for all employees
 */
function getMonthlyReportsData(month, year) {
  var ss = getDbSpreadsheet();
  var empSheet = ss.getSheetByName("Employees");
  var attSheet = ss.getSheetByName("Attendance");
  
  var empData = empSheet.getDataRange().getValues();
  var attData = attSheet.getDataRange().getValues();
  
  var reports = [];

  // Loop through all employees (except admins)
  for (var i = 1; i < empData.length; i++) {
    var empId = empData[i][0];
    var empName = empData[i][1];
    var empRole = empData[i][4];
    var empDept = empData[i][5];
    var empStatus = empData[i][8];

    if (empRole === "Admin" || empStatus === "Inactive") continue;

    var present = 0;
    var late = 0;
    var absent = 0;
    var leave = 0;

    // Filter and count attendance logs for target month & year
    for (var j = 1; j < attData.length; j++) {
      if (attData[j][1] === empId) {
        var dateStr = cellValueToDateStr(attData[j][3]); // safe conversion
        var status = attData[j][6];

        var parts = dateStr.split('-');
        var rMonth = parseInt(parts[1], 10);
        var rYear = parseInt(parts[2], 10);

        if (rMonth === month && rYear === year) {
          if (status.indexOf("Present") !== -1 || status === "Late") present++;
          else if (status === "Absent") absent++;
          else if (status === "On Leave") leave++;
        }
      }
    }

    reports.push({
      employeeId: empId,
      name: empName,
      department: empDept,
      present: present,
      late: late,
      absent: absent,
      leave: leave
    });
  }

  return reports;
}

/**
 * Retrieve performance reports data based on custom from/to dates
 */
function getCustomReportsData(fromDateStr, toDateStr, hospitalId) {
  var ss = getDbSpreadsheet();
  var empSheet = ss.getSheetByName("Employees");
  var attSheet = ss.getSheetByName("Attendance");
  var leaveSheet = ss.getSheetByName("Leaves");
  
  var empData = empSheet ? empSheet.getDataRange().getValues() : [];
  var attData = attSheet ? attSheet.getDataRange().getValues() : [];
  var leaveData = leaveSheet ? leaveSheet.getDataRange().getValues() : [];
  
  var reports = [];
  
  var fromDate = fromDateStr ? new Date(fromDateStr) : null;
  var toDate = toDateStr ? new Date(toDateStr) : null;
  
  if (fromDate) fromDate.setHours(0, 0, 0, 0);
  if (toDate) toDate.setHours(23, 59, 59, 999);

  // Loop through all employees (except admins)
  for (var i = 1; i < empData.length; i++) {
    var empId = empData[i][0];
    var empName = empData[i][1];
    var empRole = empData[i][4];
    var empDept = empData[i][5] || "General";
    var empStatus = empData[i][8];
    var empHospitalId = empData[i][9] || "";

    if (empRole === "Admin" || empStatus === "Inactive") continue;
    if (hospitalId && empHospitalId !== hospitalId) continue;

    var days = {};
    for (var d = 1; d <= 31; d++) {
      days[d] = [];
    }

    var presentCount = 0;
    var lateCount = 0;
    var absentCount = 0;
    var leaveCount = 0;

    // Attendance Data
    for (var j = 1; j < attData.length; j++) {
      if (attData[j][1] === empId) {
        var dateStr = cellValueToDateStr(attData[j][3]);
        if (!dateStr) continue;

        var parts = dateStr.split('-');
        if (parts.length < 3) continue;
        var logDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
        logDate.setHours(0, 0, 0, 0);

        var match = true;
        if (fromDate && logDate < fromDate) match = false;
        if (toDate && logDate > toDate) match = false;

        if (match) {
          var dayOfMonth = parseInt(parts[0], 10);
          var checkInTime = attData[j][4] ? formatTimeStr(attData[j][4]) : "";
          var checkOutTime = attData[j][5] ? formatTimeStr(attData[j][5]) : "";
          var status = attData[j][6] || "";

          if (status.indexOf("Present") !== -1) presentCount++;
          else if (status === "Late") { presentCount++; lateCount++; }
          else if (status === "Absent") absentCount++;
          else if (status === "On Leave") leaveCount++;
          
          if (dayOfMonth >= 1 && dayOfMonth <= 31) {
            days[dayOfMonth].push({
              in: checkInTime,
              out: checkOutTime,
              status: status,
              workHours: attData[j][7] || 0
            });
          }
        }
      }
    }

    // Approved Leaves Data
    for (var k = 1; k < leaveData.length; k++) {
      if (leaveData[k][1] === empId && (leaveData[k][8] === "Approved" || leaveData[k][8] === "Approved by System")) {
        var lFromStr = cellValueToDateStr(leaveData[k][4]);
        var lToStr = cellValueToDateStr(leaveData[k][5]);
        if (!lFromStr || !lToStr) continue;

        var fParts = lFromStr.split('-');
        var tParts = lToStr.split('-');
        var lFromDate = new Date(parseInt(fParts[2], 10), parseInt(fParts[1], 10) - 1, parseInt(fParts[0], 10));
        var lToDate = new Date(parseInt(tParts[2], 10), parseInt(tParts[1], 10) - 1, parseInt(tParts[0], 10));
        lFromDate.setHours(0, 0, 0, 0);
        lToDate.setHours(23, 59, 59, 999);

        // Check overlapping range
        var checkFrom = fromDate || lFromDate;
        var checkTo = toDate || lToDate;

        var curr = new Date(Math.max(lFromDate.getTime(), checkFrom.getTime()));
        var endBound = new Date(Math.min(lToDate.getTime(), checkTo.getTime()));

        while (curr <= endBound) {
          var dayNum = curr.getDate();
          if (days[dayNum] && days[dayNum].length === 0) {
            days[dayNum].push({
              in: "",
              out: "",
              status: "On Leave",
              workHours: 0
            });
            leaveCount++;
          }
          curr.setDate(curr.getDate() + 1);
        }
      }
    }

    reports.push({
      employeeId: empId,
      name: empName,
      department: empDept,
      hospitalId: empHospitalId,
      days: days,
      present: presentCount,
      late: lateCount,
      absent: absentCount,
      leave: leaveCount
    });
  }

  return reports;
}

/* =========================================================================
   DATE/TIME CONVERTER UTILITIES
   ========================================================================= */

/**
 * Format JS Date object to dd-MM-yyyy string
 */
function formatDateStr(date) {
  return Utilities.formatDate(date, TIMEZONE, "dd-MM-yyyy");
}

/**
 * Format JS Date object to hh:mm a string
 */
function formatTimeStr(date) {
  return Utilities.formatDate(date, TIMEZONE, "hh:mm a");
}

/**
 * Parse time string "hh:mm a" into a comparable numerical value (minutes since midnight)
 */
function parseTimeString(timeStr) {
  if (!timeStr) return 0;
  
  // Format Date objects using formatTimeStr, or cast to trimmed String
  if (timeStr instanceof Date) {
    timeStr = formatTimeStr(timeStr);
  } else {
    timeStr = String(timeStr).trim();
  }
  
  // Format expectation: "hh:mm AM" or "hh:mm PM"
  var match = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return 0;

  var hours = parseInt(match[1], 10);
  var minutes = parseInt(match[2], 10);
  var meridian = match[3].toUpperCase();

  if (meridian === "PM" && hours !== 12) {
    hours += 12;
  } else if (meridian === "AM" && hours === 12) {
    hours = 0;
  }

  return hours * 60 + minutes;
}

/**
 * Helper to fetch spreadsheet. If container-bound active sheet is not found,
 * it falls back to looking up/creating a sheet in Google Drive using Script Properties.
 */
function getDbSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // Ignore error
  }
  
  if (ss) {
    return ss;
  }
  
  // Standalone script fallback
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty("SPREADSHEET_ID");
  
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      // If deleted/inaccessible
      props.deleteProperty("SPREADSHEET_ID");
    }
  }
  
  // Create a new spreadsheet in user's Drive
  try {
    ss = SpreadsheetApp.create("Hospital Database (Auto-Created)");
    props.setProperty("SPREADSHEET_ID", ss.getId());
    return ss;
  } catch (e) {
    throw new Error("Could not access or create Google Sheet database. Please link to a sheet.");
  }
}

/**
 * Expose URL of the database sheet
 */
function getSpreadsheetUrl() {
  try {
    var ss = getDbSpreadsheet();
    return ss.getUrl();
  } catch (e) {
    return null;
  }
}

/**
 * Diagnostics helper function for setup troubleshooting
 */
function runSetupDiagnostics() {
  var report = {
    spreadsheetFound: false,
    spreadsheetUrl: null,
    sheetsCreated: [],
    errors: [],
    employeesCount: 0,
    settingsCount: 0,
    attendanceCount: 0,
    leavesCount: 0
  };

  try {
    // 1. Force setup sheets to run
    setupSheets();
    
    // 2. Fetch Spreadsheet details
    var ss = getDbSpreadsheet();
    report.spreadsheetFound = true;
    report.spreadsheetUrl = ss.getUrl();

    // 3. Inspect sheets stats
    var sheets = ["Employees", "Attendance", "Leaves", "Settings"];
    sheets.forEach(name => {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        report.sheetsCreated.push(name);
        var rowCount = sheet.getLastRow();
        if (name === "Employees") report.employeesCount = rowCount;
        if (name === "Settings") report.settingsCount = rowCount;
        if (name === "Attendance") report.attendanceCount = rowCount;
        if (name === "Leaves") report.leavesCount = rowCount;
      }
    });

  } catch (e) {
    report.errors.push(e.message);
  }

  return report;
}

/**
 * Helper to normalize any cell date value (String or Date object) into dd-MM-yyyy string
 */
function cellValueToDateStr(cellVal) {
  if (!cellVal) return "";
  
  // Format Date objects using standard dd-MM-yyyy format
  if (cellVal instanceof Date) {
    return formatDateStr(cellVal);
  }
  
  var str = String(cellVal).trim();
  
  // Already in dd-MM-yyyy format
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) {
    return str;
  }
  
  // ISO date format yyyy-MM-dd
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return isoMatch[3] + "-" + isoMatch[2] + "-" + isoMatch[1];
  }
  
  // Slash date format yyyy/MM/dd
  var slashMatch1 = str.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  if (slashMatch1) {
    return slashMatch1[3] + "-" + slashMatch1[2] + "-" + slashMatch1[1];
  }
  
  // Slash date format dd/MM/yyyy or d/m/yyyy
  var slashMatch2 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch2) {
    var day = String(slashMatch2[1]).padStart(2, '0');
    var month = String(slashMatch2[2]).padStart(2, '0');
    return day + "-" + month + "-" + slashMatch2[3];
  }
  
  // General parsing try
  try {
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return formatDateStr(parsed);
    }
  } catch (e) {
    // Ignore error
  }
  
  return str;
}

/**
 * Saves a base64 encoded image to Google Drive inside 'Hospital Attendance Selfies' folder
 */
function saveSelfieToDrive(employeeId, dateStr, base64Data) {
  if (!base64Data) return "";
  try {
    var parts = base64Data.split(',');
    var contentType = parts[0].split(':')[1].split(';')[0];
    var rawData = parts[1];
    var decoded = Utilities.base64Decode(rawData);
    var blob = Utilities.newBlob(decoded, contentType, "Selfie_" + employeeId + "_" + dateStr.replace(/-/g, "") + ".png");
    
    var folderName = "Hospital Attendance Selfies";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log("Selfie upload failed: " + e.message);
    return "Upload failed: " + e.message;
  }
}

/**
 * Compiles a system-wide timeline of employee actions (attendance check-ins, check-outs, leave requests, leave actions)
 */
function getSystemActivityLogs() {
  var ss = getDbSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance");
  var leaveSheet = ss.getSheetByName("Leaves");
  
  var attData = attSheet ? attSheet.getDataRange().getValues() : [];
  var leaveData = leaveSheet ? leaveSheet.getDataRange().getValues() : [];
  
  var logs = [];

  // 1. Process Attendance records
  if (attData.length > 1) {
    for (var i = 1; i < attData.length; i++) {
      var row = attData[i];
      var empId = row[1];
      var name = row[2];
      var dateStr = cellValueToDateStr(row[3]);
      var checkin = row[4];
      var checkout = row[5];
      var status = row[6];
      var remarks = row[7];
      var selfie = row[8] || "";

      if (checkin) {
        logs.push({
          timestamp: parseDateTime(dateStr + " " + checkin),
          timeDisplay: dateStr + " " + checkin,
          employee: name + " (" + empId + ")",
          type: "Check-In",
          badgeClass: "badge present",
          icon: "fa-sign-in-alt",
          description: "Clocked In. Status: " + status + (remarks ? " (" + remarks + ")" : ""),
          link: selfie
        });
      }
      if (checkout) {
        logs.push({
          timestamp: parseDateTime(dateStr + " " + checkout),
          timeDisplay: dateStr + " " + checkout,
          employee: name + " (" + empId + ")",
          type: "Check-Out",
          badgeClass: "badge late",
          icon: "fa-sign-out-alt",
          description: "Clocked Out. Completed Shift.",
          link: ""
        });
      }
    }
  }

  // 2. Process Leave records
  if (leaveData.length > 1) {
    for (var j = 1; j < leaveData.length; j++) {
      var lRow = leaveData[j];
      var leaveId = lRow[0];
      var empId = lRow[1];
      var name = lRow[2];
      var leaveType = lRow[3];
      var fromDate = cellValueToDateStr(lRow[4]);
      var toDate = cellValueToDateStr(lRow[5]);
      var days = lRow[6];
      var reason = lRow[7];
      var status = lRow[8];
      var appliedOn = cellValueToDateStr(lRow[9]);
      var actionBy = lRow[10];
      var actionOn = cellValueToDateStr(lRow[11]);

      if (appliedOn) {
        logs.push({
          timestamp: parseDateTime(appliedOn),
          timeDisplay: appliedOn,
          employee: name + " (" + empId + ")",
          type: "Leave Applied",
          badgeClass: "badge pending",
          icon: "fa-calendar-plus",
          description: "Requested " + days + " days of " + leaveType + " leave. Reason: " + reason,
          link: ""
        });
      }

      if (actionOn) {
        var actionText = status === "Approved" ? "Approved" : "Rejected";
        var badge = status === "Approved" ? "badge present" : "badge absent";
        logs.push({
          timestamp: parseDateTime(actionOn),
          timeDisplay: actionOn,
          employee: "Admin (" + actionBy + ")",
          type: "Leave " + actionText,
          badgeClass: badge,
          icon: status === "Approved" ? "fa-check-double" : "fa-times-circle",
          description: actionText + " leave request " + leaveId + " for " + name + " (" + empId + ")",
          link: ""
        });
      }
    }
  }

  // 3. Sort logs descending by timestamp
  logs.sort(function(a, b) {
    return b.timestamp - a.timestamp;
  });

  return logs;
}

/**
 * Helper to parse custom "dd-MM-yyyy hh:mm a" into JS Date timestamp
 */
function parseDateTime(dtStr) {
  if (!dtStr) return 0;
  try {
    var parts = dtStr.split(' ');
    var dateParts = parts[0].split('-');
    var day = parseInt(dateParts[0], 10);
    var month = parseInt(dateParts[1], 10) - 1;
    var year = parseInt(dateParts[2], 10);

    if (parts.length < 2) {
      return new Date(year, month, day).getTime();
    }

    var timeParts = parts[1].split(':');
    var hours = parseInt(timeParts[0], 10);
    var minutes = parseInt(timeParts[1], 10);
    var meridian = parts[2] ? parts[2].toUpperCase() : "AM";
    
    // Check if meridian is at index 3 due to double spacing
    for (var i = 2; i < parts.length; i++) {
      var val = parts[i].toUpperCase();
      if (val === "AM" || val === "PM") {
        meridian = val;
        break;
      }
    }
    
    if (meridian === "PM" && hours !== 12) {
      hours += 12;
    } else if (meridian === "AM" && hours === 12) {
      hours = 0;
    }

    return new Date(year, month, day, hours, minutes).getTime();
  } catch (e) {
    return 0;
  }
}

/**
 * Logs a new employee activity
 */
function addEmployeeActivity(employeeId, details, duration, category) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    
    // Find name
    var empName = "";
    var empData = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeId) {
        empName = empData[i][1];
        break;
      }
    }
    
    if (!empName) return { success: false, message: "Employee not found." };
    
    var now = new Date();
    var dateStr = formatDateStr(now);
    var timeStr = formatTimeStr(now);
    var activityId = "ACT" + String(now.getTime()).substring(5);
    
    var props = PropertiesService.getScriptProperties();
    var activitiesJson = props.getProperty("SYSTEM_ACTIVITIES") || "[]";
    var activities = JSON.parse(activitiesJson);
    
    activities.push({
      activityId: activityId,
      employeeId: employeeId,
      name: empName,
      date: dateStr,
      time: timeStr,
      details: details,
      duration: duration,
      category: category
    });
    
    // Keep only last 500 logs to prevent script properties size limit
    if (activities.length > 500) {
      activities.shift();
    }
    
    props.setProperty("SYSTEM_ACTIVITIES", JSON.stringify(activities));
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error saving activity: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns activity log for a specific employee on a specific date (default today)
 */
function getEmployeeActivities(employeeId, dateStr) {
  try {
    if (!dateStr) {
      dateStr = formatDateStr(new Date());
    }
    var props = PropertiesService.getScriptProperties();
    var activitiesJson = props.getProperty("SYSTEM_ACTIVITIES") || "[]";
    var activities = JSON.parse(activitiesJson);
    
    var list = [];
    // Reverse to show newest first
    for (var i = activities.length - 1; i >= 0; i--) {
      var act = activities[i];
      if (act.employeeId === employeeId && act.date === dateStr) {
        list.push(act);
      }
    }
    return list;
  } catch(e) {
    return [];
  }
}

/**
 * Deletes an employee's activity log entry
 */
function deleteEmployeeActivity(activityId, employeeId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var activitiesJson = props.getProperty("SYSTEM_ACTIVITIES") || "[]";
    var activities = JSON.parse(activitiesJson);
    
    var index = -1;
    for (var i = 0; i < activities.length; i++) {
      if (activities[i].activityId === activityId && activities[i].employeeId === employeeId) {
        index = i;
        break;
      }
    }
    
    if (index === -1) {
      return { success: false, message: "Activity not found." };
    }
    
    activities.splice(index, 1);
    props.setProperty("SYSTEM_ACTIVITIES", JSON.stringify(activities));
    return { success: true };
  } catch(e) {
    return { success: false, message: "Error deleting activity: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns all activities logged by employees on a specific date for Admin
 */
function getActivitiesByDate(dateStr) {
  try {
    if (!dateStr) {
      dateStr = formatDateStr(new Date());
    }
    var props = PropertiesService.getScriptProperties();
    var activitiesJson = props.getProperty("SYSTEM_ACTIVITIES") || "[]";
    var activities = JSON.parse(activitiesJson);
    
    var list = [];
    // Reverse to show newest first
    for (var i = activities.length - 1; i >= 0; i--) {
      var act = activities[i];
      if (act.date === dateStr) {
        list.push(act);
      }
    }
    return list;
  } catch(e) {
    return [];
  }
}

/**
 * Force deletes and re-seeds database with initial dummy datasets
 */
function forceResetAndSeedDatabase() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    var ss = getDbSpreadsheet();
    if (!ss) return { success: false, message: "Spreadsheet container not found." };
    
    var sheets = ["Employees", "Attendance", "Leaves", "Hospitals"];
    sheets.forEach(name => {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
        }
      }
    });
    
    // Clear script properties settings and activities fallback
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty("SYSTEM_SETTINGS");
    props.deleteProperty("SYSTEM_ACTIVITIES");
    
    // Run setupSheets to recreate headers and seed dummy data
    setupSheets();
    
    return { success: true, message: "Database reset complete. All default dummy data restored." };
  } catch (e) {
    return { success: false, message: "Error resetting database: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetches dashboard stats for Admin View
 */
function getAdminDashboardStats(hospitalId) {
  try {
    var ss = getDbSpreadsheet();
    var empSheet = ss ? ss.getSheetByName("Employees") : null;
    var attSheet = ss ? ss.getSheetByName("Attendance") : null;
    var leaveSheet = ss ? ss.getSheetByName("Leaves") : null;
    var settingsSheet = ss ? ss.getSheetByName("Settings") : null;
    
    var totalStaff = 0;
    var presentToday = 0;
    var lateToday = 0;
    var absentToday = 0;
    var leavesPending = 0;
    
    var todayStr = formatDateStr(new Date());
    var staffStatusList = [];
    
    // 1. Fetch active non-admin employees
    var empList = [];
    if (empSheet) {
      var empData = empSheet.getDataRange().getValues();
      for (var i = 1; i < empData.length; i++) {
        if (empData[i][4] !== "Admin" && empData[i][8] === "Active") {
          var assignedHosp = empData[i][9] || "";
          if (hospitalId && assignedHosp !== hospitalId) {
            continue;
          }
          empList.push({
            employeeId: empData[i][0],
            name: empData[i][1],
            department: empData[i][5],
            designation: empData[i][6]
          });
          totalStaff++;
        }
      }
    }
    
    // 2. Fetch today's attendance logs
    var attMap = {};
    if (attSheet) {
      var attData = attSheet.getDataRange().getValues();
      for (var j = 1; j < attData.length; j++) {
        if (cellValueToDateStr(attData[j][3]) === todayStr) {
          var empId = attData[j][1];
          var status = attData[j][6];
          
          var isStaffInHospital = empList.some(function(e) { return e.employeeId === empId; });
          if (isStaffInHospital) {
            attMap[empId] = {
              checkIn: attData[j][4] ? (attData[j][4] instanceof Date ? formatTimeStr(attData[j][4]) : String(attData[j][4])) : "",
              checkOut: attData[j][5] ? (attData[j][5] instanceof Date ? formatTimeStr(attData[j][5]) : String(attData[j][5])) : "",
              status: status,
              remarks: attData[j][7]
            };
            
            if (status.indexOf("Present") !== -1 || status === "Late") presentToday++;
            else if (status === "Absent") absentToday++;
          }
        }
      }
    }
    
    // 3. Compile the presence list for all active non-admin employees
    empList.forEach(emp => {
      var log = attMap[emp.employeeId];
      if (log) {
        staffStatusList.push({
          employeeId: emp.employeeId,
          name: emp.name,
          department: emp.department,
          checkIn: log.checkIn || "-",
          checkOut: log.checkOut || "-",
          status: log.status
        });
      } else {
        // Default to Absent if no log exists
        staffStatusList.push({
          employeeId: emp.employeeId,
          name: emp.name,
          department: emp.department,
          checkIn: "-",
          checkOut: "-",
          status: "Absent"
        });
        absentToday++;
      }
    });
    
    // 4. Pending leaves
    if (leaveSheet) {
      var leaveData = leaveSheet.getDataRange().getValues();
      for (var k = 1; k < leaveData.length; k++) {
        if (leaveData[k][8] === "Pending") {
          var empId = leaveData[k][1];
          var isStaffInHospital = empList.some(function(e) { return e.employeeId === empId; });
          if (isStaffInHospital) {
            leavesPending++;
          }
        }
      }
    }
    
    // 5. Settings
    var officeStart = "09:00 AM";
    var lateMark = "09:15 AM";
    var workingDays = "Mon-Sat";
    var hospitalLatitude = 28.6139;
    var hospitalLongitude = 77.2090;
    if (settingsSheet && settingsSheet.getLastRow() > 1) {
      var settingsRows = settingsSheet.getDataRange().getValues();
      var headers = settingsRows[0];
      if (settingsRows.length > 1) {
        var settings = settingsRows[1];
        var rawStart = settings[0];
        var rawLate = settings[1];
        
        if (rawStart instanceof Date) {
          officeStart = formatTimeStr(rawStart);
        } else if (rawStart) {
          officeStart = String(rawStart);
        }
        
        if (rawLate instanceof Date) {
          lateMark = formatTimeStr(rawLate);
        } else if (rawLate) {
          lateMark = String(rawLate);
        }
        
        if (settings[2]) {
          workingDays = String(settings[2]);
        }

        var idxLat = headers.indexOf("HospitalLatitude");
        if (idxLat !== -1 && settings[idxLat] !== undefined && settings[idxLat] !== "") {
          hospitalLatitude = parseFloat(settings[idxLat]);
        }

        var idxLng = headers.indexOf("HospitalLongitude");
        if (idxLng !== -1 && settings[idxLng] !== undefined && settings[idxLng] !== "") {
          hospitalLongitude = parseFloat(settings[idxLng]);
        }
      }
    }
    
    return {
      totalStaff: totalStaff,
      presentToday: presentToday + lateToday,
      lateToday: lateToday,
      absentToday: absentToday,
      leavesPending: leavesPending,
      officeStart: officeStart,
      lateMark: lateMark,
      workingDays: workingDays,
      hospitalLatitude: hospitalLatitude,
      hospitalLongitude: hospitalLongitude,
      staffStatusList: staffStatusList
    };
  } catch (e) {
    Logger.log("Error in getAdminDashboardStats: " + e.message);
    return {
      totalStaff: 0,
      presentToday: 0,
      lateToday: 0,
      absentToday: 0,
      leavesPending: 0,
      officeStart: "09:00 AM",
      lateMark: "09:15 AM",
      workingDays: "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday",
      staffStatusList: []
    };
  }
}

/**
 * Edit an existing employee's details (Name, Email, Password, Role, Department, Designation)
 */
function editEmployee(employeeId, details) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    if (!empSheet) return { success: false, message: "Database system error. Employees sheet missing." };
    
    var data = empSheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        rowIndex = i + 1; // 1-indexed conversion
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Employee record not found." };
    }
    
    // Check if email is already taken by another employee
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] !== employeeId && data[j][2].toLowerCase() === details.email.toLowerCase()) {
        return { success: false, message: "Email is already registered under another account." };
      }
    }
    
    // Update values: Name, Email, Password (if provided), Role, Department, Designation, HospitalID, ShiftTime
    empSheet.getRange(rowIndex, 2).setValue(details.name);
    empSheet.getRange(rowIndex, 3).setValue(details.email);
    if (details.password) {
      empSheet.getRange(rowIndex, 4).setValue(details.password);
    }
    empSheet.getRange(rowIndex, 5).setValue(details.role);
    empSheet.getRange(rowIndex, 6).setValue("");
    empSheet.getRange(rowIndex, 7).setValue(details.designation);
    empSheet.getRange(rowIndex, 10).setValue(details.hospitalId || "");
    
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error updating employee: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete an employee from roster
 */
function deleteEmployee(employeeId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var empSheet = ss.getSheetByName("Employees");
    if (!empSheet) return { success: false, message: "Database system error. Employees sheet missing." };
    
    var data = empSheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Employee record not found." };
    }
    
    empSheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error deleting employee: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetch system settings details (including coordinates)
 */
function getSystemSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    var settingsJson = props.getProperty("SYSTEM_SETTINGS");
    var officeStartTime = "09:00 AM";
    var lateMarkAfter = "09:15 AM";
    var workingDays = "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday";
    var hospitalLatitude = 28.6139;
    var hospitalLongitude = 77.2090;
    
    if (settingsJson) {
      var s = JSON.parse(settingsJson);
      if (s.officeStartTime) officeStartTime = s.officeStartTime;
      if (s.lateMarkAfter) lateMarkAfter = s.lateMarkAfter;
      if (s.workingDays) workingDays = s.workingDays;
      if (s.hospitalLatitude !== undefined) hospitalLatitude = parseFloat(s.hospitalLatitude);
      if (s.hospitalLongitude !== undefined) hospitalLongitude = parseFloat(s.hospitalLongitude);
    }
    
    return {
      officeStartTime: officeStartTime,
      lateMarkAfter: lateMarkAfter,
      workingDays: workingDays,
      hospitalLatitude: hospitalLatitude,
      hospitalLongitude: hospitalLongitude
    };
  } catch (e) {
    return { 
      officeStartTime: "09:00 AM", 
      lateMarkAfter: "09:15 AM", 
      workingDays: "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday",
      hospitalLatitude: 28.6139,
      hospitalLongitude: 77.2090
    };
  }
}

/**
 * Save/Update system settings (including coordinates)
 */
function updateSystemSettings(startTime, lateTime, workingDays, hospitalLat, hospitalLng) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    props.setProperty("SYSTEM_SETTINGS", JSON.stringify({
      officeStartTime: startTime,
      lateMarkAfter: lateTime,
      workingDays: workingDays,
      hospitalLatitude: parseFloat(hospitalLat),
      hospitalLongitude: parseFloat(hospitalLng)
    }));
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error updating settings: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add or Update an attendance log (Admin CRUD action)
 */
function addOrUpdateAttendance(record) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var attSheet = ss.getSheetByName("Attendance");
    var empSheet = ss.getSheetByName("Employees");
    if (!attSheet || !empSheet) return { success: false, message: "Database system error. Sheets missing." };
    
    // Find Employee Name if adding new
    var empName = "";
    var empData = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === record.employeeId) {
        empName = empData[i][1];
        break;
      }
    }
    if (!empName) return { success: false, message: "Employee not found." };
    
    var dateParts = record.date.split('-'); // Format expected: dd-MM-yyyy (from input filter conversion)
    var dateStr = record.date;
    
    if (record.attendanceId) {
      // UPDATE existing attendance log
      var attData = attSheet.getDataRange().getValues();
      var rowIndex = -1;
      for (var j = 1; j < attData.length; j++) {
        if (attData[j][0] === record.attendanceId) {
          rowIndex = j + 1;
          break;
        }
      }
      
      if (rowIndex === -1) return { success: false, message: "Attendance record not found." };
      
      attSheet.getRange(rowIndex, 5).setValue(record.checkInTime); // CheckIn
      attSheet.getRange(rowIndex, 6).setValue(record.checkOutTime); // CheckOut
      attSheet.getRange(rowIndex, 7).setValue(record.status); // Status
      attSheet.getRange(rowIndex, 8).setValue(record.remarks); // Remarks
      
      return { success: true };
    } else {
      // CREATE new attendance log
      // Check for duplicate date log first
      var attData = attSheet.getDataRange().getValues();
      for (var k = 1; k < attData.length; k++) {
        if (attData[k][1] === record.employeeId && cellValueToDateStr(attData[k][3]) === dateStr) {
          return { success: false, message: "Employee already has an attendance log for " + dateStr };
        }
      }
      
      // Auto-generate Attendance ID: ATT + yyyyMMdd + EmployeeID
      var cleanDateId = dateStr.replace(/-/g, ""); // yyyyMMdd-like or ddMMMM-like
      var attId = "ATT" + cleanDateId.split('').reverse().join('') + record.employeeId; // simple unique hash
      if (dateParts.length === 3) {
        attId = "ATT" + dateParts[2] + dateParts[1] + dateParts[0] + record.employeeId;
      }
      
      attSheet.appendRow([
        attId,
        record.employeeId,
        empName,
        "'" + dateStr,
        record.checkInTime,
        record.checkOutTime,
        record.status,
        record.remarks,
        "" // No Selfie URL for manual admin override
      ]);
      
      return { success: true };
    }
  } catch (e) {
    return { success: false, message: "Error updating attendance: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete attendance log entry
 */
function deleteAttendance(attendanceId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var attSheet = ss.getSheetByName("Attendance");
    if (!attSheet) return { success: false, message: "Database system error. Attendance sheet missing." };
    
    var data = attSheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === attendanceId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Attendance record not found." };
    }
    
    attSheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error deleting attendance: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Update details of an existing employee activity
 */
function updateEmployeeActivity(activityId, employeeId, details, duration, category) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var props = PropertiesService.getScriptProperties();
    var activitiesJson = props.getProperty("SYSTEM_ACTIVITIES") || "[]";
    var activities = JSON.parse(activitiesJson);
    
    var index = -1;
    for (var i = 0; i < activities.length; i++) {
      if (activities[i].activityId === activityId && activities[i].employeeId === employeeId) {
        index = i;
        break;
      }
    }
    
    if (index === -1) {
      return { success: false, message: "Activity not found." };
    }
    
    activities[index].details = details;
    activities[index].duration = duration;
    activities[index].category = category;
    
    props.setProperty("SYSTEM_ACTIVITIES", JSON.stringify(activities));
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error updating activity: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Refresh user session data in real-time from Google Sheet
 */
function refreshUserSession(employeeId) {
  var sheet = getDbSpreadsheet().getSheetByName("Employees");
  if (!sheet) return { success: false, message: "Database system error." };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === employeeId) {
      if (row[8] === "Active") {
        return {
          success: true,
          user: {
            employeeId: row[0],
            name: row[1],
            email: row[2],
            role: row[4],
            department: row[5],
            designation: row[6],
            hospitalId: row[9] || "",
            shiftTime: row[10] || ""
          }
        };
      } else {
        return { success: false, message: "Account has been deactivated." };
      }
    }
  }
  return { success: false, message: "Employee not found." };
}

/**
 * Calculates distance between two coordinates in meters using the Haversine formula
 */
function getHaversineDistance(lat1, lon1, lat2, lon2) {
  var R = 6371000; // Radius of the Earth in meters
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = 
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
    Math.sin(dLon/2) * Math.sin(dLon/2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  var d = R * c; // Distance in meters
  return d;
}

/**
 * Retrieve list of all registered hospitals
 */
function getAllHospitals(hospitalId) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName("Hospitals");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    var hospId = data[i][0];
    if (hospitalId && hospId !== hospitalId) continue;
    list.push({
      hospitalId: hospId,
      name: data[i][1],
      latitude: parseFloat(data[i][2]),
      longitude: parseFloat(data[i][3]),
      range: parseInt(data[i][4], 10)
    });
  }
  return list;
}

/**
 * Add or update a hospital record
 */
function addOrUpdateHospital(record) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Hospitals");
    if (!sheet) return { success: false, message: "Database system error. Hospitals sheet missing." };
    
    var data = sheet.getDataRange().getValues();
    
    if (record.hospitalId) {
      // UPDATE existing hospital
      var rowIndex = -1;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === record.hospitalId) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex === -1) return { success: false, message: "Hospital record not found." };
      
      sheet.getRange(rowIndex, 2).setValue(record.name);
      sheet.getRange(rowIndex, 3).setValue(parseFloat(record.latitude));
      sheet.getRange(rowIndex, 4).setValue(parseFloat(record.longitude));
      sheet.getRange(rowIndex, 5).setValue(parseInt(record.range, 10));
      
      return { success: true };
    } else {
      // CREATE new hospital
      // Generate ID: HOSP + serial number
      var maxIdVal = 0;
      for (var j = 1; j < data.length; j++) {
        var idNum = parseInt(data[j][0].replace("HOSP", ""), 10);
        if (!isNaN(idNum) && idNum > maxIdVal) {
          maxIdVal = idNum;
        }
      }
      var nextId = "HOSP" + String(maxIdVal + 1).padStart(3, '0');
      
      sheet.appendRow([
        nextId,
        record.name,
        parseFloat(record.latitude),
        parseFloat(record.longitude),
        parseInt(record.range, 10)
      ]);
      
      return { success: true, hospitalId: nextId };
    }
  } catch (e) {
    return { success: false, message: "Error saving hospital: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete a hospital record
 */
function deleteHospital(hospitalId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName("Hospitals");
    if (!sheet) return { success: false, message: "Database system error. Hospitals sheet missing." };
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === hospitalId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Hospital record not found." };
    }
    
    // Check if any active employee is assigned to this hospital
    var empSheet = ss.getSheetByName("Employees");
    if (empSheet) {
      var empData = empSheet.getDataRange().getValues();
      for (var k = 1; k < empData.length; k++) {
        if (empData[k][9] === hospitalId && empData[k][8] === "Active") {
          return { success: false, message: "Cannot delete hospital. Active staff member(s) are assigned to it." };
        }
      }
    }
    
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, message: "Error deleting hospital: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get system activity logs for Admin (Global Attendance History)
 */
function getSystemActivityLogs() {
  var ss = getDbSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance");
  if (!attSheet) return [];
  var data = attSheet.getDataRange().getValues();
  
  var logs = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var attId = row[0];
    var empId = row[1];
    var empName = row[2];
    var dateStr = cellValueToDateStr(row[3]);
    var checkIn = row[4] ? (row[4] instanceof Date ? formatTimeStr(row[4]) : String(row[4])) : "";
    var checkOut = row[5] ? (row[5] instanceof Date ? formatTimeStr(row[5]) : String(row[5])) : "";
    
    if (checkIn) {
      logs.push({
        activityId: attId + "_IN",
        employeeId: empId,
        employeeName: empName,
        category: "Check-In",
        details: "Checked in.",
        date: dateStr,
        time: checkIn,
        duration: 0
      });
    }
    if (checkOut) {
      logs.push({
        activityId: attId + "_OUT",
        employeeId: empId,
        employeeName: empName,
        category: "Check-Out",
        details: "Checked out.",
        date: dateStr,
        time: checkOut,
        duration: 0
      });
    }
  }
  
  return logs.reverse(); // Newest first
}

/**
 * Get activity logs for a specific Employee
 */
function getEmployeeActivityLogs(employeeId) {
  var ss = getDbSpreadsheet();
  var attSheet = ss.getSheetByName("Attendance");
  if (!attSheet) return [];
  var data = attSheet.getDataRange().getValues();
  
  var logs = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] !== employeeId) continue;
    
    var attId = row[0];
    var dateStr = cellValueToDateStr(row[3]);
    var checkIn = row[4] ? (row[4] instanceof Date ? formatTimeStr(row[4]) : String(row[4])) : "";
    var checkOut = row[5] ? (row[5] instanceof Date ? formatTimeStr(row[5]) : String(row[5])) : "";
    
    if (checkIn) {
      logs.push({
        activityId: attId + "_IN",
        employeeId: employeeId,
        employeeName: row[2],
        category: "Check-In",
        details: "Checked in successfully.",
        date: dateStr,
        time: checkIn,
        duration: 0
      });
    }
    if (checkOut) {
      logs.push({
        activityId: attId + "_OUT",
        employeeId: employeeId,
        employeeName: row[2],
        category: "Check-Out",
        details: "Checked out successfully.",
        date: dateStr,
        time: checkOut,
        duration: 0
      });
    }
  }
  
  return logs.reverse(); // Newest first
}
