/**
 * Create a normalized **name key** for a person, robust to:
 * - middle names
 * - swapped first/last name (Japanese style)
 *
 * Steps:
 * 1) Take FIRST WORD of lastName and firstName.
 * 2) Keep only letters, lowercase.
 * 3) Sort the two tokens alphabetically and join with "|".
 */
function createNameKeyForArchived_(lastName, firstName) {
  function firstWordLettersOnly(s) {
    if (!s) return '';
    var w = String(s).toLowerCase().trim().split(/\s+/)[0];
    return w.replace(/[^a-z]/g, '');
  }

  var t1 = firstWordLettersOnly(lastName);
  var t2 = firstWordLettersOnly(firstName);
  if (!t1 || !t2) return '';

  var parts = [t1, t2].sort();
  return parts[0] + '|' + parts[1];
}

/**
 * FAST delete:
 * - Reads all names in Archived (from external Directory file in Config!B2).
 * - Deletes matching rows (by normalized name) in:
 *   - Appsheet Sunserv     (B last, C first, data from row 2)
 *   - Appsheet Event       (B last, C first, data from row 2)
 *   - Sunday Service       (C last, D first, data from row 4)
 *   - Event Attendance     (C last, D first, data from row 5)
 *   - Appsheet Pastoral    (B last, C first, data from row 2)
 *   - Pastoral Check-In    (C last, D first, data from row 5)
 *
 * Run this from the ATTENDANCE spreadsheet.
 */
function deleteArchivedFromAttendanceTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    throw new Error('Config sheet not found in this spreadsheet.');
  }

  Logger.log("=== START DELETE ARCHIVED FROM ATTENDANCE ===");

  // Config!B2 = URL/ID of the DIRECTORY file that contains "Archived"
  var dirRef = configSheet.getRange('B2').getValue();
  if (!dirRef) {
    Logger.log("Config!B2 is empty. Cannot proceed.");
    return;
  }

  var dirUrl = String(dirRef);
  if (!dirUrl.startsWith('http')) {
    dirUrl = 'https://docs.google.com/spreadsheets/d/' + dirUrl + '/edit';
  }

  Logger.log("Opening Directory file: " + dirUrl);

  var directorySs = SpreadsheetApp.openByUrl(dirUrl);
  var archivedSheet = directorySs.getSheetByName('Archived');
  if (!archivedSheet) {
    Logger.log("Archived sheet NOT FOUND in the Directory file.");
    return;
  }

  var ARCH_START_ROW = 4;
  var COL_LAST_NAME = 2;  // C
  var COL_FIRST_NAME = 3; // D

  var lastRowArch = archivedSheet.getLastRow();
  if (lastRowArch < ARCH_START_ROW) {
    Logger.log("No archived rows found. Nothing to delete.");
    return;
  }

    var archValues = archivedSheet
    .getRange(ARCH_START_ROW, 1, lastRowArch - ARCH_START_ROW + 1, archivedSheet.getLastColumn())
    .getValues();

  var COL_FLAG = 0; // Column A = status/flag

  var archivedNameSet = {};
  for (var i = 0; i < archValues.length; i++) {
    var row = archValues[i];

    var flag = (row[COL_FLAG] || '').toString().trim().toUpperCase();

    // Decide if this archived row should be used to delete from attendance tabs
    var includeForDeletion = false;

    if (!flag) {
      // No status written yet → treat as normal archived → delete from attendance
      includeForDeletion = true;
    } else if (flag === 'PERMANENTLY ARCHIVED BECAUSE ASCENDED') {
      // Ascended → permanently archived → delete from attendance
      includeForDeletion = true;
    } else if (
      flag === 'RETURNED TO DIRECTORY BECAUSE ACTIVE AGAIN OR NEW MEMBER' ||
      flag === 'DUPLICATE ARCHIVED RECORD - IGNORE THIS ROW, KEEP THE FIRST ONE' ||
      flag === 'ALREADY IN DIRECTORY - NO ACTION NEEDED'
    ) {
      // These should NOT cause deletion from attendance → skip
      includeForDeletion = false;
    }

    if (includeForDeletion) {
      var key = createNameKeyForArchived_(row[COL_LAST_NAME], row[COL_FIRST_NAME]);
      if (key) archivedNameSet[key] = true;
    }
  }


  var totalDeleted = 0;

  // NEW helper: actually delete rows (like right-click → Delete row)
  function deleteRowsByNameKey_(sheet, startRow, lastNameCol, firstNameCol, sheetName) {
    if (!sheet) {
      Logger.log(sheetName + " TAB NOT FOUND - skipped.");
      return 0;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < startRow) {
      Logger.log(sheetName + " has no data rows.");
      return 0;
    }

    var lastCol = sheet.getLastColumn();
    var numRows = lastRow - startRow + 1;
    var data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

    var lnIndex = lastNameCol - 1;   // because range starts at column 1
    var fnIndex = firstNameCol - 1;

    var rowsToDelete = [];

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var key = createNameKeyForArchived_(row[lnIndex], row[fnIndex]);

      if (key && archivedNameSet[key]) {
        rowsToDelete.push(startRow + r);  // store absolute row number
      }
    }

    // delete from bottom to top so row numbers stay valid
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }

    Logger.log(sheetName + " - deleted: " + rowsToDelete.length);
    return rowsToDelete.length;
  }

  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Appsheet Sunserv'), 2, 2, 3, "Appsheet Sunserv");
  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Appsheet Event'), 2, 2, 3, "Appsheet Event");
  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Sunday Service'), 4, 3, 4, "Sunday Service");
  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Event Attendance'), 5, 3, 4, "Event Attendance");
  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Appsheet Pastoral'), 2, 2, 3, "Appsheet Pastoral");
  totalDeleted += deleteRowsByNameKey_(ss.getSheetByName('Pastoral Check-In'), 5, 3, 4, "Pastoral Check-In");

  Logger.log("=== SUMMARY ===");
  Logger.log("Total rows deleted: " + totalDeleted);
  Logger.log("=== FINISHED ===");
}
