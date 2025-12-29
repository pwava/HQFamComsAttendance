/**
 * Syncs names from external Directory into 6 tabs:
 * - Event Attendance
 * - Sunday Service
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 * - Pastoral Check-In
 *
 * Directory is in an external spreadsheet whose URL/ID is in Config!B2,
 * sheet name "Directory", with:
 *   - Col Z: Personal ID
 *   - Col C: Last Name
 *   - Col D: First Name
 *   - Col G: Birthdate
 *   - Headers on rows 1–3
 *
 * The tabs have:
 *   - Event Attendance: Col B personal ID, Col C last name, Col D first name, rows 1–4 headers
 *   - Sunday Service:   Col B personal ID, Col C last name, Col D first name, rows 1–3 headers
 *   - Pastoral Check-In: Col B personal ID, Col C last name, Col D first name, rows 1–3 headers
 *   - Appsheet Sunserv: Col A personal ID, Col B last name, Col C first name, row 1 header
 *   - Appsheet Event:   Col A personal ID, Col B last name, Col C first name, row 1 header
 *   - Appsheet Pastoral: Col A personal ID, Col B last name, Col C first name, row 1 header
 *
 * RULES:
 * - Do NOT copy Directory rows if Directory Personal ID (Col Z) is blank.
 * - Do NOT use Attendance/Appsheet rows WITHOUT Personal ID as a source to copy into other tabs.
 * - Directory uniqueness key includes: ID + Last + First + Birthdate (Directory Col G),
 *   but syncing into tabs uses ID + Last + First (because tabs do not store birthdate).
 */
function syncDirectoryNamesToAllTabs() {
  var CONFIG_SHEET_NAME = 'Config';
  var DIRECTORY_SHEET_NAME = 'Directory';

  var DIRECTORY_ID_COL = 26;         // Z  <<< CHANGED
  var DIRECTORY_LAST_NAME_COL = 3;   // C
  var DIRECTORY_FIRST_NAME_COL = 4;  // D
  var DIRECTORY_BIRTHDATE_COL = 7;   // G
  var DIRECTORY_HEADER_ROWS = 3;     // 1–3 are headers

  var SHEETS_CONFIG = [
    { name: 'Event Attendance', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 4 },
    { name: 'Sunday Service', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { name: 'Appsheet Sunserv', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Event', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Pastoral', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Pastoral Check-In', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    throw new Error("Config sheet '" + CONFIG_SHEET_NAME + "' not found.");
  }

  // --- Get external Directory spreadsheet ID from Config!B2 ---
  var externalRef = configSheet.getRange('B2').getValue();
  if (!externalRef) {
    throw new Error("Config!B2 is empty. Please put the Directory spreadsheet URL or ID there.");
  }
  var externalId = extractSpreadsheetIdFromString_(String(externalRef));

  var externalSs = SpreadsheetApp.openById(externalId);
  var directorySheet = externalSs.getSheetByName(DIRECTORY_SHEET_NAME);
  if (!directorySheet) {
    throw new Error("Directory sheet '" + DIRECTORY_SHEET_NAME + "' not found in external spreadsheet.");
  }

  var lastDirRow = directorySheet.getLastRow();
  var dirNumRows = Math.max(0, lastDirRow - DIRECTORY_HEADER_ROWS);

  var dirIds = [];
  var dirLastNames = [];
  var dirFirstNames = [];
  var dirBirthdates = [];

  if (dirNumRows > 0) {
    dirIds = directorySheet.getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_ID_COL, dirNumRows, 1).getValues();
    dirLastNames = directorySheet.getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_LAST_NAME_COL, dirNumRows, 1).getValues();
    dirFirstNames = directorySheet.getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_FIRST_NAME_COL, dirNumRows, 1).getValues();
    dirBirthdates = directorySheet.getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_BIRTHDATE_COL, dirNumRows, 1).getValues();
  }

  // Build directory entries:
  // - keyFull = ID + Last + First + Birthdate (Directory-only uniqueness)
  // - keyMatch = ID + Last + First (used for syncing into tabs, since tabs do not store birthdate)
  var directoryEntries = [];
  var directoryMatchMap = {}; // keyMatch -> entry (first wins)

  for (var i = 0; i < dirNumRows; i++) {
    var personalId = (dirIds[i][0] || '').toString().trim();
    if (!personalId) continue; // IMPORTANT: skip if Directory ID is blank

    var lastName = (dirLastNames[i][0] || '').toString().trim();
    var firstName = (dirFirstNames[i][0] || '').toString().trim();
    var birthRaw = dirBirthdates[i][0];

    if (!lastName && !firstName) continue;

    var birthKey = normalizeBirthdateKey_(birthRaw); // YYYY-MM-DD or ''
    var keyFull = buildPersonKey_(personalId, lastName, firstName, birthKey);
    var keyMatch = buildPersonKey_(personalId, lastName, firstName, ''); // no birthdate for syncing

    if (!keyFull || !keyMatch) continue;

    var entry = {
      personalId: personalId,
      lastName: lastName,
      firstName: firstName,
      birthKey: birthKey,
      keyFull: keyFull,
      keyMatch: keyMatch
    };

    directoryEntries.push(entry);

    // for syncing into tabs (no birthdate), keep first occurrence per keyMatch
    if (!directoryMatchMap[keyMatch]) {
      directoryMatchMap[keyMatch] = entry;
    }
  }

  var directoryEntriesForSync = [];
  for (var k in directoryMatchMap) {
    if (Object.prototype.hasOwnProperty.call(directoryMatchMap, k)) {
      directoryEntriesForSync.push(directoryMatchMap[k]);
    }
  }

  // --- For each sheet, append any missing names from Directory (one-way) ---
  for (var s = 0; s < SHEETS_CONFIG.length; s++) {
    var cfg = SHEETS_CONFIG[s];
    var sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Sheet '" + cfg.name + "' not found. Skipping.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var dataStartRow = cfg.headerRows + 1;

    var existingKeys = {};

    if (lastRow >= dataStartRow) {
      var numRows = lastRow - cfg.headerRows;

      var idValues = sheet.getRange(dataStartRow, cfg.idCol, numRows, 1).getValues();
      var lastNameValues = sheet.getRange(dataStartRow, cfg.lastNameCol, numRows, 1).getValues();
      var firstNameValues = sheet.getRange(dataStartRow, cfg.firstNameCol, numRows, 1).getValues();

      for (var r = 0; r < numRows; r++) {
        var pid = (idValues[r][0] || '').toString().trim();
        if (!pid) continue; // IMPORTANT: if no Personal ID, do NOT use as key/source

        var ln = (lastNameValues[r][0] || '').toString().trim();
        var fn = (firstNameValues[r][0] || '').toString().trim();
        if (!ln && !fn) continue;

        var keyMatchExisting = buildPersonKey_(pid, ln, fn, '');
        if (keyMatchExisting) existingKeys[keyMatchExisting] = true;
      }
    }

    var rowsToAppend = [];
    for (var e = 0; e < directoryEntriesForSync.length; e++) {
      var entry2 = directoryEntriesForSync[e];
      if (!existingKeys[entry2.keyMatch]) {
        existingKeys[entry2.keyMatch] = true;

        var newRow = new Array(numCols);
        for (var c = 0; c < numCols; c++) newRow[c] = '';

        newRow[cfg.idCol - 1] = entry2.personalId;
        newRow[cfg.lastNameCol - 1] = entry2.lastName;
        newRow[cfg.firstNameCol - 1] = entry2.firstName;

        rowsToAppend.push(newRow);
      }
    }

    if (rowsToAppend.length > 0) {
      var appendStartRow = getNextAvailableRow_(sheet, dataStartRow, cfg.lastNameCol, cfg.firstNameCol);
      sheet.getRange(appendStartRow, 1, rowsToAppend.length, numCols).setValues(rowsToAppend);
      Logger.log("Sheet '" + sheet.getName() + "': appended " + rowsToAppend.length + " names from Directory.");
    } else {
      Logger.log("Sheet '" + sheet.getName() + "': no new names needed from Directory.");
    }
  }

  // --- Build union of names (Directory + tabs with Personal ID only) ---
  var unionEntries = buildUnionEntries_(ss, directoryEntriesForSync);

  // --- Ensure all union names are present in Attendance tabs ---
  syncUnionNamesIntoAttendanceTabs_(ss, unionEntries);

  // --- Ensure all union names are present in Appsheet tabs ---
  syncUnionNamesIntoAppsheetTabs_(ss, unionEntries);

  // --- After syncing, sort the 6 tabs by Attendance Stats status + name ---
  sortSyncedTabsByAttendanceStatus();
}

/**
 * Finds the next truly empty row (based on last/first name columns),
 * starting from dataStartRow.
 */
function getNextAvailableRow_(sheet, dataStartRow, lastNameCol, firstNameCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return dataStartRow;

  var numRows = lastRow - dataStartRow + 1;
  var lastNames = sheet.getRange(dataStartRow, lastNameCol, numRows, 1).getValues();
  var firstNames = sheet.getRange(dataStartRow, firstNameCol, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var ln = (lastNames[i][0] || '').toString().trim();
    var fn = (firstNames[i][0] || '').toString().trim();
    if (!ln && !fn) {
      return dataStartRow + i;
    }
  }
  return lastRow + 1;
}

/**
 * Normalizes a birthdate value into a stable key string: YYYY-MM-DD or ''.
 */
function normalizeBirthdateKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (!s) return '';
  return s;
}

/**
 * Builds a normalized person key from:
 * Personal ID + Last Name + First Name + BirthdateKey (optional).
 *
 * Keeps digits (A–Z, 0–9, accented letters).
 */
function buildPersonKey_(personalId, lastName, firstName, birthdateKey) {
  var pid = (personalId || '').toString().trim().toLowerCase();
  var ln = (lastName || '').toString().trim().toLowerCase();
  var fn = (firstName || '').toString().trim().toLowerCase();
  var bd = (birthdateKey || '').toString().trim().toLowerCase();

  if (!pid) return null;
  if (!ln && !fn) return null;

  var cleanPid = pid.replace(/[^A-Za-z0-9\u00C0-\u024F]/g, '');
  var cleanLn = ln.replace(/[^A-Za-z0-9\u00C0-\u024F]/g, '');
  var cleanFn = fn.replace(/[^A-Za-z0-9\u00C0-\u024F]/g, '');
  var cleanBd = bd.replace(/[^A-Za-z0-9\u00C0-\u024F-]/g, '');

  if (!cleanPid) return null;
  if (!cleanLn && !cleanFn) return null;

  return cleanPid + '|' + cleanLn + '|' + cleanFn + '|' + cleanBd;
}

/**
 * Extracts a spreadsheet ID from either a raw ID or a full URL string.
 */
function extractSpreadsheetIdFromString_(input) {
  if (/^[\w-]{25,}$/.test(input)) return input;

  var match = input.match(/[-\w]{25,}/);
  if (match && match[0]) return match[0];

  throw new Error('Could not extract Spreadsheet ID from: ' + input);
}

/**
 * Loose sheet matcher:
 * - tries exact getSheetByName first
 * - then matches by normalized name (trim, collapse spaces, lowercase)
 */
function getSheetByNameLoose_(ss, targetName) {
  var exact = ss.getSheetByName(targetName);
  if (exact) return exact;

  var normTarget = normalizeSheetName_(targetName);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (normalizeSheetName_(s.getName()) === normTarget) {
      return s;
    }
  }
  return null;
}

/**
 * Normalizes sheet names to avoid hidden-space/case issues.
 */
function normalizeSheetName_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Returns a numeric rank for a given status string.
 * Lower number = higher priority in sorting.
 */
function getStatusRank_(rawStatus) {
  var s = String(rawStatus || '').toLowerCase().trim();
  if (s === 'core') return 0;
  if (s === 'active') return 1;
  if (s === 'inactive') return 2;
  if (s === 'archived') return 3;
  return 4;
}

/**
 * Sorts the following tabs by status from 'Attendance Stats' and then by name:
 * - Event Attendance
 * - Sunday Service
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 * - Pastoral Check-In
 *
 * Attendance Stats:
 *   - Col C: Last Name
 *   - Col D: First Name
 *   - Col F: Status
 *   - Rows 1–2 headers, data from row 3
 *
 * NOTE: This sorting still uses Last/First only (as before).
 */
function sortSyncedTabsByAttendanceStatus() {
  var ATTENDANCE_STATS_SHEET_NAME = 'Attendance Stats';
  var ATT_LAST_NAME_COL = 3; // C
  var ATT_FIRST_NAME_COL = 4; // D
  var ATT_STATUS_COL = 6; // F
  var ATT_HEADER_ROWS = 2;

  var SORT_SHEETS_CONFIG = [
    { name: 'Event Attendance', lastNameCol: 3, firstNameCol: 4, headerRows: 4 },
    { name: 'Sunday Service', lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { name: 'Appsheet Sunserv', lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Event', lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Pastoral', lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Pastoral Check-In', lastNameCol: 3, firstNameCol: 4, headerRows: 3 }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName(ATTENDANCE_STATS_SHEET_NAME);
  if (!attSheet) {
    throw new Error("Attendance Stats sheet '" + ATTENDANCE_STATS_SHEET_NAME + "' not found.");
  }

  var lastRowStats = attSheet.getLastRow();
  if (lastRowStats <= ATT_HEADER_ROWS) {
    Logger.log('Attendance Stats has no data rows to use for sorting.');
    return;
  }

  var statsNumRows = lastRowStats - ATT_HEADER_ROWS;
  var statsLastNames = attSheet.getRange(ATT_HEADER_ROWS + 1, ATT_LAST_NAME_COL, statsNumRows, 1).getValues();
  var statsFirstNames = attSheet.getRange(ATT_HEADER_ROWS + 1, ATT_FIRST_NAME_COL, statsNumRows, 1).getValues();
  var statsStatuses = attSheet.getRange(ATT_HEADER_ROWS + 1, ATT_STATUS_COL, statsNumRows, 1).getValues();

  var statusMap = {};
  for (var i = 0; i < statsNumRows; i++) {
    var ln = (statsLastNames[i][0] || '').toString().trim();
    var fn = (statsFirstNames[i][0] || '').toString().trim();
    var status = statsStatuses[i][0];
    var key = (ln.toLowerCase() + '|' + fn.toLowerCase());
    statusMap[key] = getStatusRank_(status);
  }

  for (var j = 0; j < SORT_SHEETS_CONFIG.length; j++) {
    var cfg = SORT_SHEETS_CONFIG[j];
    var sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Sheet '" + cfg.name + "' not found. Skipping sort.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    var headerRows = cfg.headerRows;
    if (lastRow <= headerRows) {
      Logger.log("Sheet '" + cfg.name + "' has no data rows to sort.");
      continue;
    }

    var numRows = lastRow - headerRows;
    var numCols = sheet.getLastColumn();
    var dataRange = sheet.getRange(headerRows + 1, 1, numRows, numCols);
    var dataValues = dataRange.getValues();

    var rowsWithMeta = [];
    for (var x = 0; x < dataValues.length; x++) {
      var row = dataValues[x];
      var lnRaw = (row[cfg.lastNameCol - 1] || '').toString().trim();
      var fnRaw = (row[cfg.firstNameCol - 1] || '').toString().trim();
      var key2 = (lnRaw.toLowerCase() + '|' + fnRaw.toLowerCase());
      var rank = Object.prototype.hasOwnProperty.call(statusMap, key2) ? statusMap[key2] : 4;

      rowsWithMeta.push({
        row: row,
        rank: rank,
        ln: lnRaw.toLowerCase(),
        fn: fnRaw.toLowerCase(),
        originalIndex: x
      });
    }

    rowsWithMeta.sort(function (a, b) {
      if (a.rank !== b.rank) return a.rank - b.rank;
      if (a.ln !== b.ln) return a.ln.localeCompare(b.ln);
      if (a.fn !== b.fn) return a.fn.localeCompare(b.fn);
      return a.originalIndex - b.originalIndex;
    });

    var sortedValues = [];
    for (var y = 0; y < rowsWithMeta.length; y++) {
      sortedValues.push(rowsWithMeta[y].row);
    }

    dataRange.setValues(sortedValues);
    Logger.log("Sheet '" + cfg.name + "' sorted by status and name.");
  }
}

/**
 * Build a union of people from:
 * - Directory entries (already filtered: must have Personal ID)
 * - Attendance tabs + Appsheet tabs, BUT ONLY rows that have Personal ID
 */
function buildUnionEntries_(ss, directoryEntriesForSync) {
  var SOURCE_SHEETS = [
    // Attendance tabs
    { name: 'Event Attendance', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 4 },
    { name: 'Sunday Service', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { name: 'Pastoral Check-In', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },

    // Appsheet tabs
    { name: 'Appsheet Sunserv', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Event', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Pastoral', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 }
  ];

  var unionMap = {}; // keyMatch -> entry

  // Start with Directory entries (keyMatch)
  for (var i = 0; i < directoryEntriesForSync.length; i++) {
    var entry = directoryEntriesForSync[i];
    if (!unionMap[entry.keyMatch]) {
      unionMap[entry.keyMatch] = {
        personalId: entry.personalId,
        lastName: entry.lastName,
        firstName: entry.firstName,
        keyMatch: entry.keyMatch
      };
    }
  }

  // Add from sheets (ONLY if Personal ID present)
  for (var s = 0; s < SOURCE_SHEETS.length; s++) {
    var cfg = SOURCE_SHEETS[s];
    var sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Union source sheet '" + cfg.name + "' not found. Skipping.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    var dataStartRow = cfg.headerRows + 1;
    if (lastRow < dataStartRow) continue;

    var numRows = lastRow - cfg.headerRows;

    var ids = sheet.getRange(dataStartRow, cfg.idCol, numRows, 1).getValues();
    var lastNames = sheet.getRange(dataStartRow, cfg.lastNameCol, numRows, 1).getValues();
    var firstNames = sheet.getRange(dataStartRow, cfg.firstNameCol, numRows, 1).getValues();

    for (var i2 = 0; i2 < numRows; i2++) {
      var pid = (ids[i2][0] || '').toString().trim();
      if (!pid) continue; // IMPORTANT: do not copy rows without Personal ID

      var ln = (lastNames[i2][0] || '').toString().trim();
      var fn = (firstNames[i2][0] || '').toString().trim();
      if (!ln && !fn) continue;

      var keyMatch = buildPersonKey_(pid, ln, fn, '');
      if (!keyMatch) continue;

      if (!unionMap[keyMatch]) {
        unionMap[keyMatch] = {
          personalId: pid,
          lastName: ln,
          firstName: fn,
          keyMatch: keyMatch
        };
      }
    }
  }

  var out = [];
  for (var k in unionMap) {
    if (Object.prototype.hasOwnProperty.call(unionMap, k)) {
      out.push(unionMap[k]);
    }
  }
  return out;
}

/**
 * Ensure all union people exist in Attendance tabs:
 * - Event Attendance
 * - Sunday Service
 * - Pastoral Check-In
 */
function syncUnionNamesIntoAttendanceTabs_(ss, unionEntries) {
  var ATTENDANCE_TABS_CONFIG = [
    { name: 'Event Attendance', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 4 },
    { name: 'Sunday Service', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { name: 'Pastoral Check-In', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 }
  ];

  for (var t = 0; t < ATTENDANCE_TABS_CONFIG.length; t++) {
    var cfg = ATTENDANCE_TABS_CONFIG[t];
    var sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Attendance tab '" + cfg.name + "' not found. Skipping union sync.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    var dataStartRow = cfg.headerRows + 1;
    var numCols = sheet.getLastColumn();
    var existingKeys = {};

    if (lastRow >= dataStartRow) {
      var numRows = lastRow - cfg.headerRows;

      var idValues = sheet.getRange(dataStartRow, cfg.idCol, numRows, 1).getValues();
      var lastNameValues = sheet.getRange(dataStartRow, cfg.lastNameCol, numRows, 1).getValues();
      var firstNameValues = sheet.getRange(dataStartRow, cfg.firstNameCol, numRows, 1).getValues();

      for (var r = 0; r < numRows; r++) {
        var pid = (idValues[r][0] || '').toString().trim();
        if (!pid) continue; // IMPORTANT: ignore rows without ID

        var ln = (lastNameValues[r][0] || '').toString().trim();
        var fn = (firstNameValues[r][0] || '').toString().trim();
        if (!ln && !fn) continue;

        var keyMatch = buildPersonKey_(pid, ln, fn, '');
        if (keyMatch) existingKeys[keyMatch] = true;
      }
    }

    var rowsToAppend = [];
    for (var i = 0; i < unionEntries.length; i++) {
      var entry = unionEntries[i];
      if (!entry.personalId) continue;

      if (!existingKeys[entry.keyMatch]) {
        existingKeys[entry.keyMatch] = true;

        var newRow = new Array(numCols);
        for (var c = 0; c < numCols; c++) newRow[c] = '';

        newRow[cfg.idCol - 1] = entry.personalId;
        newRow[cfg.lastNameCol - 1] = entry.lastName;
        newRow[cfg.firstNameCol - 1] = entry.firstName;

        rowsToAppend.push(newRow);
      }
    }

    if (rowsToAppend.length > 0) {
      var appendStartRow = getNextAvailableRow_(sheet, dataStartRow, cfg.lastNameCol, cfg.firstNameCol);
      sheet.getRange(appendStartRow, 1, rowsToAppend.length, numCols).setValues(rowsToAppend);
      Logger.log("Attendance tab '" + sheet.getName() + "': appended " + rowsToAppend.length + " union names.");
    } else {
      Logger.log("Attendance tab '" + cfg.name + "': no union names needed.");
    }
  }
}

/**
 * Ensure all union people exist in Appsheet tabs:
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 */
function syncUnionNamesIntoAppsheetTabs_(ss, unionEntries) {
  var APPSHEET_TABS_CONFIG = [
    { name: 'Appsheet Sunserv', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Event', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { name: 'Appsheet Pastoral', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 }
  ];

  for (var t = 0; t < APPSHEET_TABS_CONFIG.length; t++) {
    var cfg = APPSHEET_TABS_CONFIG[t];
    var sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Appsheet tab '" + cfg.name + "' not found. Skipping union sync.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    var dataStartRow = cfg.headerRows + 1;
    var numCols = sheet.getLastColumn();
    var existingKeys = {};

    if (lastRow >= dataStartRow) {
      var numRows = lastRow - cfg.headerRows;

      var idValues = sheet.getRange(dataStartRow, cfg.idCol, numRows, 1).getValues();
      var lastNameValues = sheet.getRange(dataStartRow, cfg.lastNameCol, numRows, 1).getValues();
      var firstNameValues = sheet.getRange(dataStartRow, cfg.firstNameCol, numRows, 1).getValues();

      for (var r = 0; r < numRows; r++) {
        var pid = (idValues[r][0] || '').toString().trim();
        if (!pid) continue; // IMPORTANT: ignore rows without ID for copy logic

        var ln = (lastNameValues[r][0] || '').toString().trim();
        var fn = (firstNameValues[r][0] || '').toString().trim();
        if (!ln && !fn) continue;

        var keyMatch = buildPersonKey_(pid, ln, fn, '');
        if (keyMatch) existingKeys[keyMatch] = true;
      }
    }

    var rowsToAppend = [];
    for (var i = 0; i < unionEntries.length; i++) {
      var entry = unionEntries[i];
      if (!entry.personalId) continue;

      if (!existingKeys[entry.keyMatch]) {
        existingKeys[entry.keyMatch] = true;

        var newRow = new Array(numCols);
        for (var c = 0; c < numCols; c++) newRow[c] = '';

        newRow[cfg.idCol - 1] = entry.personalId;
        newRow[cfg.lastNameCol - 1] = entry.lastName;
        newRow[cfg.firstNameCol - 1] = entry.firstName;

        rowsToAppend.push(newRow);
      }
    }

    if (rowsToAppend.length > 0) {
      var appendStartRow = getNextAvailableRow_(sheet, dataStartRow, cfg.lastNameCol, cfg.firstNameCol);
      sheet.getRange(appendStartRow, 1, rowsToAppend.length, numCols).setValues(rowsToAppend);
      Logger.log("Appsheet tab '" + sheet.getName() + "': appended " + rowsToAppend.length + " union names.");
    } else {
      Logger.log("Appsheet tab '" + cfg.name + "': no union names needed.");
    }
  }
}

/**
 * Debug helper:
 * - Counts how many unique (ID+Last+First) each tab has.
 * - Shows how many names in Appsheet tabs are NOT in Directory (by ID+Last+First).
 *
 * IMPORTANT: Rows without Personal ID are ignored in comparisons.
 */
function checkNameCountsAndExtras() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var CONFIG_SHEET_NAME = 'Config';
  var DIRECTORY_SHEET_NAME = 'Directory';

  var tabConfigs = [
    { id: 'directory', label: 'Directory', sheetName: DIRECTORY_SHEET_NAME, idCol: 26, lastNameCol: 3, firstNameCol: 4, headerRows: 3, external: true }, // Z  <<< CHANGED
    { id: 'event', label: 'Event Attendance', sheetName: 'Event Attendance', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 4 },
    { id: 'sunday', label: 'Sunday Service', sheetName: 'Sunday Service', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { id: 'pastoralCheck', label: 'Pastoral Check-In', sheetName: 'Pastoral Check-In', idCol: 2, lastNameCol: 3, firstNameCol: 4, headerRows: 3 },
    { id: 'appSunserv', label: 'Appsheet Sunserv', sheetName: 'Appsheet Sunserv', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { id: 'appEvent', label: 'Appsheet Event', sheetName: 'Appsheet Event', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 },
    { id: 'appPastoral', label: 'Appsheet Pastoral', sheetName: 'Appsheet Pastoral', idCol: 1, lastNameCol: 2, firstNameCol: 3, headerRows: 1 }
  ];

  var externalSs = null;
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) {
    var externalRef = configSheet.getRange('B2').getValue();
    if (externalRef) {
      var externalId = extractSpreadsheetIdFromString_(String(externalRef));
      externalSs = SpreadsheetApp.openById(externalId);
    } else {
      Logger.log("checkNameCountsAndExtras: Config!B2 is empty, cannot read external Directory.");
    }
  } else {
    Logger.log("checkNameCountsAndExtras: Config sheet not found.");
  }

  var results = {};

  for (var t = 0; t < tabConfigs.length; t++) {
    var cfg = tabConfigs[t];
    var sheet = null;

    if (cfg.external) {
      if (!externalSs) continue;
      sheet = externalSs.getSheetByName(cfg.sheetName);
    } else {
      sheet = getSheetByNameLoose_(ss, cfg.sheetName);
    }

    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    var dataStartRow = cfg.headerRows + 1;

    if (lastRow < dataStartRow) {
      results[cfg.id] = { label: cfg.label, uniqueCount: 0, map: {} };
      Logger.log(cfg.label + ": no data rows.");
      continue;
    }

    var numRows = lastRow - cfg.headerRows;

    var ids = sheet.getRange(dataStartRow, cfg.idCol, numRows, 1).getValues();
    var lastNames = sheet.getRange(dataStartRow, cfg.lastNameCol, numRows, 1).getValues();
    var firstNames = sheet.getRange(dataStartRow, cfg.firstNameCol, numRows, 1).getValues();

    var nameMap = {};

    for (var i = 0; i < numRows; i++) {
      var pid = (ids[i][0] || '').toString().trim();
      if (!pid) continue;

      var ln = (lastNames[i][0] || '').toString().trim();
      var fn = (firstNames[i][0] || '').toString().trim();
      if (!ln && !fn) continue;

      var keyMatch = buildPersonKey_(pid, ln, fn, '');
      if (!keyMatch) continue;

      if (!nameMap[keyMatch]) {
        nameMap[keyMatch] = { personalId: pid, lastName: ln, firstName: fn };
      }
    }

    var count = 0;
    for (var kk in nameMap) {
      if (Object.prototype.hasOwnProperty.call(nameMap, kk)) count++;
    }

    results[cfg.id] = { label: cfg.label, uniqueCount: count, map: nameMap };
    Logger.log(cfg.label + ": unique (with ID only) = " + count);
  }

  var dirRes = results['directory'];
  if (!dirRes) return;

  var dirMap = dirRes.map;
  var appsheetIds = ['appSunserv', 'appEvent', 'appPastoral'];

  for (var a = 0; a < appsheetIds.length; a++) {
    var id = appsheetIds[a];
    var res = results[id];
    if (!res) continue;

    var extras = [];
    for (var key in res.map) {
      if (!Object.prototype.hasOwnProperty.call(res.map, key)) continue;
      if (!dirMap[key]) {
        var v = res.map[key];
        extras.push(v.personalId + ' | ' + v.lastName + ', ' + v.firstName);
      }
    }

    Logger.log(res.label + ": (with ID) NOT in Directory = " + extras.length);

    if (extras.length > 0) {
      var previewCount = Math.min(30, extras.length);
      Logger.log(res.label + " preview: " + extras.slice(0, previewCount).join(" | "));
    }
  }
}

