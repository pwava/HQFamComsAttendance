/**
 * @OnlyCurrentDoc
 *
 * This script updates the membership status and personal details on 'Sunday Service'
 * and 'Event Attendance' sheets by cross-referencing with a central 'Directory' spreadsheet.
 *
 * UPDATE (as requested):
 * - Matching now prioritizes: Personal ID (Column B) + Last Name (Col C) + First Name (Col D)
 * - If Personal ID is missing, it falls back to name-only matching (same behavior as before).
 * - Also updates these tabs:
 *   - Appsheet SunServ
 *   - Appsheet Event
 *   - Appsheet Pastoral
 *   For AppSheet tabs:
 *     - Last Name = Column B
 *     - First Name = Column C
 *     - Gender = Column D
 *     - Lineage = Column E
 *     - Age = Column F
 *     - Type = Column G
 */

/**
 * Main function to be run manually.
 * It reads the Directory spreadsheet, builds a map of members,
 * and then processes 'Sunday Service', 'Event Attendance', 'Attendance Log', and AppSheet tabs.
 */
function processMemberStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found. Please create one with the Directory link in cell B2.");
    return;
  }

  // Get the URL or ID of the Directory Spreadsheet from Config!B2
  const directoryUrl = configSheet.getRange('B2').getValue();
  if (!directoryUrl) {
    Logger.log("Error: Cell B2 in 'Config' sheet is empty. Please provide the Directory Spreadsheet URL or ID.");
    return;
  }

  // Build the fast lookup maps from the Directory
  const directoryLookup = buildDirectoryMap(directoryUrl);
  if (!directoryLookup) {
    Logger.log('Failed to build directory map. Aborting.');
    return;
  }

  Logger.log(
    `Directory maps built successfully. byPidName=${directoryLookup.byPidName.size}, byName=${directoryLookup.byName.size}`
  );

  // Process the 'Sunday Service' sheet, starting from row 4
  processSheet(ss, 'Sunday Service', 4, directoryLookup);

  // Process the 'Event Attendance' sheet, starting from row 5
  processSheet(ss, 'Event Attendance', 5, directoryLookup);

  // Process the 'Attendance Log' sheet, starting from row 4
  // Column B = Personal ID, Column C = Last Name, Column D = First Name, Column E = Member/Guest
  processAttendanceLog(ss, 'Attendance Log', 4, directoryLookup);

  // AppSheet tabs (layout per your note)
  // Last = B, First = C, Gender = D, Lineage = E, Age = F, Type = G
  processAppSheetTab(ss, 'Appsheet SunServ', 2, directoryLookup);
  processAppSheetTab(ss, 'Appsheet Event', 2, directoryLookup);
  processAppSheetTab(ss, 'Appsheet Pastoral', 2, directoryLookup);

  Logger.log('Processing complete for all sheets.');
  SpreadsheetApp.flush();
}

/**
 * Creates lookup maps from the Directory spreadsheet.
 * - byPidName: PID + normalized Last + normalized First
 * - byName: normalized Last + normalized First (fallback)
 *
 * Directory assumed columns:
 *   Z = Personal ID
 *   C = Last Name
 *   D = First Name
 *   E = Gender
 *   F = Lineage
 *   H = Age
 *
 * @param {string} directoryUrl The URL or ID of the Directory spreadsheet.
 * @return {{byPidName: Map, byName: Map}|null}
 */
function buildDirectoryMap(directoryUrl) {
  let directorySpreadsheet;

  // Try opening by URL, then by ID
  try {
    directorySpreadsheet = SpreadsheetApp.openByUrl(directoryUrl);
  } catch (e) {
    try {
      directorySpreadsheet = SpreadsheetApp.openById(directoryUrl);
    } catch (e2) {
      Logger.log('Error: Could not open Directory spreadsheet. Invalid URL/ID in Config B2: ' + directoryUrl);
      Logger.log('Details: ' + e2);
      return null;
    }
  }

  const dirSheet = directorySpreadsheet.getSheetByName('Directory');
  if (!dirSheet) {
    Logger.log("Error: 'Directory' sheet not found in the linked spreadsheet.");
    return null;
  }

  const lastRow = dirSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Directory sheet is empty (no data found after row 1).');
    return { byPidName: new Map(), byName: new Map() };
  }

  // Read C:H (6 columns): C=Last, D=First, E=Gender, F=Lineage, G=?, H=Age
  const rangeCH = dirSheet.getRange(2, 3, lastRow - 1, 6);
  const valuesCH = rangeCH.getValues();

  // Read Z: Z=Personal ID
  const rangeZ = dirSheet.getRange(2, 26, lastRow - 1, 1);
  const valuesZ = rangeZ.getValues();

  const byPidName = new Map();
  const byName = new Map();

  for (let i = 0; i < valuesCH.length; i++) {
    const rowCH = valuesCH[i];

    const pid = valuesZ[i][0];     // Col Z
    const lastName = rowCH[0];     // Col C
    const firstName = rowCH[1];    // Col D
    const gender = rowCH[2];       // Col E
    const lineage = rowCH[3];      // Col F
    const age = rowCH[5];          // Col H

    if (!lastName && !firstName && !pid) continue;

    const normLast = normalizeString(lastName);
    const firstNameStr = String(firstName || '');
    const normFirst = normalizeString(firstNameStr.split(' ')[0]);

    if (!normLast || !normFirst) continue;

    const nameKey = normLast + '_' + normFirst;

    // Fallback map (name-only)
    if (!byName.has(nameKey)) {
      byName.set(nameKey, { gender: gender, lineage: lineage, age: age });
    }

    // Primary map (PID + name)
    const pidStr = String(pid || '').trim();
    if (pidStr) {
      const pidNameKey = pidStr + '_' + nameKey;
      if (!byPidName.has(pidNameKey)) {
        byPidName.set(pidNameKey, { gender: gender, lineage: lineage, age: age });
      }
    }
  }

  return { byPidName: byPidName, byName: byName };
}

/**
 * Processes a single sheet ('Sunday Service' or 'Event Attendance') to update member status.
 *
 * Assumed layout for these tabs:
 *   B = Personal ID
 *   C = Last Name
 *   D = First Name
 *   E = Gender
 *   F = Lineage
 *   G = Age
 *   H = Type
 *
 * @param {Spreadsheet} ss
 * @param {string} sheetName
 * @param {number} startRow
 * @param {{byPidName: Map, byName: Map}} directoryLookup
 */
function processSheet(ss, sheetName, startRow, directoryLookup) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Warning: Sheet '${sheetName}' not found. Skipping.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log(`Sheet '${sheetName}' has no data to process starting from row ${startRow}.`);
    return;
  }

  // Read/write B:H (7 columns)
  const numRows = lastRow - startRow + 1;
  const numCols = 7; // B, C, D, E, F, G, H
  const range = sheet.getRange(startRow, 2, numRows, numCols);
  const values = range.getValues();

  let membersFound = 0;
  let guestsFound = 0;

  for (let i = 0; i < values.length; i++) {
    const pid = values[i][0];       // Col B
    const lastName = values[i][1];  // Col C
    const firstName = values[i][2]; // Col D

    // If name cells are blank, mark as Guest (Type col H) and continue
    if (!lastName && !firstName) {
      values[i][6] = 'Guest'; // Col H
      guestsFound++;
      continue;
    }

    const normLast = normalizeString(lastName);
    const firstNameStr = String(firstName || '');
    const normFirst = normalizeString(firstNameStr.split(' ')[0]);

    let match = null;

    if (normLast && normFirst) {
      const nameKey = normLast + '_' + normFirst;

      const pidStr = String(pid || '').trim();
      if (pidStr) {
        const pidNameKey = pidStr + '_' + nameKey;
        match = directoryLookup.byPidName.get(pidNameKey) || null;
      }

      if (!match) {
        match = directoryLookup.byName.get(nameKey) || null;
      }

      if (sheetName === 'Sunday Service' && i === 0) {
        Logger.log(
          `[${sheetName}] First generated keys: pid="${String(pid || '').trim()}", nameKey="${nameKey}"`
        );
      }
    }

    if (match) {
      // Found: Update E, F, G, H
      values[i][3] = match.gender;  // Col E
      values[i][4] = match.lineage; // Col F
      values[i][5] = match.age;     // Col G
      values[i][6] = 'Member';      // Col H
      membersFound++;
    } else {
      // Not Found: Mark as Guest; preserve existing E/F/G
      values[i][6] = 'Guest';       // Col H
      guestsFound++;
    }
  }

  range.setValues(values);

  // --- Alignment updates (kept exactly as your original logic, applied to C:H) ---
  sheet.getRange(startRow, 3, numRows, 6).setVerticalAlignment('middle');     // C:H
  sheet.getRange(startRow, 3, numRows, 2).setHorizontalAlignment('left');     // C:D
  sheet.getRange(startRow, 5, numRows, 4).setHorizontalAlignment('center');   // E:H

  Logger.log(`Processed ${numRows} rows for '${sheetName}'. Found: ${membersFound} Members, ${guestsFound} Guests.`);
}

/**
 * Processes the 'Attendance Log' sheet to tag Member/Guest in column E.
 *
 * Assumed layout:
 *   B = Personal ID
 *   C = Last Name
 *   D = First Name
 *   E = Member/Guest
 *
 * @param {Spreadsheet} ss
 * @param {string} sheetName
 * @param {number} startRow
 * @param {{byPidName: Map, byName: Map}} directoryLookup
 */
function processAttendanceLog(ss, sheetName, startRow, directoryLookup) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Warning: Sheet '${sheetName}' not found. Skipping Attendance Log processing.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log(`Sheet '${sheetName}' has no data to process starting from row ${startRow}.`);
    return;
  }

  // Read/write B:E (4 columns): PID, Last, First, Type
  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, 2, numRows, 4);
  const values = range.getValues();

  let membersFound = 0;
  let guestsFound = 0;

  for (let i = 0; i < values.length; i++) {
    const pid = values[i][0];       // Col B
    const lastName = values[i][1];  // Col C
    const firstName = values[i][2]; // Col D

    // If both name cells are blank, skip (do not overwrite Type)
    if (!lastName && !firstName) continue;

    const normLast = normalizeString(lastName);
    const firstNameStr = String(firstName || '');
    const normFirst = normalizeString(firstNameStr.split(' ')[0]);

    let match = null;

    if (normLast && normFirst) {
      const nameKey = normLast + '_' + normFirst;

      const pidStr = String(pid || '').trim();
      if (pidStr) {
        const pidNameKey = pidStr + '_' + nameKey;
        match = directoryLookup.byPidName.get(pidNameKey) || null;
      }

      if (!match) {
        match = directoryLookup.byName.get(nameKey) || null;
      }
    }

    if (match) {
      values[i][3] = 'Member'; // Col E
      membersFound++;
    } else {
      values[i][3] = 'Guest';  // Col E
      guestsFound++;
    }
  }

  range.setValues(values);
  Logger.log(`Processed ${numRows} rows for '${sheetName}'. Found: ${membersFound} Members, ${guestsFound} Guests.`);
}

/**
 * Updates AppSheet tabs:
 *   Last = B, First = C, Gender = D, Lineage = E, Age = F, Type = G
 *
 * Matching used: name-only (Last+First) against Directory (fallback map).
 *
 * @param {Spreadsheet} ss
 * @param {string} sheetName
 * @param {number} startRow
 * @param {{byPidName: Map, byName: Map}} directoryLookup
 */
function processAppSheetTab(ss, sheetName, startRow, directoryLookup) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Warning: Sheet '${sheetName}' not found. Skipping.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log(`Sheet '${sheetName}' has no data to process starting from row ${startRow}.`);
    return;
  }

  // Read/write B:G (6 columns)
  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, 2, numRows, 6);
  const values = range.getValues();

  let membersFound = 0;
  let guestsFound = 0;

  for (let i = 0; i < values.length; i++) {
    const lastName = values[i][0];  // Col B
    const firstName = values[i][1]; // Col C

    if (!lastName && !firstName) {
      // If blank names, do not overwrite other fields; just set type
      values[i][5] = 'Guest'; // Col G
      guestsFound++;
      continue;
    }

    const normLast = normalizeString(lastName);
    const firstNameStr = String(firstName || '');
    const normFirst = normalizeString(firstNameStr.split(' ')[0]);

    let match = null;
    if (normLast && normFirst) {
      const nameKey = normLast + '_' + normFirst;
      match = directoryLookup.byName.get(nameKey) || null;
    }

    if (match) {
      values[i][2] = match.gender;  // Col D
      values[i][3] = match.lineage; // Col E
      values[i][4] = match.age;     // Col F
      values[i][5] = 'Member';      // Col G
      membersFound++;
    } else {
      values[i][5] = 'Guest';       // Col G
      guestsFound++;
    }
  }

  range.setValues(values);
  Logger.log(`Processed ${numRows} rows for '${sheetName}'. Found: ${membersFound} Members, ${guestsFound} Guests.`);
}

/**
 * Normalizes a string for matching.
 * Converts to lowercase, trims whitespace, and removes non-alphabetic characters.
 * @param {string} str The string to normalize.
 * @return {string} The normalized string.
 */
function normalizeString(str) {
  if (!str || typeof str !== 'string') {
    return '';
  }
  return str.toLowerCase().trim().replace(/[^a-z]/g, '');
}

