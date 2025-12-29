/**
 * Processes the 'Attendance Log' sheet and updates:
 * - Sunday Service
 * - Event Attendance
 * - Pastoral Check-In
 *
 * REFINEMENT (per request):
 * - Matching uses Personal ID + Last Name + First Name.
 * - Personal ID in Attendance Log is Column B.
 * - Destination rows WITHOUT Personal ID in Column B are NOT used for matching.
 * - If no match, script adds a new row (same as original behavior).
 */
function processAttendanceLogV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CONFIGURATION ---
  const logSheetName = 'Attendance Log';
  const sunServiceSheetName = 'Sunday Service';
  const eventSheetName = 'Event Attendance';
  const pastoralSheetName = 'Pastoral Check-In';

  const typeColumnIndex = 6; // Column F in destination sheets

  // 'Sunday Service' sheet config
  const sunServiceDateRow = 2;
  const sunServiceDataStartCol = 9; // Column I
  const sunServiceDataStartRow = 4; // data starts on row 4

  // 'Event Attendance' sheet config
  const eventDateRow = 2;
  const eventNameRow = 3;
  const eventCountRow = 4;
  const eventDataStartCol = 9; // Column I
  const eventDataStartRow = 5; // data starts on row 5

  // 'Pastoral Check-In' sheet config
  const pastoralDataStartRow = 4; // data starts on row 4
  const pastoralRecentDateCol = 5;  // E
  const pastoralPreviousDateCol = 6; // F
  const pastoralNotesCol = 7; // G
  const pastoralExtraCol = 8; // H

  // Attendance Log columns
  // We read B:L
  // B=Personal ID, C=Last, D=First, E=Type, F=Event, G=Date, H=Timestamp, I=Status, J=Remarks, K=Notes, L=Extra
  const logStatusColumn = 9;  // Column I
  const logRemarksColumn = 10; // Column J
  const logNumColsToRead = 11; // B..L

  // Indices inside B:L array
  const logPersonalIdIndex = 0; // Col B
  const logLastNameIndex = 1;   // Col C
  const logFirstNameIndex = 2;  // Col D
  const logTypeIndex = 3;       // Col E
  const logEventNameIndex = 4;  // Col F
  const logEventDateIndex = 5;  // Col G
  const logStatusColIndex = 7;  // Col I
  const logRemarksColIndex = 8; // Col J
  const logNotesColIndex = 9;   // Col K
  const logExtraColIndex = 10;  // Col L

  const logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    Logger.log('Error: Source sheet "' + logSheetName + '" not found.');
    return;
  }

  const lastLogRow = logSheet.getLastRow();
  if (lastLogRow < 2) {
    Logger.log('No data rows in Attendance Log.');
    return;
  }

  // 1) Read all attendance log data at once (B:L)
  const logRange = logSheet.getRange(2, 2, lastLogRow - 1, logNumColsToRead);
  const logData = logRange.getValues();

  const attendanceRecords = [];

  // Filter for unprocessed rows
  for (let i = 0; i < logData.length; i++) {
    const row = logData[i];
    const status = row[logStatusColIndex];

    if (status === 'Logged') continue;

    const personalId = (row[logPersonalIdIndex] || '').toString().trim();
    const lastNameRaw = (row[logLastNameIndex] || '').toString().trim();
    const firstNameRaw = (row[logFirstNameIndex] || '').toString().trim();

    const eventName = row[logEventNameIndex];
    let eventDate = row[logEventDateIndex];

    // We accept missing first OR last name (per your note), but Personal ID is expected for strong matching.
    // If there's absolutely no name and no Personal ID, skip.
    if (!personalId && !lastNameRaw && !firstNameRaw) {
      continue;
    }

    // Fix/validate date
    if (!(eventDate instanceof Date) && eventDate) {
      try {
        eventDate = new Date(eventDate);
        if (isNaN(eventDate.getTime())) throw new Error('Invalid date string');
      } catch (e) {
        logData[i][logRemarksColIndex] = 'Skipped: Invalid date format.';
        continue;
      }
    }

    if (!eventName || !(eventDate instanceof Date)) continue;

    const formattedFullDate = (eventDate.getMonth() + 1) + '-' + eventDate.getDate() + '-' + eventDate.getFullYear();
    const formattedShortDate = (eventDate.getMonth() + 1) + '-' + eventDate.getDate();

    const key = buildAttendanceKey_(personalId, lastNameRaw, firstNameRaw);

    attendanceRecords.push({
      personalId: personalId,
      lastName: lastNameRaw,
      firstName: firstNameRaw,
      key: key,
      eventName: eventName.toString().trim(),
      eventDate: eventDate,
      formattedFullDate: formattedFullDate,
      formattedShortDate: formattedShortDate,
      type: row[logTypeIndex],
      notes: row[logNotesColIndex],
      extra: row[logExtraColIndex],
      originalLogRownum: i + 2
    });
  }

  if (attendanceRecords.length === 0) {
    Logger.log('No *new* valid attendance records found in the log.');
    return;
  }

  // 2) Prepare caches to hold sheet data
  const sunServiceSheet = ss.getSheetByName(sunServiceSheetName);
  const eventSheet = ss.getSheetByName(eventSheetName);
  const pastoralSheet = ss.getSheetByName(pastoralSheetName);

  let sunServiceData = null;
  let eventSheetData = null;
  let pastoralData = null;

  if (sunServiceSheet) {
    sunServiceData = prepareSheetDataWithPersonalId_(
      sunServiceSheet,
      sunServiceDataStartRow,
      sunServiceDataStartCol,
      [sunServiceDateRow],
      false,
      false
    );
  } else {
    Logger.log('Warning: "' + sunServiceSheetName + '" not found. Skipping.');
  }

  if (eventSheet) {
    eventSheetData = prepareSheetDataWithPersonalId_(
      eventSheet,
      eventDataStartRow,
      eventDataStartCol,
      [eventDateRow, eventNameRow],
      true,
      true
    );
  } else {
    Logger.log('Warning: "' + eventSheetName + '" not found. Skipping.');
  }

  if (pastoralSheet) {
    pastoralData = preparePastoralSheetDataWithPersonalId_(pastoralSheet, pastoralDataStartRow);
  } else {
    Logger.log('Warning: "' + pastoralSheetName + '" not found. Skipping.');
  }

  // 3) Process records in memory
  let recordsWereLogged = false;
  const processedLogs = new Set();

  for (const record of attendanceRecords) {
    const logDataIndex = record.originalLogRownum - 2;

    const logKey = record.key + '|' + record.eventName + '|' + record.formattedFullDate;
    if (processedLogs.has(logKey)) {
      logData[logDataIndex][logStatusColIndex] = 'Logged';
      logData[logDataIndex][logRemarksColIndex] = 'Duplicate log entry processed.';
      recordsWereLogged = true;
      continue;
    }

    try {
      const eventName = record.eventName;

      if (/sunday service/i.test(eventName)) {
        if (!sunServiceData) continue;

        const rowNum = sunServiceData.keyMap.get(record.key) || null;
        const colNum = sunServiceData.dateMap.get(record.formattedShortDate) || null;

        if (rowNum && colNum) {
          const arrayRow = rowNum - sunServiceDataStartRow;
          const arrayCol = colNum - sunServiceDataStartCol;

          if (sunServiceData.checkboxes[arrayRow] && sunServiceData.checkboxes[arrayRow][arrayCol] !== undefined) {
            sunServiceData.checkboxes[arrayRow][arrayCol] = true;
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = '';
            recordsWereLogged = true;
            processedLogs.add(logKey);
          }
        } else if (!rowNum) {
          // Add new row
          const nextRow = sunServiceData.nextBlankRow;

          // Column B = Personal ID, Column C = Last, Column D = First
          if (record.personalId) sunServiceSheet.getRange(nextRow, 2).setValue(record.personalId);
          if (record.lastName) sunServiceSheet.getRange(nextRow, 3).setValue(capitalizeName(record.lastName));
          if (record.firstName) sunServiceSheet.getRange(nextRow, 4).setValue(capitalizeName(record.firstName));

          SpreadsheetApp.flush();

          sunServiceData.keyMap.set(record.key, nextRow);

          const numCols = sunServiceData.checkboxes[0] ? sunServiceData.checkboxes[0].length : 0;
          const newCheckboxRow = Array(numCols).fill(false);

          if (colNum) {
            const arrayCol = colNum - sunServiceDataStartCol;
            newCheckboxRow[arrayCol] = true;
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = 'New person added.';
            processedLogs.add(logKey);
          } else {
            logData[logDataIndex][logRemarksColIndex] = 'New person added, but event date not found.';
          }

          sunServiceData.checkboxes.push(newCheckboxRow);
          sunServiceData.numRows++;
          sunServiceData.nextBlankRow++;
          recordsWereLogged = true;
        } else if (rowNum && !colNum) {
          logData[logDataIndex][logRemarksColIndex] = 'Date not found in Sunday Service sheet.';
          recordsWereLogged = true;
        }

      } else if (/pastoral check-?in/i.test(eventName)) {
        if (!pastoralData) continue;

        const rowNum = pastoralData.keyMap.get(record.key) || null;

        if (rowNum) {
          const recentCell = pastoralSheet.getRange(rowNum, pastoralRecentDateCol);
          const prevCell = pastoralSheet.getRange(rowNum, pastoralPreviousDateCol);
          const existingRecent = recentCell.getValue();

          if (existingRecent) {
            prevCell.setValue(existingRecent);
            prevCell.setHorizontalAlignment('center');
            prevCell.setVerticalAlignment('middle');
          }

          recentCell.setValue(record.eventDate);
          recentCell.setHorizontalAlignment('center');
          recentCell.setVerticalAlignment('middle');

          const notesCell = pastoralSheet.getRange(rowNum, pastoralNotesCol);
          notesCell.setValue(record.notes);
          notesCell.setHorizontalAlignment('left');
          notesCell.setVerticalAlignment('middle');

          const extraCell = pastoralSheet.getRange(rowNum, pastoralExtraCol);
          extraCell.setValue(record.extra);
          extraCell.setHorizontalAlignment('left');

          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = '';
          recordsWereLogged = true;
          processedLogs.add(logKey);

        } else {
          // Add new row
          const nextRow = pastoralData.nextBlankRow;

          if (record.personalId) pastoralSheet.getRange(nextRow, 2).setValue(record.personalId); // B
          if (record.lastName) pastoralSheet.getRange(nextRow, 3).setValue(capitalizeName(record.lastName)); // C
          if (record.firstName) pastoralSheet.getRange(nextRow, 4).setValue(capitalizeName(record.firstName)); // D

          const recentCell = pastoralSheet.getRange(nextRow, pastoralRecentDateCol);
          recentCell.setValue(record.eventDate);
          recentCell.setHorizontalAlignment('center');
          recentCell.setVerticalAlignment('middle');

          const notesCell = pastoralSheet.getRange(nextRow, pastoralNotesCol);
          notesCell.setValue(record.notes);
          notesCell.setHorizontalAlignment('left');
          notesCell.setVerticalAlignment('middle');

          const extraCell = pastoralSheet.getRange(nextRow, pastoralExtraCol);
          extraCell.setValue(record.extra);
          extraCell.setHorizontalAlignment('left');

          SpreadsheetApp.flush();

          pastoralData.keyMap.set(record.key, nextRow);
          pastoralData.nextBlankRow++;
          pastoralData.numRows++;

          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = 'New person added.';
          recordsWereLogged = true;
          processedLogs.add(logKey);
        }

      } else {
        // Other events -> Event Attendance
        if (!eventSheetData) continue;

        const eventKey = record.formattedFullDate + '_' + record.eventName.trim().toLowerCase();
        let colNum = eventSheetData.dateMap.get(eventKey) || null;

        if (!colNum) {
          // Find a placeholder col or append a new one
          const lastCol = eventSheet.getLastColumn();
          let placeholderCol = null;

          if (lastCol >= eventDataStartCol) {
            const width = lastCol - eventDataStartCol + 1;
            const nameRowValues = eventSheet.getRange(eventNameRow, eventDataStartCol, 1, width).getValues()[0];
            const dateRowValues = eventSheet.getRange(eventDateRow, eventDataStartCol, 1, width).getValues()[0];

            for (let i = 0; i < width; i++) {
              const nameCell = nameRowValues[i];
              const dateCell = dateRowValues[i];
              if (nameCell === 'Post event name here' && !dateCell) {
                placeholderCol = eventDataStartCol + i;
                break;
              }
            }
          }

          colNum = (placeholderCol !== null) ? placeholderCol : (eventSheetData.lastDataCol + 1);

          eventSheet.getRange(eventDateRow, colNum).setValue(record.eventDate);
          eventSheet.getRange(eventNameRow, colNum).setValue(record.eventName);

          const colLetter = eventSheet.getRange(1, colNum).getA1Notation().replace(/\d+/g, '');
          const formula = '=COUNTIF(' + colLetter + eventDataStartRow + ':' + colLetter + ', TRUE)';
          eventSheet.getRange(eventCountRow, colNum).setFormula(formula);

          if (eventSheetData.numRows > 0) {
            eventSheet.getRange(eventDataStartRow, colNum, eventSheetData.numRows, 1).insertCheckboxes();
          }

          SpreadsheetApp.flush();

          eventSheetData.dateMap.set(eventKey, colNum);

          if (placeholderCol === null) {
            eventSheetData.lastDataCol = colNum;
            eventSheetData.checkboxes.forEach(function (r) { r.push(false); });
          } else if (colNum > eventSheetData.lastDataCol) {
            eventSheetData.lastDataCol = colNum;
          }
        }

        const rowNum = eventSheetData.keyMap.get(record.key) || null;

        if (rowNum) {
          const arrayRow = rowNum - eventDataStartRow;
          const arrayCol = colNum - eventDataStartCol;

          if (eventSheetData.checkboxes[arrayRow] && eventSheetData.checkboxes[arrayRow][arrayCol] !== undefined) {
            eventSheetData.checkboxes[arrayRow][arrayCol] = true;
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = '';
            recordsWereLogged = true;
            processedLogs.add(logKey);
          }
        } else {
          // Add new row
          const nextRow = eventSheetData.nextBlankRow;

          if (record.personalId) eventSheet.getRange(nextRow, 2).setValue(record.personalId); // B
          if (record.lastName) eventSheet.getRange(nextRow, 3).setValue(capitalizeName(record.lastName)); // C
          if (record.firstName) eventSheet.getRange(nextRow, 4).setValue(capitalizeName(record.firstName)); // D
          eventSheet.getRange(nextRow, typeColumnIndex).setValue(record.type); // F

          SpreadsheetApp.flush();

          eventSheetData.keyMap.set(record.key, nextRow);

          const numCols = eventSheetData.checkboxes[0] ? eventSheetData.checkboxes[0].length : 0;
          const newCheckboxRow = Array(numCols).fill(false);

          const arrayCol = colNum - eventDataStartCol;
          if (arrayCol >= 0 && arrayCol < newCheckboxRow.length) newCheckboxRow[arrayCol] = true;

          eventSheetData.checkboxes.push(newCheckboxRow);
          eventSheetData.numRows++;
          eventSheetData.nextBlankRow++;

          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = 'New person added.';
          recordsWereLogged = true;
          processedLogs.add(logKey);
        }
      }

    } catch (e) {
      Logger.log('Error processing record at log row ' + record.originalLogRownum + ': ' + e.message);
    }
  }

  // 4) Write updates back
  if (sunServiceData && sunServiceData.checkboxes.length > 0 && sunServiceData.checkboxes[0].length > 0) {
    sunServiceSheet.getRange(
      sunServiceDataStartRow,
      sunServiceDataStartCol,
      sunServiceData.checkboxes.length,
      sunServiceData.checkboxes[0].length
    ).setValues(sunServiceData.checkboxes);
  }

  if (eventSheetData && eventSheetData.checkboxes.length > 0 && eventSheetData.checkboxes[0].length > 0) {
    eventSheet.getRange(
      eventDataStartRow,
      eventDataStartCol,
      eventSheetData.checkboxes.length,
      eventSheetData.checkboxes[0].length
    ).setValues(eventSheetData.checkboxes);
  }

  if (recordsWereLogged) {
    // Write back only Status (Col I) and Remarks (Col J)
    const statusData = logData.map(function (r) { return [r[logStatusColIndex], r[logRemarksColIndex]]; });
    logSheet.getRange(2, logStatusColumn, statusData.length, 2).setValues(statusData);
  }

  Logger.log('Attendance processing complete.');
}

/**
 * Builds the matching key used across Attendance Log and destination tabs.
 * Key = PersonalID|Last|First (normalized)
 *
 * NOTE:
 * - Allows missing first OR last name (per your note)
 * - Personal ID is the primary anchor
 */
function buildAttendanceKey_(personalId, lastName, firstName) {
  const pid = (personalId || '').toString().trim();
  const ln = normalizeKeyPart_(lastName);
  const fn = normalizeKeyPart_(firstName);

  // If PID is blank, we still return a weaker key so the script can fall back to "add new row" behavior.
  return pid + '|' + ln + '|' + fn;
}

/**
 * Normalizes a key part (keeps letters+digits, strips spaces/punct).
 */
function normalizeKeyPart_(v) {
  return (v || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u00C0-\u024F]/g, '');
}

/**
 * Reads destination sheet and builds:
 * - keyMap: key -> row number (ONLY for rows that have Personal ID in Column B)
 * - dateMap: date/event -> column
 * - checkboxes: grid values
 * - nextBlankRow: first truly empty row (based on B/C/D)
 */
function prepareSheetDataWithPersonalId_(sheet, dataStartRow, dataStartCol, dateKeyRows, useFullDate, isEventSheet) {
  // Find last row based on any data in B/C/D
  const bcdAll = sheet.getRange('B1:D' + sheet.getMaxRows()).getValues();
  let actualLastDataRow = 0;
  for (let i = bcdAll.length - 1; i >= 0; i--) {
    const pid = bcdAll[i][0];
    const ln = bcdAll[i][1];
    const fn = bcdAll[i][2];
    if (pid || ln || fn) {
      actualLastDataRow = i + 1;
      break;
    }
  }

  const nextBlankRow = actualLastDataRow < dataStartRow ? dataStartRow : actualLastDataRow + 1;
  const dataRowCount = actualLastDataRow >= dataStartRow ? (actualLastDataRow - dataStartRow + 1) : 0;

  // Build keyMap ONLY for rows with Personal ID in Col B
  const keyMap = new Map();
  if (dataRowCount > 0) {
    const slice = bcdAll.slice(dataStartRow - 1, actualLastDataRow);
    for (let i = 0; i < slice.length; i++) {
      const pid = (slice[i][0] || '').toString().trim(); // Col B
      const ln = (slice[i][1] || '').toString().trim();  // Col C
      const fn = (slice[i][2] || '').toString().trim();  // Col D

      if (!pid) continue; // IMPORTANT: no Personal ID -> do NOT use for matching

      const key = buildAttendanceKey_(pid, ln, fn);
      if (key && !keyMap.has(key)) {
        keyMap.set(key, i + dataStartRow);
      }
    }
  }

  // Build Date-to-Column Map
  const dateMap = new Map();
  const lastSheetCol = sheet.getLastColumn() || dataStartCol;

  const dateValues = sheet.getRange(dateKeyRows[0], 1, 1, lastSheetCol).getValues()[0];
  const nameValues = dateKeyRows[1] ? sheet.getRange(dateKeyRows[1], 1, 1, lastSheetCol).getValues()[0] : null;

  let lastDataCol = dataStartCol - 1;

  for (let i = dataStartCol - 1; i < lastSheetCol; i++) {
    const dateVal = dateValues[i];

    if (dateVal instanceof Date) {
      const formattedDate = useFullDate
        ? (dateVal.getMonth() + 1) + '-' + dateVal.getDate() + '-' + dateVal.getFullYear()
        : (dateVal.getMonth() + 1) + '-' + dateVal.getDate();

      let key;
      if (nameValues) {
        const eventName = nameValues[i] ? nameValues[i].toString().trim().toLowerCase() : '';
        key = formattedDate + '_' + eventName;
      } else {
        key = formattedDate;
      }

      dateMap.set(key, i + 1);
      lastDataCol = i + 1;

    } else if (dateVal === '' && (!nameValues || nameValues[i] === '')) {
      break;
    } else if (i >= dataStartCol - 1) {
      lastDataCol = i + 1;
    }
  }

  // Get checkbox values
  const numCols = lastDataCol >= dataStartCol ? (lastDataCol - dataStartCol + 1) : 0;
  let checkboxes = [];

  if (dataRowCount > 0) {
    if (numCols > 0) {
      const range = sheet.getRange(dataStartRow, dataStartCol, dataRowCount, numCols);
      checkboxes = range.getValues();
      if (isEventSheet) range.insertCheckboxes();
    } else {
      checkboxes = Array(dataRowCount).fill(0).map(function () { return []; });
    }
  }

  return {
    sheet: sheet,
    keyMap: keyMap,
    dateMap: dateMap,
    checkboxes: checkboxes,
    lastDataCol: lastDataCol,
    numRows: dataRowCount,
    nextBlankRow: nextBlankRow
  };
}

/**
 * Pastoral Check-In helper:
 * builds keyMap + nextBlankRow based on B/C/D,
 * but ONLY maps rows that have Personal ID in Col B.
 */
function preparePastoralSheetDataWithPersonalId_(sheet, dataStartRow) {
  const bcdAll = sheet.getRange('B1:D' + sheet.getMaxRows()).getValues();
  let actualLastDataRow = 0;
  for (let i = bcdAll.length - 1; i >= 0; i--) {
    const pid = bcdAll[i][0];
    const ln = bcdAll[i][1];
    const fn = bcdAll[i][2];
    if (pid || ln || fn) {
      actualLastDataRow = i + 1;
      break;
    }
  }

  const nextBlankRow = actualLastDataRow < dataStartRow ? dataStartRow : actualLastDataRow + 1;
  const dataRowCount = actualLastDataRow >= dataStartRow ? (actualLastDataRow - dataStartRow + 1) : 0;

  const keyMap = new Map();
  if (dataRowCount > 0) {
    const slice = bcdAll.slice(dataStartRow - 1, actualLastDataRow);
    for (let i = 0; i < slice.length; i++) {
      const pid = (slice[i][0] || '').toString().trim();
      const ln = (slice[i][1] || '').toString().trim();
      const fn = (slice[i][2] || '').toString().trim();

      if (!pid) continue; // IMPORTANT: no Personal ID -> do NOT use for matching

      const key = buildAttendanceKey_(pid, ln, fn);
      if (key && !keyMap.has(key)) {
        keyMap.set(key, i + dataStartRow);
      }
    }
  }

  return {
    sheet: sheet,
    keyMap: keyMap,
    numRows: dataRowCount,
    nextBlankRow: nextBlankRow
  };
}

/**
 * Capitalizes the first letter of each part of a name.
 */
function capitalizeName(nameStr) {
  if (!nameStr) return '';
  return nameStr.toLowerCase()
    .replace(/\b(\w)|(-(\w))/g, function (match, p1, p2, p3) {
      if (p1) return p1.toUpperCase();
      if (p3) return '-' + p3.toUpperCase();
      return match;
    });
}



/**
 * Reverse of processAttendanceLogV2:
 * Reads checkboxes from "Sunday Service" and "Event Attendance"
 * and writes attendance rows into "Attendance Log".
 *
 * - Sunday Service:
 *   - Data starts at row 4.
 *   - Last name in Col C, first name in Col D.
 *   - Dates are in row 2 (starting at Col I / index 9).
 *   - If checkbox is TRUE, log a row:
 *     - Col A: unique id (e.g. "7a60d5eb").
 *     - Col B: "FirstName LastName".
 *     - Col C: Last Name.
 *     - Col D: First Name.
 *     - Col F: "Sunday Service".
 *     - Col G: Date from row 2.
 *     - Col H: Timestamp (now).
 *     - Col I: "Logged".
 *
 * - Event Attendance:
 *   - Data starts at row 5.
 *   - Last name in Col C, first name in Col D.
 *   - Dates in row 2, event names in row 3.
 *   - Only use columns where:
 *       Row 2 has a date AND
 *       Row 3 has an event name AND
 *       Row 3 is NOT "Post event name here".
 *   - For each TRUE checkbox:
 *     - Same mapping as above, but Col F = event name (row 3).
 *
 * - Prevents duplicate log rows by checking existing
 *   (LastName, FirstName, EventName, Date) combinations
 *   already in "Attendance Log".
 */
function exportSheetsAttendanceToLogV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Attendance Log');
  if (!logSheet) {
    Logger.log('Sheet "Attendance Log" not found.');
    return;
  }

  const sunServiceSheet = ss.getSheetByName('Sunday Service');
  const eventSheet = ss.getSheetByName('Event Attendance');

  const timezone = ss.getSpreadsheetTimeZone() || 'GMT';

  // --- Helper: generate 8-char hex ID like "7a60d5eb" ---
  function generateUniqueId_() {
    return Math.random().toString(16).slice(2, 10);
  }

  // --- Helper: normalize key for dedupe (lname, fname, event, date) ---
  function makeKey_(lastName, firstName, eventName, dateObjOrStr) {
    if (!lastName || !firstName || !eventName || !dateObjOrStr) return '';
    let dateKey;
    if (dateObjOrStr instanceof Date) {
      dateKey = Utilities.formatDate(dateObjOrStr, timezone, 'yyyy-MM-dd');
    } else {
      const d = new Date(dateObjOrStr);
      if (!isNaN(d.getTime())) {
        dateKey = Utilities.formatDate(d, timezone, 'yyyy-MM-dd');
      } else {
        dateKey = String(dateObjOrStr);
      }
    }
    return [
      String(lastName).trim().toLowerCase(),
      String(firstName).trim().toLowerCase(),
      String(eventName).trim().toLowerCase(),
      dateKey
    ].join('|');
  }

  // --- Build set of existing log keys so we don't duplicate ---
  const existingKeys = new Set();
  const lastLogRow = logSheet.getLastRow();
  if (lastLogRow > 1) {
    // Read Col C (Last), D (First), F (Event), G (Date)
    const existingRange = logSheet.getRange(2, 1, lastLogRow - 1, 7).getValues();
    // [A,B,C,D,E,F,G]
    for (let i = 0; i < existingRange.length; i++) {
      const row = existingRange[i];
      const lastName = row[2];  // Col C
      const firstName = row[3]; // Col D
      const eventName = row[5]; // Col F
      const dateVal = row[6];   // Col G
      const key = makeKey_(lastName, firstName, eventName, dateVal);
      if (key) existingKeys.add(key);
    }
  }

  const newRows = [];
  const newKeys = new Set();
  const now = new Date();

  // --- Helper: push a new log row if not duplicate ---
  function maybeAddLogRow_(lastName, firstName, eventName, dateVal) {
    if (!lastName && !firstName) return;
    if (!eventName || !dateVal) return;

    let dateObj = dateVal;
    if (!(dateObj instanceof Date)) {
      const tmp = new Date(dateVal);
      if (isNaN(tmp.getTime())) return;
      dateObj = tmp;
    }

    const key = makeKey_(lastName, firstName, eventName, dateObj);
    if (!key) return;
    if (existingKeys.has(key) || newKeys.has(key)) return;

    newKeys.add(key);
    existingKeys.add(key);

    const fullName = `${firstName || ''} ${lastName || ''}`.trim();
    const uniqueId = generateUniqueId_();

    // Col A–I: [ID, FullName, LastName, FirstName, (Type blank), EventName, EventDate, Timestamp, Status]
    newRows.push([
      uniqueId,
      fullName,
      lastName || '',
      firstName || '',
      '',                    // Col E (Type) – left blank
      eventName,
      dateObj,
      new Date(),            // Col H: timestamp (now)
      'Logged'               // Col I: status
    ]);
  }

  // --- 1) From "Sunday Service" ---
  if (sunServiceSheet) {
    const sunDataStartRow = 4;   // data row 4
    const sunDataStartCol = 9;   // column I
    const lastRow = sunServiceSheet.getLastRow();
    const lastCol = sunServiceSheet.getLastColumn();

    if (lastRow >= sunDataStartRow && lastCol >= sunDataStartCol) {
      const numRows = lastRow - sunDataStartRow + 1;
      const numCols = lastCol - sunDataStartCol + 1;

      const nameValues = sunServiceSheet.getRange(sunDataStartRow, 3, numRows, 2).getValues(); // C–D
      const checkboxValues = sunServiceSheet.getRange(sunDataStartRow, sunDataStartCol, numRows, numCols).getValues();
      const dateRowValues = sunServiceSheet.getRange(2, sunDataStartCol, 1, numCols).getValues()[0];

      for (let r = 0; r < numRows; r++) {
        const lastName = nameValues[r][0];
        const firstName = nameValues[r][1];

        if (!lastName && !firstName) continue;

        for (let c = 0; c < numCols; c++) {
          const checked = checkboxValues[r][c];
          if (checked === true) {
            const dateVal = dateRowValues[c];
            if (!dateVal) continue;
            maybeAddLogRow_(lastName, firstName, 'Sunday Service', dateVal);
          }
        }
      }
    }
  } else {
    Logger.log('Sheet "Sunday Service" not found.');
  }

  // --- 2) From "Event Attendance" ---
  if (eventSheet) {
    const evtDataStartRow = 5;   // data row 5
    const evtDataStartCol = 9;   // column I
    const lastRow = eventSheet.getLastRow();
    const lastCol = eventSheet.getLastColumn();

    if (lastRow >= evtDataStartRow && lastCol >= evtDataStartCol) {
      const numRows = lastRow - evtDataStartRow + 1;
      const numCols = lastCol - evtDataStartCol + 1;

      const nameValues = eventSheet.getRange(evtDataStartRow, 3, numRows, 2).getValues(); // C–D
      const checkboxValues = eventSheet.getRange(evtDataStartRow, evtDataStartCol, numRows, numCols).getValues();

      const dateRowValues = eventSheet.getRange(2, evtDataStartCol, 1, numCols).getValues()[0];
      const eventNameValues = eventSheet.getRange(3, evtDataStartCol, 1, numCols).getValues()[0];

      for (let c = 0; c < numCols; c++) {
        const dateVal = dateRowValues[c];
        const eventName = eventNameValues[c];

        // Only process if:
        // - Row 2 has date
        // - Row 3 has event name
        // - Row 3 is NOT "Post event name here"
        if (!dateVal) continue;
        if (!eventName) continue;
        if (String(eventName).trim() === 'Post event name here') continue;

        for (let r = 0; r < numRows; r++) {
          const checked = checkboxValues[r][c];
          if (checked === true) {
            const lastName = nameValues[r][0];
            const firstName = nameValues[r][1];
            if (!lastName && !firstName) continue;
            maybeAddLogRow_(lastName, firstName, eventName, dateVal);
          }
        }
      }
    }
  } else {
    Logger.log('Sheet "Event Attendance" not found.');
  }

  // --- Write new rows into Attendance Log ---
  if (newRows.length > 0) {
    const startRow = lastLogRow > 1 ? lastLogRow + 1 : 2;
    logSheet.getRange(startRow, 1, newRows.length, 9).setValues(newRows);
    Logger.log(`Added ${newRows.length} new attendance rows into "Attendance Log".`);
  } else {
    Logger.log('No new attendance rows to add into "Attendance Log".');
  }
}
