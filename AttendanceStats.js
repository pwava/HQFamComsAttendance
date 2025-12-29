/**
 * A more robust function to clean and standardize names.
 * Removes all spaces and keeps only letters/numbers/commas for a consistent key.
 */
function normalizeName(name) {
  if (!name) return '';
  return name.toString().toLowerCase().replace(/[^a-z0-9,]/g, '').trim();
}

/**
 * Normalize Personal ID for matching (keep letters+numbers only).
 */
function normalizePersonalId(id) {
  if (!id) return '';
  return String(id).toLowerCase().trim().replace(/[^a-z0-9]/g, '');
}

/**
 * Build the matching key:
 * PersonalID + Last + First
 * (Last/First can be blank; PersonalID can be blank)
 */
function buildMatchKey(personalId, lastName, firstName) {
  const pid = normalizePersonalId(personalId);
  const last = normalizeName(lastName || '');
  const first = normalizeName(firstName || '');
  return `${pid}|${last}|${first}`;
}

/**
 * Creates a person-key used ONLY for issuing/reusing Personal IDs.
 * Works even if only one of (last/first) is present.
 */
function buildPersonKeyForId_(lastName, firstName) {
  const last = normalizeName(lastName || '');
  const first = normalizeName(firstName || '');
  // If both blank, no key.
  if (!last && !first) return '';
  return `${last}|${first}`;
}

/**
 * Generates a unique 9-character (letters+numbers) Personal ID in ALL CAPS.
 * Ensures no duplicates against the provided normalized-id Set.
 */
function generateUniquePersonalId_(existingIdSet) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let id = "";
  let norm = "";

  do {
    id = "";
    for (let i = 0; i < 9; i++) {
      id += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    norm = normalizePersonalId(id);
  } while (existingIdSet.has(norm));

  existingIdSet.add(norm);
  return id;
}

/**
 * Ensures rows with names but missing Personal ID get assigned a unique ID (9 chars, A-Z/0-9, all caps),
 * reusing an existing ID for the same person-key when available.
 *
 * Sheets covered (Personal ID is Column B):
 * - Sunday Service
 * - Event Attendance
 * - Pastoral Check-In
 * - Appsheet SunServ
 * - Appsheet Event
 * - Appsheet Pastoral
 * - Attendance Log
 *
 * Also tries to borrow ID from Directory by (Last, First) when available before generating a new one.
 *
 * Returns:
 * {
 *   existingIdSet,
 *   personKeyToId
 * }
 */
function ensurePersonalIdsAcrossAttendanceTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- External Directory ---
  let externalDirectorySs = null;
  try {
    const configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      const directoryId = configSheet.getRange("B2").getValue();
      if (directoryId) {
        externalDirectorySs = SpreadsheetApp.openById(directoryId);
      } else {
        Logger.log('❌ "Config!B2" is empty.');
      }
    } else {
      Logger.log('❌ "Config" sheet not found.');
    }
  } catch (e) {
    Logger.log(`❌ Error opening Directory sheet: ${e}`);
  }

  const directorySheet = externalDirectorySs ? externalDirectorySs.getSheetByName("Directory") : null;
  const dData = directorySheet ? directorySheet.getDataRange().getValues() : [];

  // Build:
  // - existingIdSet (normalized ids) from Directory + all attendance tabs
  // - directoryNameToId (borrow by name)
  // - personKeyToId (reuse for same person across tabs)
  const existingIdSet = new Set();
  const directoryNameToId = new Map();
  const personKeyToId = new Map();

  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const pid = row[25];    // Column Z
      const last = row[2];    // Column C
      const first = row[3];   // Column D

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) existingIdSet.add(pidNorm);

      const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
      if (nameKey && pidNorm && !directoryNameToId.has(nameKey)) {
        directoryNameToId.set(nameKey, pid);
      }

      const pKey = buildPersonKeyForId_(last, first);
      if (pKey && pidNorm && !personKeyToId.has(pKey)) {
        personKeyToId.set(pKey, String(pid).toUpperCase().trim());
      }
    });
  }

  // Helper: scan a sheet, collect existing IDs, then fill missing IDs in column B.
  const scanAndFillSheet_ = (sheetName, startRow, lastColIndexForNames) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < startRow) return;

    // Read needed columns:
    // Column B (2) for Personal ID, Column C (3) Last, Column D (4) First
    // We will read B:D for the range and write back ONLY column B.
    const numRows = lastRow - startRow + 1;
    const bToD = sh.getRange(startRow, 2, numRows, 3).getValues(); // B:D

    // First pass: collect IDs already present in this sheet into sets/maps
    bToD.forEach(r => {
      const pid = r[0];
      const last = r[1];
      const first = r[2];

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) existingIdSet.add(pidNorm);

      const pKey = buildPersonKeyForId_(last, first);
      if (pKey && pidNorm && !personKeyToId.has(pKey)) {
        personKeyToId.set(pKey, String(pid).toUpperCase().trim());
      }
    });

    // Second pass: fill missing IDs when there is at least a name part
    let changed = false;
    for (let i = 0; i < bToD.length; i++) {
      const pid = bToD[i][0];
      const last = bToD[i][1];
      const first = bToD[i][2];

      if (normalizePersonalId(pid)) continue;

      const pKey = buildPersonKeyForId_(last, first);
      if (!pKey) continue;

      // 1) Reuse from prior sheets (same person-key)
      let finalId = personKeyToId.get(pKey) || "";

      // 2) Borrow from Directory by exact "Last, First" (if both present)
      if (!finalId) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) {
          finalId = String(directoryNameToId.get(nameKey) || "").toUpperCase().trim();
        }
      }

      // 3) Generate new unique ID
      if (!finalId) {
        finalId = generateUniquePersonalId_(existingIdSet);
      } else {
        // Ensure the borrowed/reused id is also reserved in the set
        const borrowedNorm = normalizePersonalId(finalId);
        if (borrowedNorm) existingIdSet.add(borrowedNorm);
      }

      bToD[i][0] = finalId; // write into Column B
      personKeyToId.set(pKey, finalId);
      changed = true;
    }

    if (changed) {
      const colBOnly = bToD.map(r => [r[0]]);
      sh.getRange(startRow, 2, numRows, 1).setValues(colBOnly);
    }
  };

  // Sheets:
  // Event Attendance has 3 header rows; data starts row 4
  scanAndFillSheet_("Event Attendance", 4);
  // Sunday Service has 2 header rows; data starts row 3
  scanAndFillSheet_("Sunday Service", 3);
  // Pastoral Check-In commonly starts after header row; use row 3 to be safe
  scanAndFillSheet_("Pastoral Check-In", 3);

  // AppSheet tabs (assumed header row 1; data starts row 2)
  scanAndFillSheet_("Appsheet SunServ", 2);
  scanAndFillSheet_("Appsheet Event", 2);
  scanAndFillSheet_("Appsheet Pastoral", 2);

  // Attendance Log (assumed header row 1; data starts row 2)
  scanAndFillSheet_("Attendance Log", 2);

  return { existingIdSet, personKeyToId };
}

/**
 * Fetches all raw data from sheets.
 * Reads Directory from external sheet ID in Config!B2.
 * Reads Event Attendance and Sunday Service.
 * Reads Attendance Log for Pastoral Check-In.
 */
function getDataFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let externalDirectorySs = null;

  // --- External Directory ---
  try {
    const configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      const directoryId = configSheet.getRange("B2").getValue();
      if (directoryId) {
        externalDirectorySs = SpreadsheetApp.openById(directoryId);
      } else {
        Logger.log('❌ "Config!B2" is empty.');
      }
    } else {
      Logger.log('❌ "Config" sheet not found.');
    }
  } catch (e) {
    Logger.log(`❌ Error opening Directory sheet: ${e}`);
  }

  const getSheetData = (sheetName, spreadsheet = ss) => {
    const sheet = spreadsheet ? spreadsheet.getSheetByName(sheetName) : null;
    if (!sheet) {
      Logger.log(`❌ Sheet "${sheetName}" not found.`);
      return [];
    }
    return sheet.getDataRange().getValues();
  };

  return {
    dData: getSheetData("Directory", externalDirectorySs),
    eData: getSheetData("Event Attendance", ss),
    sData: getSheetData("Sunday Service", ss),
    lData: getSheetData("Attendance Log", ss)
  };
}

/**
 * Collects attendance from:
 *  - Directory (for lookup only)
 *  - Sunday Service
 *  - Event Attendance
 *  - Attendance Log (Pastoral Check-In only)
 *
 * IMPORTANT CHANGE:
 * - Personal ID is Column B in ALL relevant sheets.
 * - Matching uses PersonalID + Last + First (names can be partially blank).
 * - No BEL generation in this script.
 */
function matchOrAssignBelCodes() {
  // NEW: Ensure Personal IDs exist (and are written back to their source sheets) before collecting.
  ensurePersonalIdsAcrossAttendanceTabs_();

  const data = getDataFromSheets();
  if (!data) return { rawData: [], dData: [] };

  const { sData, eData, dData, lData } = data;

  // Directory lookup:
  // - ID set for guest detection
  // - optional name->id mapping (helps when attendance row missing ID)
  const directoryIdSet = new Set();
  const directoryNameToId = new Map();

  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const pid = row[25];    // Column Z = Personal ID
      const last = row[2];    // Column C = Last
      const first = row[3];   // Column D = First

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) directoryIdSet.add(pidNorm);

      const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
      if (nameKey && pidNorm && !directoryNameToId.has(nameKey)) {
        directoryNameToId.set(nameKey, pid);
      }
    });
  }

  const results = [];

  // --- EVENT ATTENDANCE ---
  if (eData && eData.length > 3) {
    const dates = eData[1];
    const names = eData[2];

    eData.slice(3).forEach(row => {
      const personalId = row[1]; // Column B
      const last = row[2];       // Column C
      const first = row[3];      // Column D

      // If Personal ID missing, try to borrow from Directory using name (no generation here)
      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          const eventName = names[c];
          if (date && eventName) {
            results.push([pidFinal || "", first || "", last || "", eventName, eventName, date, false, matchKey]);
          }
        }
      }
    });
  }

  // --- SUNDAY SERVICE ---
  if (sData && sData.length > 2) {
    const dates = sData[1];

    sData.slice(2).forEach(row => {
      const personalId = row[1]; // Column B
      const last = row[2];       // Column C
      const first = row[3];      // Column D

      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          if (date) {
            results.push([pidFinal || "", first || "", last || "", "Sunday Service", "Sunday Service", date, false, matchKey]);
          }
        }
      }
    });
  }

  // --- PASTORAL CHECK-IN FROM ATTENDANCE LOG (DEDUPLICATED PER DATE) ---
  if (lData && lData.length > 1) {
    const pastoralSeen = new Set(); // MatchKey|DATE

    lData.slice(1).forEach(row => {
      const event = row[5];
      if (!event) return;
      if (String(event).toLowerCase().trim() !== "pastoral check-in") return;

      const personalId = row[1]; // Column B = Personal ID
      const last = row[2];
      const first = row[3];

      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      const date = row[6];
      if (!date) return;

      const dedupeKey = `${matchKey}|${new Date(date).toDateString()}`;
      if (pastoralSeen.has(dedupeKey)) return;

      pastoralSeen.add(dedupeKey);
      results.push([pidFinal || "", first || "", last || "", "Pastoral Check-In", "Pastoral Check-In", date, false, matchKey]);
    });
  }

  return { rawData: results, dData, directoryIdSet };
}

/**
 * Calculates stats (Q1–Q4, Total, Last Event, Guest Flag) for YEAR 2026.
 * FIXED: Column M returns full event name even if it contains hyphens,
 * and Pastoral Check-In is forced to exact label.
 *
 * IMPORTANT CHANGE:
 * - Grouping is by PersonalID+Last+First match key (stored at index 7 in raw records).
 * - Guest detection includes matching Personal ID.
 */
function calculateAttendanceStats() {
  const { rawData, dData, directoryIdSet } = matchOrAssignBelCodes();
  if (rawData.length === 0) return [];

  // Fallback name set (only used if Personal ID is missing)
  const directoryNamesSet = new Set();
  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const lastName = row[2];
      const firstName = row[3];
      if (lastName || firstName) {
        directoryNamesSet.add(normalizeName(`${lastName || ''}, ${firstName || ''}`));
      }
    });
  }

  const reportYear = 2026;

  const q1_start = new Date(reportYear, 0, 1);
  const q1_end   = new Date(reportYear, 3, 0);
  const q2_start = new Date(reportYear, 3, 1);
  const q2_end   = new Date(reportYear, 6, 0);
  const q3_start = new Date(reportYear, 6, 1);
  const q3_end   = new Date(reportYear, 9, 0);
  const q4_start = new Date(reportYear, 9, 1);
  const q4_end   = new Date(reportYear, 11, 31);

  const grouped = new Map(); // matchKey -> records[]

  rawData.forEach(row => {
    if (row.length < 8) return;

    const personalId = row[0];
    const firstName = row[1];
    const lastName = row[2];
    const eventName = row[3];
    const dateVal = row[5];
    const isVolunteer = row[6];
    const matchKey = row[7];

    const date = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(date.getTime())) return;

    // Only count records inside 2026 for Q/Totals/Last Event (2026)
    if (date.getFullYear() !== reportYear) return;

    const isSundayService = /sunday service/i.test(eventName);
    const eventKey = isSundayService
      ? `Sunday Service-${date.toDateString()}`
      : `${eventName}-${date.toDateString()}`;

    const record = {
      personalId,
      firstName,
      lastName,
      date,
      eventKey,
      isVolunteer: isVolunteer === true
    };

    if (!grouped.has(matchKey)) grouped.set(matchKey, []);
    grouped.get(matchKey).push(record);
  });

  const summary = [];
  grouped.forEach((records, matchKey) => {
    if (!records || records.length === 0) return;

    const q1Events = new Set(),
          q2Events = new Set(),
          q3Events = new Set(),
          q4Events = new Set();

    records.forEach(r => {
      if (r.date >= q1_start && r.date <= q1_end) q1Events.add(r.eventKey);
      if (r.date >= q2_start && r.date <= q2_end) q2Events.add(r.eventKey);
      if (r.date >= q3_start && r.date <= q3_end) q3Events.add(r.eventKey);
      if (r.date >= q4_start && r.date <= q4_end) q4Events.add(r.eventKey);
    });

    records.sort((a, b) => b.date.getTime() - a.date.getTime());
    const mostRecentRecord = records[0];

    const lastDashIndex = mostRecentRecord.eventKey.lastIndexOf("-");
    let lastEventName = lastDashIndex > -1
      ? mostRecentRecord.eventKey.substring(0, lastDashIndex)
      : mostRecentRecord.eventKey;

    if (/pastoral\s*check[-\s]*in/i.test(lastEventName)) {
      lastEventName = "Pastoral Check-In";
    }

    const pidNorm = normalizePersonalId(mostRecentRecord.personalId);
    const nameNorm = normalizeName(`${mostRecentRecord.lastName || ''}, ${mostRecentRecord.firstName || ''}`);

    // Guest logic:
    // - If Personal ID exists: guest if NOT in Directory Personal IDs
    // - Else (no ID): fallback to name-based check
    let isGuest = false;
    if (pidNorm) {
      isGuest = !(directoryIdSet && directoryIdSet.has(pidNorm));
    } else {
      isGuest = !directoryNamesSet.has(nameNorm);
    }

    const guestStatus = isGuest ? "Guest" : "";

    summary.push([
      mostRecentRecord.personalId || "",     // Personal ID (will go to Column B in Attendance Stats)
      mostRecentRecord.firstName || "",
      mostRecentRecord.lastName || "",
      q1Events.size,
      q2Events.size,
      q3Events.size,
      q4Events.size,
      q1Events.size + q2Events.size + q3Events.size + q4Events.size,
      mostRecentRecord.date,
      lastEventName,
      guestStatus,
      matchKey
    ]);
  });

  return summary;
}

/**
 * Update activity level (Column F) based on attendance in the past 91 days.
 * If last attendance is over 12 months ago → "Archive".
 *
 * IMPORTANT CHANGE:
 * - Matching to Attendance Stats is by (Personal ID + Last + First)
 */
function updateActivityLevels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() < 3) return;

  const { rawData } = matchOrAssignBelCodes();
  const today = new Date();

  const cutoff91 = new Date(today);
  cutoff91.setDate(cutoff91.getDate() - 91);

  const cutoff12mo = new Date(today);
  cutoff12mo.setFullYear(cutoff12mo.getFullYear() - 1);

  // Map: matchKey -> { count91, lastDate }
  const attMap = new Map();

  rawData.forEach(r => {
    if (!r || r.length < 8) return;

    const dateVal = r[5];
    const matchKey = r[7];

    if (!matchKey || !dateVal) return;

    const d = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(d.getTime())) return;

    if (!attMap.has(matchKey)) attMap.set(matchKey, { count91: 0, lastDate: null });
    const obj = attMap.get(matchKey);

    if (!obj.lastDate || d > obj.lastDate) obj.lastDate = d;
    if (d >= cutoff91) obj.count91++;
  });

  const lastRow = sheet.getLastRow();

  // Read Personal ID (B), Last (C), First (D)
  const pidLastFirst = sheet.getRange(3, 2, lastRow - 2, 3).getValues(); // B3:D
  const out = pidLastFirst.map(([pid, last, first]) => {
    const key = buildMatchKey(pid, last, first);
    const info = key ? attMap.get(key) : null;

    const lastDate = info ? info.lastDate : null;
    const count91 = info ? info.count91 : 0;

    if (lastDate && lastDate < cutoff12mo) return ["Archive"];
    if (count91 >= 12) return ["Core"];
    if (count91 >= 3) return ["Active"];
    return ["Inactive"];
  });

  sheet.getRange(3, 6, out.length, 1).setValues(out);
}

/**
 * Sort + formatting for final output.
 * - Guests on top
 * - Core, Active, Inactive, Archive
 * - Then sort alphabetically by Last, then First
 *
 * IMPORTANT CHANGE:
 * - Column B (Personal ID) always left aligned + vertical middle
 * - Column M (Last Event) left aligned + vertical middle
 */
function performFinalSort() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() <= 2) return;

  const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const data = range.getValues();

  const order = { "Core": 1, "Active": 2, "Inactive": 3, "Archive": 4 };

  data.sort((a, b) => {
    const guestA = a[4] === "Guest";
    const guestB = b[4] === "Guest";
    if (guestA !== guestB) return guestA ? -1 : 1;

    const aLvl = order[a[5]] || 99;
    const bLvl = order[b[5]] || 99;
    if (aLvl !== bLvl) return aLvl - bLvl;

    const lastA = (a[2] || "").toString().toLowerCase();
    const lastB = (b[2] || "").toString().toLowerCase();
    if (lastA < lastB) return -1;
    if (lastA > lastB) return 1;

    const firstA = (a[3] || "").toString().toLowerCase();
    const firstB = (b[3] || "").toString().toLowerCase();
    if (firstA < firstB) return -1;
    if (firstA > firstB) return 1;

    return 0;
  });

  range.setValues(data);
  range.clearFormat();

  const numRows = range.getNumRows();

  // Center E–L (Guest, Activity, Q1-Q4, Total, Last Date)
  const centerRange = sheet.getRange(3, 5, numRows, 8); // E (5) to L (12)
  centerRange.setHorizontalAlignment("center");
  centerRange.setVerticalAlignment("middle");

  // Column B Personal ID: left + middle
  const colBRange = sheet.getRange(3, 2, numRows, 1);
  colBRange.setHorizontalAlignment("left");
  colBRange.setVerticalAlignment("middle");

  // Column M Last Event: left + middle
  const colMRange = sheet.getRange(3, 13, numRows, 1);
  colMRange.setHorizontalAlignment("left");
  colMRange.setVerticalAlignment("middle");

  Logger.log("✅ Final sort and alignment complete (Column B & M left-aligned).");
}

/**
 * Main update flow.
 *
 * OUTPUT COLUMNS (13):
 * A: blank
 * B: Personal ID
 * C: Last Name
 * D: First Name
 * E: Guest
 * F: Activity Level (filled by updateActivityLevels)
 * G: Q1
 * H: Q2
 * I: Q3
 * J: Q4
 * K: Total
 * L: Last Date
 * M: Last Event
 */
function updateAttendanceStatsSheet() {
  Logger.log("🚀 Starting the process to update the 'Attendance Stats' sheet...");

  const data = calculateAttendanceStats();
  if (data.length === 0) {
    Logger.log("No data to update.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance Stats");
  if (!sheet) return;

  const output = data.map(row => {
    const [
      personalId, first, last, q1, q2, q3, q4, total, lastDate, lastEvent, guest
    ] = row;

    const formattedDate = lastDate instanceof Date
      ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy")
      : "";

    return [
      "",              // A
      personalId || "",// B
      last || "",      // C
      first || "",     // D
      guest || "",     // E
      "",              // F (Activity Level computed later)
      q1 || 0,         // G
      q2 || 0,         // H
      q3 || 0,         // I
      q4 || 0,         // J
      total || 0,      // K
      formattedDate,   // L
      lastEvent || ""  // M
    ];
  });

  const maxRows = sheet.getMaxRows();
  if (sheet.getLastRow() > 2) {
    sheet.getRange(3, 1, maxRows - 2, 13).clearContent().clearFormat();
  }

  sheet.getRange(3, 1, output.length, 13).setValues(output);

  updateActivityLevels();
  performFinalSort();

  Logger.log("✅ Finished updating Attendance Stats.");
}

/**
 * Manual run.
 */
function runManualUpdate() {
  updateAttendanceStatsSheet();
  SpreadsheetApp.getUi().alert('The "Attendance Stats" sheet has been successfully updated.');
}
function ensurePersonalIdsAcrossAttendanceTabs_ForSpreadsheet_(ss) {
  // --- External Directory ---
  let externalDirectorySs = null;
  try {
    const configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      const directoryId = configSheet.getRange("B2").getValue();
      if (directoryId) {
        externalDirectorySs = SpreadsheetApp.openById(directoryId);
      } else {
        Logger.log('❌ "Config!B2" is empty.');
      }
    } else {
      Logger.log('❌ "Config" sheet not found.');
    }
  } catch (e) {
    Logger.log(`❌ Error opening Directory sheet: ${e}`);
  }

  const directorySheet = externalDirectorySs ? externalDirectorySs.getSheetByName("Directory") : null;
  const dData = directorySheet ? directorySheet.getDataRange().getValues() : [];

  const existingIdSet = new Set();
  const directoryNameToId = new Map();
  const personKeyToId = new Map();

  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const pid = row[25];    // Column Z
      const last = row[2];    // Column C
      const first = row[3];   // Column D

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) existingIdSet.add(pidNorm);

      const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
      if (nameKey && pidNorm && !directoryNameToId.has(nameKey)) {
        directoryNameToId.set(nameKey, pid);
      }

      const pKey = buildPersonKeyForId_(last, first);
      if (pKey && pidNorm && !personKeyToId.has(pKey)) {
        personKeyToId.set(pKey, String(pid).toUpperCase().trim());
      }
    });
  }

  const scanAndFillSheet_ = (sheetName, startRow) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < startRow) return;

    const numRows = lastRow - startRow + 1;
    const bToD = sh.getRange(startRow, 2, numRows, 3).getValues(); // B:D

    // First pass: collect IDs already present
    bToD.forEach(r => {
      const pid = r[0];
      const last = r[1];
      const first = r[2];

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) existingIdSet.add(pidNorm);

      const pKey = buildPersonKeyForId_(last, first);
      if (pKey && pidNorm && !personKeyToId.has(pKey)) {
        personKeyToId.set(pKey, String(pid).toUpperCase().trim());
      }
    });

    // Second pass: fill missing IDs when there is at least a name part
    let changed = false;
    for (let i = 0; i < bToD.length; i++) {
      const pid = bToD[i][0];
      const last = bToD[i][1];
      const first = bToD[i][2];

      if (normalizePersonalId(pid)) continue;

      const pKey = buildPersonKeyForId_(last, first);
      if (!pKey) continue;

      let finalId = personKeyToId.get(pKey) || "";

      if (!finalId) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) {
          finalId = String(directoryNameToId.get(nameKey) || "").toUpperCase().trim();
        }
      }

      if (!finalId) {
        finalId = generateUniquePersonalId_(existingIdSet);
      } else {
        const borrowedNorm = normalizePersonalId(finalId);
        if (borrowedNorm) existingIdSet.add(borrowedNorm);
      }

      bToD[i][0] = finalId;
      personKeyToId.set(pKey, finalId);
      changed = true;
    }

    if (changed) {
      const colBOnly = bToD.map(r => [r[0]]);
      sh.getRange(startRow, 2, numRows, 1).setValues(colBOnly);
    }
  };

  scanAndFillSheet_("Event Attendance", 4);
  scanAndFillSheet_("Sunday Service", 3);
  scanAndFillSheet_("Pastoral Check-In", 3);

  scanAndFillSheet_("Appsheet SunServ", 2);
  scanAndFillSheet_("Appsheet Event", 2);
  scanAndFillSheet_("Appsheet Pastoral", 2);

  scanAndFillSheet_("Attendance Log", 2);

  return { existingIdSet, personKeyToId };
}

function getDataFromSheets_ForSpreadsheet_(ss) {
  let externalDirectorySs = null;

  try {
    const configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      const directoryId = configSheet.getRange("B2").getValue();
      if (directoryId) {
        externalDirectorySs = SpreadsheetApp.openById(directoryId);
      } else {
        Logger.log('❌ "Config!B2" is empty.');
      }
    } else {
      Logger.log('❌ "Config" sheet not found.');
    }
  } catch (e) {
    Logger.log(`❌ Error opening Directory sheet: ${e}`);
  }

  const getSheetData = (sheetName, spreadsheet = ss) => {
    const sheet = spreadsheet ? spreadsheet.getSheetByName(sheetName) : null;
    if (!sheet) {
      Logger.log(`❌ Sheet "${sheetName}" not found.`);
      return [];
    }
    return sheet.getDataRange().getValues();
  };

  return {
    dData: getSheetData("Directory", externalDirectorySs),
    eData: getSheetData("Event Attendance", ss),
    sData: getSheetData("Sunday Service", ss),
    lData: getSheetData("Attendance Log", ss)
  };
}

function matchOrAssignBelCodes_ForSpreadsheet_(ss) {
  ensurePersonalIdsAcrossAttendanceTabs_ForSpreadsheet_(ss);

  const data = getDataFromSheets_ForSpreadsheet_(ss);
  if (!data) return { rawData: [], dData: [] };

  const { sData, eData, dData, lData } = data;

  const directoryIdSet = new Set();
  const directoryNameToId = new Map();

  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const pid = row[25];    // Column Z
      const last = row[2];    // Column C
      const first = row[3];   // Column D

      const pidNorm = normalizePersonalId(pid);
      if (pidNorm) directoryIdSet.add(pidNorm);

      const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
      if (nameKey && pidNorm && !directoryNameToId.has(nameKey)) {
        directoryNameToId.set(nameKey, pid);
      }
    });
  }

  const results = [];

  // --- EVENT ATTENDANCE ---
  if (eData && eData.length > 3) {
    const dates = eData[1];
    const names = eData[2];

    eData.slice(3).forEach(row => {
      const personalId = row[1]; // B
      const last = row[2];       // C
      const first = row[3];      // D

      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          const eventName = names[c];
          if (date && eventName) {
            results.push([pidFinal || "", first || "", last || "", eventName, eventName, date, false, matchKey]);
          }
        }
      }
    });
  }

  // --- SUNDAY SERVICE ---
  if (sData && sData.length > 2) {
    const dates = sData[1];

    sData.slice(2).forEach(row => {
      const personalId = row[1]; // B
      const last = row[2];       // C
      const first = row[3];      // D

      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          if (date) {
            results.push([pidFinal || "", first || "", last || "", "Sunday Service", "Sunday Service", date, false, matchKey]);
          }
        }
      }
    });
  }

  // --- PASTORAL CHECK-IN FROM ATTENDANCE LOG ---
  if (lData && lData.length > 1) {
    const pastoralSeen = new Set();

    lData.slice(1).forEach(row => {
      const event = row[5];
      if (!event) return;
      if (String(event).toLowerCase().trim() !== "pastoral check-in") return;

      const personalId = row[1]; // B
      const last = row[2];
      const first = row[3];

      let pidFinal = personalId;
      if (!normalizePersonalId(pidFinal)) {
        const nameKey = normalizeName(`${last || ''}, ${first || ''}`);
        if (nameKey && directoryNameToId.has(nameKey)) pidFinal = directoryNameToId.get(nameKey);
      }

      const matchKey = buildMatchKey(pidFinal, last, first);
      if (matchKey === "||") return;

      const date = row[6];
      if (!date) return;

      const dedupeKey = `${matchKey}|${new Date(date).toDateString()}`;
      if (pastoralSeen.has(dedupeKey)) return;

      pastoralSeen.add(dedupeKey);
      results.push([pidFinal || "", first || "", last || "", "Pastoral Check-In", "Pastoral Check-In", date, false, matchKey]);
    });
  }

  return { rawData: results, dData, directoryIdSet };
}

function calculateAttendanceStats_ForSpreadsheet_(ss) {
  const { rawData, dData, directoryIdSet } = matchOrAssignBelCodes_ForSpreadsheet_(ss);
  if (rawData.length === 0) return [];

  const directoryNamesSet = new Set();
  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const lastName = row[2];
      const firstName = row[3];
      if (lastName || firstName) {
        directoryNamesSet.add(normalizeName(`${lastName || ''}, ${firstName || ''}`));
      }
    });
  }

  const reportYear = 2026;

  const q1_start = new Date(reportYear, 0, 1);
  const q1_end   = new Date(reportYear, 3, 0);
  const q2_start = new Date(reportYear, 3, 1);
  const q2_end   = new Date(reportYear, 6, 0);
  const q3_start = new Date(reportYear, 6, 1);
  const q3_end   = new Date(reportYear, 9, 0);
  const q4_start = new Date(reportYear, 9, 1);
  const q4_end   = new Date(reportYear, 11, 31);

  const grouped = new Map();

  rawData.forEach(row => {
    if (row.length < 8) return;

    const personalId = row[0];
    const firstName = row[1];
    const lastName = row[2];
    const eventName = row[3];
    const dateVal = row[5];
    const isVolunteer = row[6];
    const matchKey = row[7];

    const date = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(date.getTime())) return;

    if (date.getFullYear() !== reportYear) return;

    const isSundayService = /sunday service/i.test(eventName);
    const eventKey = isSundayService
      ? `Sunday Service-${date.toDateString()}`
      : `${eventName}-${date.toDateString()}`;

    const record = {
      personalId,
      firstName,
      lastName,
      date,
      eventKey,
      isVolunteer: isVolunteer === true
    };

    if (!grouped.has(matchKey)) grouped.set(matchKey, []);
    grouped.get(matchKey).push(record);
  });

  const summary = [];
  grouped.forEach((records, matchKey) => {
    if (!records || records.length === 0) return;

    const q1Events = new Set(),
          q2Events = new Set(),
          q3Events = new Set(),
          q4Events = new Set();

    records.forEach(r => {
      if (r.date >= q1_start && r.date <= q1_end) q1Events.add(r.eventKey);
      if (r.date >= q2_start && r.date <= q2_end) q2Events.add(r.eventKey);
      if (r.date >= q3_start && r.date <= q3_end) q3Events.add(r.eventKey);
      if (r.date >= q4_start && r.date <= q4_end) q4Events.add(r.eventKey);
    });

    records.sort((a, b) => b.date.getTime() - a.date.getTime());
    const mostRecentRecord = records[0];

    const lastDashIndex = mostRecentRecord.eventKey.lastIndexOf("-");
    let lastEventName = lastDashIndex > -1
      ? mostRecentRecord.eventKey.substring(0, lastDashIndex)
      : mostRecentRecord.eventKey;

    if (/pastoral\s*check[-\s]*in/i.test(lastEventName)) {
      lastEventName = "Pastoral Check-In";
    }

    const pidNorm = normalizePersonalId(mostRecentRecord.personalId);
    const nameNorm = normalizeName(`${mostRecentRecord.lastName || ''}, ${mostRecentRecord.firstName || ''}`);

    let isGuest = false;
    if (pidNorm) {
      isGuest = !(directoryIdSet && directoryIdSet.has(pidNorm));
    } else {
      isGuest = !directoryNamesSet.has(nameNorm);
    }

    const guestStatus = isGuest ? "Guest" : "";

    summary.push([
      mostRecentRecord.personalId || "",
      mostRecentRecord.firstName || "",
      mostRecentRecord.lastName || "",
      q1Events.size,
      q2Events.size,
      q3Events.size,
      q4Events.size,
      q1Events.size + q2Events.size + q3Events.size + q4Events.size,
      mostRecentRecord.date,
      lastEventName,
      guestStatus,
      matchKey
    ]);
  });

  return summary;
}

function updateActivityLevels_ForSpreadsheet_(ss) {
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() < 3) return;

  const { rawData } = matchOrAssignBelCodes_ForSpreadsheet_(ss);
  const today = new Date();

  const cutoff91 = new Date(today);
  cutoff91.setDate(cutoff91.getDate() - 91);

  const cutoff12mo = new Date(today);
  cutoff12mo.setFullYear(cutoff12mo.getFullYear() - 1);

  const attMap = new Map();

  rawData.forEach(r => {
    if (!r || r.length < 8) return;

    const dateVal = r[5];
    const matchKey = r[7];

    if (!matchKey || !dateVal) return;

    const d = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(d.getTime())) return;

    if (!attMap.has(matchKey)) attMap.set(matchKey, { count91: 0, lastDate: null });
    const obj = attMap.get(matchKey);

    if (!obj.lastDate || d > obj.lastDate) obj.lastDate = d;
    if (d >= cutoff91) obj.count91++;
  });

  const lastRow = sheet.getLastRow();
  const pidLastFirst = sheet.getRange(3, 2, lastRow - 2, 3).getValues(); // B3:D

  const out = pidLastFirst.map(([pid, last, first]) => {
    const key = buildMatchKey(pid, last, first);
    const info = key ? attMap.get(key) : null;

    const lastDate = info ? info.lastDate : null;
    const count91 = info ? info.count91 : 0;

    if (lastDate && lastDate < cutoff12mo) return ["Archive"];
    if (count91 >= 12) return ["Core"];
    if (count91 >= 3) return ["Active"];
    return ["Inactive"];
  });

  sheet.getRange(3, 6, out.length, 1).setValues(out);
}

function performFinalSort_ForSpreadsheet_(ss) {
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() <= 2) return;

  const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const data = range.getValues();

  const order = { "Core": 1, "Active": 2, "Inactive": 3, "Archive": 4 };

  data.sort((a, b) => {
    const guestA = a[4] === "Guest";
    const guestB = b[4] === "Guest";
    if (guestA !== guestB) return guestA ? -1 : 1;

    const aLvl = order[a[5]] || 99;
    const bLvl = order[b[5]] || 99;
    if (aLvl !== bLvl) return aLvl - bLvl;

    const lastA = (a[2] || "").toString().toLowerCase();
    const lastB = (b[2] || "").toString().toLowerCase();
    if (lastA < lastB) return -1;
    if (lastA > lastB) return 1;

    const firstA = (a[3] || "").toString().toLowerCase();
    const firstB = (b[3] || "").toString().toLowerCase();
    if (firstA < firstB) return -1;
    if (firstA > firstB) return 1;

    return 0;
  });

  range.setValues(data);
  range.clearFormat();

  const numRows = range.getNumRows();

  const centerRange = sheet.getRange(3, 5, numRows, 8); // E–L
  centerRange.setHorizontalAlignment("center");
  centerRange.setVerticalAlignment("middle");

  const colBRange = sheet.getRange(3, 2, numRows, 1);
  colBRange.setHorizontalAlignment("left");
  colBRange.setVerticalAlignment("middle");

  const colMRange = sheet.getRange(3, 13, numRows, 1);
  colMRange.setHorizontalAlignment("left");
  colMRange.setVerticalAlignment("middle");

  Logger.log("✅ Final sort and alignment complete (Column B & M left-aligned).");

