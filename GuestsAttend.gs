/**
 * Normalize Personal ID for matching (keep letters+numbers only).
 */
function normalizePersonalId(id) {
  if (!id) return '';
  return String(id).toLowerCase().trim().replace(/[^a-z0-9]/g, '');
}

/**
 * Build a member key used for matching:
 * PersonalID + Last + First
 *
 * - If Personal ID exists: "pid|lastname|firstname"
 * - If Personal ID is blank: fall back to original name-only behavior:
 *   - both last & first: "lastname,firstname"
 *   - only last: "lastname"
 *   - only first: "firstname"
 */
function buildGuestKey(personalId, lastName, firstName) {
  const pid = normalizePersonalId(personalId);
  const ln = String(lastName || "").trim().toLowerCase();
  const fn = String(firstName || "").trim().toLowerCase();

  if (pid) return pid + "|" + ln + "|" + fn;

  if (ln && fn) return ln + "," + fn;
  if (ln) return ln;
  if (fn) return fn;
  return "";
}

/**
 * Main function to be run manually to update the "Guests" tab.
 * This script gathers data from "Sunday Service", "Event Attendance",
 * and an external "Directory" sheet to populate guest information.
 * This function CLEARS all data from B4:G and rewrites it.
 */
function updateGuestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get Directory Data
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);

    // 2. Get Sunday Service Data
    const serviceMap = getServiceData(serviceSheet);

    // 3. Get Event Attendance (Community Intro) Data
    const introMap = getIntroData(eventSheet);

    // 3b. Get Pastoral Check-In data from Attendance Log (for column F fallback)
    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 4. Get Unique Guest List
    const uniqueGuests = getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet);

    // 5. Process and Prepare Data for Writing
    const finalData = [];

    // Filter out anyone already in Directory (they are NOT guests)
    // Matching is PersonalID + Last + First (Directory Personal ID is Column Z)
    const filteredGuests = [];
    for (const guest of uniqueGuests.values()) {
      const pid = normalizePersonalId(guest.personalId);
      const ln = String(guest.lastName || "").trim().toLowerCase();
      const fn = String(guest.firstName || "").trim().toLowerCase();
      let inDirectory = false;

      if (pid) {
        const dirKey = buildGuestKey(pid, ln, fn);
        if (directoryMap.has(dirKey)) {
          inDirectory = true;
        }
      }

      if (!inDirectory) {
        filteredGuests.push(guest);
      }
    }

    // Sort guests alphabetically by last name, then first name.
    const sortedGuests = filteredGuests.sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      if (a.firstName < b.firstName) return -1;
      if (a.firstName > b.firstName) return 1;
      return 0;
    });

    // Build the final array for the "Guests" sheet.
    for (const guest of sortedGuests) {
      const key = buildGuestKey(guest.personalId, guest.lastName, guest.firstName);

      const serviceDate = key ? (serviceMap.get(key) || "") : "";

      // Column F (Intro) = Community Intro first, else Pastoral Check-In from Attendance Log
      let introDate = key ? (introMap.get(key) || "") : "";
      if (!introDate && key && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      // Registration date from Directory (will normally be blank for true guests,
      // because we filtered out those already in Directory)
      let regDate = "";
      if (normalizePersonalId(guest.personalId)) {
        const dirKey = buildGuestKey(guest.personalId, guest.lastName, guest.firstName);
        regDate = directoryMap.get(dirKey) || "";
      }

      // Match the column order: B, C, D, E, F, G
      // B = Personal ID (replaces Full Name)
      finalData.push([
        guest.personalId || "",
        guest.lastName,
        guest.firstName,
        serviceDate,
        introDate,
        regDate
      ]);
    }

    // 6. Write Data to Guests Sheet
    const startRow = 4;
    const numRows = finalData.length;

    // Clear old data from row 4 downwards, columns B-G
    if (guestsSheet.getLastRow() >= startRow) {
      guestsSheet.getRange(startRow, 2, guestsSheet.getLastRow() - startRow + 1, 6).clearContent();
    }

    // Write new data if any exists
    if (numRows > 0) {
      guestsSheet.getRange(startRow, 2, numRows, 6).setValues(finalData);
      guestsSheet.getRange(startRow, 5, numRows, 3).setNumberFormat("MM-dd-yy");

      // Copy Registration Date (G) into Column H when G has a date
      const regCol = guestsSheet.getRange(startRow, 7, numRows, 1).getValues(); // G
      const targetH = [];

      for (let i = 0; i < regCol.length; i++) {
        const val = regCol[i][0];
        if (val instanceof Date) {
          targetH.push([val]);
        } else {
          targetH.push([""]);
        }
      }

      guestsSheet.getRange(startRow, 8, numRows, 1).setValues(targetH); // H
      guestsSheet.getRange(startRow, 8, numRows, 1).setNumberFormat("MM-dd-yy");
    }

    Logger.log("Guest data updated successfully.");

  } catch (e) {
    Logger.log("Error in updateGuestData: " + e);
  }
}

/**
 * Finds new guests and appends them to the "Guests" tab WITH their dates.
 * This function does NOT sort the sheet or remove old guests.
 */
function addNewGuests() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get all guests who are ALREADY in the "Guests" tab
    const existingGuestsSet = new Set();
    const startRow = 4;
    if (guestsSheet.getLastRow() >= startRow) {
      // B:D = PersonalID, Last, First
      const existing = guestsSheet.getRange(startRow, 2, guestsSheet.getLastRow() - startRow + 1, 3).getValues(); // B:D
      for (const row of existing) {
        const personalId = String(row[0] || "").trim();
        const lastName = String(row[1] || "").trim();
        const firstName = String(row[2] || "").trim();
        const key = buildGuestKey(personalId, lastName, firstName);
        if (key) {
          existingGuestsSet.add(key);
        }
      }
    }

    // 2. Get ALL unique guests from the source tabs
    const allGuestsMap = getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet);

    // 3. Find only the NEW guests
    const newGuests = [];
    for (const [key, guest] of allGuestsMap.entries()) {
      if (!existingGuestsSet.has(key)) {
        newGuests.push(guest);
      }
    }

    // 4. If no new guests, stop here
    if (newGuests.length === 0) {
      Logger.log("No new guests found.");
      return;
    }

    // 5. Get all date information to populate for the new guests
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);
    const serviceMap = getServiceData(serviceSheet);
    const introMap = getIntroData(eventSheet);

    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 6. Sort new guests and format them for the sheet
    const finalData = [];
    newGuests.sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      if (a.firstName < b.firstName) return -1;
      if (a.firstName > b.firstName) return 1;
      return 0;
    });

    for (const guest of newGuests) {
      const pid = normalizePersonalId(guest.personalId);
      const ln = String(guest.lastName || "").trim().toLowerCase();
      const fn = String(guest.firstName || "").trim().toLowerCase();
      let inDirectory = false;

      if (pid) {
        const dirKey = buildGuestKey(pid, ln, fn);
        if (directoryMap.has(dirKey)) {
          inDirectory = true;
        }
      }

      // Skip if already in Directory (not a guest)
      if (inDirectory) {
        continue;
      }

      const key = buildGuestKey(guest.personalId, guest.lastName, guest.firstName);

      const serviceDate = key ? (serviceMap.get(key) || "") : "";

      let introDate = key ? (introMap.get(key) || "") : "";
      if (!introDate && key && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      let regDate = "";
      if (pid) {
        const dirKey2 = buildGuestKey(guest.personalId, guest.lastName, guest.firstName);
        regDate = directoryMap.get(dirKey2) || "";
      }

      // B: Personal ID, C: Last, D: First, E: Service, F: Intro, G: Reg
      finalData.push([
        guest.personalId || "",
        guest.lastName,
        guest.firstName,
        serviceDate,
        introDate,
        regDate
      ]);
    }

    // 7. Write the new guests starting at the next available blank row.
    const existingLastRow = guestsSheet.getLastRow();
    const dataStartRow = 4;
    const blankRows = [];

    if (existingLastRow >= dataStartRow) {
      // Check C:D (Last, First) for blanks (original behavior)
      const nameValues = guestsSheet.getRange(dataStartRow, 3, existingLastRow - dataStartRow + 1, 2).getValues(); // C:D
      for (let i = 0; i < nameValues.length; i++) {
        const ln = String(nameValues[i][0]).trim();
        const fn = String(nameValues[i][1]).trim();
        if (!ln && !fn) {
          blankRows.push(dataStartRow + i);
        }
      }
    }

    let dataIndex = 0;

    // Fill existing empty rows
    for (let i = 0; i < blankRows.length && dataIndex < finalData.length; i++, dataIndex++) {
      const rowIndex = blankRows[i];
      guestsSheet.getRange(rowIndex, 2, 1, 6).setValues([finalData[dataIndex]]);
      guestsSheet.getRange(rowIndex, 5, 1, 3).setNumberFormat("MM-dd-yy");
    }

    // Append remaining new guests
    if (dataIndex < finalData.length) {
      const remaining = finalData.slice(dataIndex);
      const appendStartRow = Math.max(existingLastRow + 1, dataStartRow);
      const newRowsRange = guestsSheet.getRange(appendStartRow, 2, remaining.length, 6);
      newRowsRange.setValues(remaining);
      newRowsRange.offset(0, 3, remaining.length, 3).setNumberFormat("MM-dd-yy");
    }

    Logger.log("Added " + finalData.length + " new guests with their dates.");

  } catch (e) {
    Logger.log("Error in addNewGuests: " + e);
  }
}

/**
 * Updates blank dates (E, F, G) for guests already in the "Guests" tab.
 * This function does NOT add or remove rows.
 */
function updateExistingGuestDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get all date data from all sources
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);
    const serviceMap = getServiceData(serviceSheet);
    const introMap = getIntroData(eventSheet);

    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 2. Get the current data from the "Guests" sheet
    const startRow = 4;
    const lastRow = guestsSheet.getLastRow();
    if (lastRow < startRow) {
      Logger.log("No guests to update.");
      return;
    }

    const numRows = lastRow - startRow + 1;
    // B:G = PersonalID, Last, First, Service, Intro, Reg
    const dataRange = guestsSheet.getRange(startRow, 2, numRows, 6); // B:G
    const values = dataRange.getValues();

    let updatesMade = 0;
    const datesToWrite = [];

    // 3. Loop through each guest and fill in blank dates
    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const personalId = String(row[0] || "").trim(); // B
      const lastName = String(row[1] || "").trim();   // C
      const firstName = String(row[2] || "").trim();  // D

      let serviceDate = row[3]; // E
      let introDate = row[4];   // F
      let regDate = row[5];     // G

      if (lastName || firstName || personalId) {
        const key = buildGuestKey(personalId, lastName, firstName);

        if (key) {
          if (!serviceDate && serviceMap.has(key)) {
            serviceDate = serviceMap.get(key);
            updatesMade++;
          }

          if (!introDate) {
            if (introMap.has(key)) {
              introDate = introMap.get(key);
              updatesMade++;
            } else if (pastoralIntroMap.has(key)) {
              introDate = pastoralIntroMap.get(key);
              updatesMade++;
            }
          }

          // Registration date from Directory uses PersonalID+Last+First
          if (!regDate && normalizePersonalId(personalId)) {
            const dirKey = buildGuestKey(personalId, lastName, firstName);
            if (directoryMap.has(dirKey)) {
              regDate = directoryMap.get(dirKey);
              updatesMade++;
            }
          }
        }
      }

      datesToWrite.push([personalId, lastName, firstName, serviceDate, introDate, regDate]);
    }

    // 4. Write updated rows (only if something actually changed)
    if (updatesMade > 0) {
      dataRange.setValues(datesToWrite);
      guestsSheet.getRange(startRow, 5, numRows, 3).setNumberFormat("MM-dd-yy");
      Logger.log("Updated " + updatesMade + " blank dates for existing guests.");
    } else {
      Logger.log("No blank dates found to update.");
    }

    // 5. ALWAYS: copy Registration (G) into Full (H) when G has a date
    const regCol = guestsSheet.getRange(startRow, 7, numRows, 1).getValues(); // G
    const hValues = [];

    for (let i = 0; i < regCol.length; i++) {
      const val = regCol[i][0];
      if (val instanceof Date) {
        hValues.push([val]);
      } else {
        hValues.push([""]);
      }
    }

    guestsSheet.getRange(startRow, 8, numRows, 1).setValues(hValues); // H
    guestsSheet.getRange(startRow, 8, numRows, 1).setNumberFormat("MM-dd-yy");

  } catch (e) {
    Logger.log("Error in updateExistingGuestDates: " + e);
  }
}

/**
 * Gets registration data from the external Directory sheet.
 * Directory Personal ID is column Z.
 *
 * @param {string} sheetId The ID of the external Google Sheet.
 * @returns {Map<string, Date>} A Map where key is "pid|lastname|firstname"
 * and value is the registration Date object.
 */
function getDirectoryData(sheetId) {
  const directoryMap = new Map();
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Directory");
    if (!sheet) {
      Logger.log("Error: 'Directory' tab not found in external sheet.");
      return directoryMap;
    }

    // Data starts from row 2
    // Read C:Z so we can get Last (C), First (D), RegDate (U), PersonalID (Z)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return directoryMap;

    const data = sheet.getRange(2, 3, lastRow - 1, 24).getValues(); // C to Z

    for (const row of data) {
      const lastName = String(row[0] || "").trim();   // C
      const firstName = String(row[1] || "").trim();  // D
      const regDate = row[18];                        // U
      const personalId = row[23];                     // Z

      const pid = normalizePersonalId(personalId);
      if (pid && (firstName || lastName) && regDate instanceof Date) {
        const key = buildGuestKey(pid, lastName, firstName);
        if (!directoryMap.has(key)) {
          directoryMap.set(key, regDate);
        }
      }
    }
  } catch (e) {
    Logger.log("Error accessing Directory sheet: " + e);
    throw new Error("Error accessing Directory sheet. Check ID and permissions. " + e.message);
  }
  return directoryMap;
}

/**
 * Gets the first service date for all GUESTS from the "Sunday Service" sheet.
 * Uses Column H = "Guest" and allows first-name-only or last-name-only.
 * Also uses Personal ID in Column B when present.
 *
 * @param {Sheet} sheet The "Sunday Service" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is buildGuestKey(personalId,last,first)
 * and value is the first service Date object.
 */
function getServiceData(sheet) {
  const serviceMap = new Map();
  const values = sheet.getDataRange().getValues();

  // Get service dates from row 2, starting column I (index 8)
  const serviceDates = values[1].slice(8);

  // Data starts from row 4 (index 3)
  for (let i = 3; i < values.length; i++) {
    const row = values[i];
    const personalId = String(row[1] || "").trim();   // Column B (index 1)
    const lastName = String(row[2] || "").trim();     // Column C (index 2)
    const firstName = String(row[3] || "").trim();    // Column D (index 3)
    const status = String(row[7] || "").trim();       // Column H (index 7)

    if ((firstName || lastName) && status === "Guest") {
      const key = buildGuestKey(personalId, lastName, firstName);
      if (!key) continue;

      if (!serviceMap.has(key)) {
        const attendance = row.slice(8);
        for (let j = 0; j < attendance.length; j++) {
          if (attendance[j] === true) {
            if (serviceDates[j] instanceof Date) {
              serviceMap.set(key, serviceDates[j]);
              break;
            }
          }
        }
      }
    }
  }
  return serviceMap;
}

/**
 * Gets the first "Intro/Orientation" date for all GUESTS
 * from the "Event Attendance" sheet.
 * Uses Column H = "Guest" and allows first-name-only or last-name-only.
 * Also uses Personal ID in Column B when present.
 *
 * @param {Sheet} sheet The "Event Attendance" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is buildGuestKey(personalId,last,first)
 * and value is the event Date object.
 */
function getIntroData(sheet) {
  const introMap = new Map();
  const values = sheet.getDataRange().getValues();

  const eventDates = values[1].slice(8); // Row 2
  const eventNames = values[2].slice(8); // Row 3

  // Data starts from row 5 (index 4)
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const personalId = String(row[1] || "").trim();  // Column B (index 1)
    const lastName = String(row[2] || "").trim();    // Column C (index 2)
    const firstName = String(row[3] || "").trim();   // Column D (index 3)
    const status = String(row[7] || "").trim();      // Column H (index 7)

    if ((firstName || lastName) && status === "Guest") {
      const key = buildGuestKey(personalId, lastName, firstName);
      if (!key) continue;

      if (!introMap.has(key)) {
        const attendance = row.slice(8);
        for (let j = 0; j < attendance.length; j++) {
          if (attendance[j] === true) {
            const eventName = String(eventNames[j] || "").toLowerCase();

            if (
              eventName.includes("community intro") ||
              eventName.includes("orientation") ||
              eventName.includes("orient")
            ) {
              const eventDate = eventDates[j];
              if (eventDate instanceof Date) {
                introMap.set(key, eventDate);
                break;
              }
            }
          }
        }
      }
    }
  }
  return introMap;
}

/**
 * Gets the first "Pastoral check -In" date for all names
 * from the "Attendance Log" sheet.
 *
 * Uses:
 * - Column B: Personal ID
 * - Column C: Last Name
 * - Column D: First Name
 * - Column B (Date) is still index 1 in this sheet as per your original script
 * - Column F: Event name (must be "Pastoral check -In")
 *
 * @param {Sheet} sheet The "Attendance Log" sheet.
 * @returns {Map<string, Date>} A Map where key is buildGuestKey(personalId,last,first)
 * and value is the first Pastoral Check-In Date object.
 */
function getPastoralIntroData(sheet) {
  const pastoralMap = new Map();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const personalId = String(row[1] || "").trim();     // Column B (index 1)
    const lastName = String(row[2] || "").trim();       // Column C (index 2)
    const firstName = String(row[3] || "").trim();      // Column D (index 3)
    const dateVal = row[1];                             // Column B (index 1) (original)
    const eventNameRaw = String(row[5] || "").trim();   // Column F (index 5)

    if (!firstName && !lastName) {
      continue;
    }

    const eventName = eventNameRaw.toLowerCase();
    if (eventName === "pastoral check -in") {
      if (dateVal instanceof Date) {
        const key = buildGuestKey(personalId, lastName, firstName);
        if (!key) continue;

        if (!pastoralMap.has(key)) {
          pastoralMap.set(key, dateVal);
        } else {
          const existing = pastoralMap.get(key);
          if (dateVal < existing) {
            pastoralMap.set(key, dateVal);
          }
        }
      }
    }
  }

  return pastoralMap;
}

/**
 * Compiles a unique list of GUESTS from:
 * - Sunday Service (Column H = "Guest", with attendance)
 * - Event Attendance (Column H = "Guest", with attendance)
 * - Attendance Log: "Pastoral check -In" rows
 *
 * Uses Personal ID from Column B of each source sheet when present.
 *
 * @returns {Map<string, Object>} key = buildGuestKey(personalId,last,first)
 * value = { personalId, firstName, lastName }
 */
function getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet) {
  const guests = new Map();

  // --- 1. Guests from Sunday Service WITH attendance ---
  const serviceValues = serviceSheet.getDataRange().getValues();
  if (serviceValues.length >= 4) {
    for (let i = 3; i < serviceValues.length; i++) {
      const row = serviceValues[i];
      const personalId = String(row[1] || "").trim(); // Col B
      const lastName = String(row[2] || "").trim();   // Col C
      const firstName = String(row[3] || "").trim();  // Col D
      const status = String(row[7] || "").trim();     // Col H

      if ((firstName || lastName) && status === "Guest") {
        const attendance = row.slice(8);
        const hasAttendance = attendance.some(v => v === true);

        if (hasAttendance) {
          const key = buildGuestKey(personalId, lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { personalId: personalId, firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  // --- 2. Guests from Event Attendance WITH attendance ---
  const eventValues = eventSheet.getDataRange().getValues();
  if (eventValues.length >= 5) {
    for (let i = 4; i < eventValues.length; i++) {
      const row = eventValues[i];
      const personalId = String(row[1] || "").trim(); // Col B
      const lastName = String(row[2] || "").trim();   // Col C
      const firstName = String(row[3] || "").trim();  // Col D
      const status = String(row[7] || "").trim();     // Col H

      if ((firstName || lastName) && status === "Guest") {
        const attendance = row.slice(8);
        const hasAttendance = attendance.some(v => v === true);

        if (hasAttendance) {
          const key = buildGuestKey(personalId, lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { personalId: personalId, firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  // --- 3. Guest candidates from Pastoral Check-In (Attendance Log) ---
  if (attendanceLogSheet) {
    const logValues = attendanceLogSheet.getDataRange().getValues();
    if (logValues.length >= 2) {
      for (let i = 1; i < logValues.length; i++) {
        const row = logValues[i];
        const personalId = String(row[1] || "").trim();  // Col B
        const lastName = String(row[2] || "").trim();    // Col C
        const firstName = String(row[3] || "").trim();   // Col D
        const eventNameRaw = String(row[5] || "").trim(); // Col F

        if (!firstName && !lastName) {
          continue;
        }

        const eventName = eventNameRaw.toLowerCase();
        if (eventName === "pastoral check -in") {
          const key = buildGuestKey(personalId, lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { personalId: personalId, firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  return guests;
}

