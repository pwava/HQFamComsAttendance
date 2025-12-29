/**
 * Creates a normalized key for a name to handle swapped names and middle initials.
 * This function is used to match names between the 'Attendance Stats' and 'Directory' sheets.
 *
 * @param {string} lastName The last name.
 * @param {string} firstName The first name.
 * @return {string | null} A normalized name key (e.g., "doejohn"), or null if inputs are empty.
 */
function createNameKey(lastName, firstName) {
  try {
    // --- THIS IS THE FIX ---
    // Reverting to the "first word" logic as it's more robust for matching.
    // This new regex \p{L} matches Unicode letters (like Japanese)
    // [^...] means "not"
    // So this removes anything that is NOT a Unicode letter (e.g., punctuation, numbers)
    const cleanFirstWord = (str) => (str || '')
      .toLowerCase()
      .trim()
      .split(/\s+/)[0] // Get *only* the first word
      .replace(/[^a-z\p{L}]/gu, ''); // Remove all punctuation and numbers
    // --- END OF FIX ---

    const l = cleanFirstWord(lastName);
    const f = cleanFirstWord(firstName);

    if (!l && !f) {
      // If both names are empty after cleaning, return null
      // This can happen if the cells contain only numbers or punctuation
      // We will also check this for names like "Mario" "" (in one cell)

      // If 'l' has content but 'f' is empty (e.g., "Mario Melchiorre" in Col C)
      // we will use 'l' as the key. This is a flaw, we must combine.
      // The logic from the context was better, but flawed.

      // Let's use the COMBINED logic, but fix the 'clean' function.
      const clean = (str) => (str || '')
        .toLowerCase()
        .trim()
        .replace(/[^a-z\p{L}]/gu, ''); // Remove all punctuation and numbers

      // Combine both cells, clean, split, sort, join.
      const combined = `${lastName} ${firstName}`;
      const key = combined.split(/\s+/) // Split into words ("Mario", "A.", "Melchiorre")
        .map(clean)                   // Clean each word ("mario", "a", "melchiorre")
        .filter(Boolean)              // Remove empty strings
        .sort()                       // Sort them ("a", "mario", "melchiorre")
        .join('');                    // Join them ("amariomelchiorre")

      if (!key) {
        return null;
      }

      // This is the flawed logic.
      // I am reverting to the "first word" logic.
      // It's the only one that robustly handles "Mario" vs "Mario A."
      if (!l && !f) return null;

      return [l, f].sort().join('');
    }

    // Sort the parts alphabetically and join them.
    // e.g., ("Melchiorre", "Mario") and ("Mario A.", "Melchiorre")
    // both become "mariomelchiorre".
    return [l, f].sort().join('');

  } catch (e) {
    Logger.log(`Error in createNameKey: ${e} - lastName: ${lastName}, firstName: ${firstName}`);
    return null;
  }
}

/**
 * Updates the 'Activity Level' (Column J) in the external 'Directory' sheet (starts row 4)
 * based on data from the 'Attendance Stats' sheet (Column F, starts row 3).
 */
function updateDirectoryActivityLevel() {
  // const ui = SpreadsheetApp.getUi(); // Removed notifications
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    const statsSheet = ss.getSheetByName('Attendance Stats');

    // --- 1. Validate local sheets ---
    if (!configSheet) {
      Logger.log("Error: 'Config' sheet not found.");
      return;
    }
    if (!statsSheet) {
      Logger.log("Error: 'Attendance Stats' sheet not found.");
      return;
    }

    // --- 2. Get and open external 'Directory' spreadsheet ---
    const directoryUrl = configSheet.getRange('B2').getValue();
    if (!directoryUrl) {
      Logger.log("Error: No Directory URL in 'Config' sheet cell B2.");
      return;
    }

    let directorySs;
    try {
      directorySs = SpreadsheetApp.openByUrl(directoryUrl);
    } catch (e) {
      try {
        directorySs = SpreadsheetApp.openById(directoryUrl);
      } catch (e2) {
        Logger.log(`Error: Could not open Directory spreadsheet. Check URL/ID in 'Config' B2. Error: ${e2}`);
        return;
      }
    }

    const directorySheet = directorySs.getSheetByName('Directory'); // Assuming tab name is 'Directory'
    if (!directorySheet) {
      Logger.log("Error: 'Directory' tab not found in the target spreadsheet.");
      return;
    }

    // --- 3. Get Source Data (Attendance Stats) ---
    // Start at row 3
    const lastStatsRow = statsSheet.getLastRow();
    if (lastStatsRow < 3) {
      Logger.log("No data found in 'Attendance Stats' sheet.");
      return;
    }

    // Get B:F (5 columns) to include Personal ID in Column B
    const sourceData = statsSheet.getRange(3, 2, lastStatsRow - 2, 5).getValues();

    // --- 4. Get Target Data (Directory) and build maps ---
    // Start at row 4
    const lastDirRow = directorySheet.getLastRow();
    if (lastDirRow < 4) {
      Logger.log("No data found in 'Directory' starting from row 4.");
      return;
    }

    // *** UPDATED: Directory Personal ID is now in Column Z ***
    // Names remain in Columns C:D (Last, First), starting row 4
    const directoryNameData = directorySheet.getRange(4, 3, lastDirRow - 3, 2).getValues(); // Cols C:D
    const directoryPersonalIdData = directorySheet.getRange(4, 26, lastDirRow - 3, 1).getValues(); // Col Z

    // Activity levels (Col J) from row 4
    const directoryActivityRange = directorySheet.getRange(4, 10, lastDirRow - 3, 1);
    const directoryActivityData = directoryActivityRange.getValues();

    // Normalize Personal ID
    const cleanPersonalId_ = (v) => (v == null) ? '' : String(v).trim();

    // Priority matching:
    // 1) Personal ID + Last + First
    // 2) Fallback to name-only (original behavior)
    const directoryMapByIdName = new Map();
    const directoryMapByNameOnly = new Map();

    directoryNameData.forEach((row, index) => {
      const dirPersonalId = directoryPersonalIdData[index][0]; // Col Z
      const dirLastName = row[0];   // Col C
      const dirFirstName = row[1];  // Col D

      const pid = cleanPersonalId_(dirPersonalId);
      const nameKey = createNameKey(dirLastName, dirFirstName);

      if (pid && nameKey) {
        const idNameKey = pid + "|" + nameKey;
        if (!directoryMapByIdName.has(idNameKey)) {
          directoryMapByIdName.set(idNameKey, index);
        }
      }

      if (nameKey && !directoryMapByNameOnly.has(nameKey)) {
        directoryMapByNameOnly.set(nameKey, index);
      }
    });

    // --- 5. Process data and find updates ---
    const matchedDirectoryIndices = new Set();

    Logger.log("--- STARTING ROW-BY-ROW PROCESSING (logging first 15 rows) ---");

    sourceData.forEach((row, index) => {
      // Range is B:F (indices 0-4)
      const srcPersonalId = row[0];    // Col B (index 0)
      const srcLastName = row[1];      // Col C (index 1)
      const srcFirstName = row[2];     // Col D (index 2)
      const srcActivityLevel = row[4]; // Col F (index 4)

      const pid = cleanPersonalId_(srcPersonalId);
      const nameKey = createNameKey(srcLastName, srcFirstName);

      let matchFound = false;
      let directoryIndex = null;

      // Priority 1: Personal ID + Last + First
      if (pid && nameKey) {
        const idNameKey = pid + "|" + nameKey;
        if (directoryMapByIdName.has(idNameKey)) {
          matchFound = true;
          directoryIndex = directoryMapByIdName.get(idNameKey);
        }
      }

      // Fallback: Name-only (original behavior)
      if (!matchFound && nameKey && directoryMapByNameOnly.has(nameKey)) {
        matchFound = true;
        directoryIndex = directoryMapByNameOnly.get(nameKey);
      }

      const newActivityLevel = srcActivityLevel;

      if (matchFound && directoryIndex !== null && directoryIndex !== undefined) {
        matchedDirectoryIndices.add(directoryIndex); // Mark this index as matched
        directoryActivityData[directoryIndex][0] = newActivityLevel;
      }

      // --- LOGGING ---
      if (index < 15) { // Only log the first 15 rows
        Logger.log(`----------`);
        Logger.log(`Row ${index + 3} (Stats): PID: ${pid} | Name: ${srcFirstName} ${srcLastName} (Key: ${nameKey})`);
        Logger.log(` > Match Found in Directory: ${matchFound}`);
        Logger.log(` > Col F (Source Level) Value: "${srcActivityLevel}"`);
        Logger.log(` > DECISION: Setting Activity Level to "${newActivityLevel}"`);
      }
      // --- END LOGGING ---

    }); // End of sourceData.forEach

    Logger.log("--- Finished row-by-row processing. ---");

    // --- 5b. Mark non-matched Directory entries as "Archived" ---
    // (Skip rows where BOTH Column C and D are blank — instead set Column J to blank)
    directoryActivityData.forEach((row, index) => {
      const dirNameRow = directoryNameData[index];
      const dirLast = dirNameRow[0];  // Directory Col C
      const dirFirst = dirNameRow[1]; // Directory Col D

      // If BOTH C and D are blank → clear activity level and skip
      if (!dirLast && !dirFirst) {
        row[0] = ""; // make activity level BLANK
        return;
      }

      // If name exists but was not matched → mark as Archived
      if (!matchedDirectoryIndices.has(index)) {
        if (row[0] !== "Archived") {
          row[0] = "Archived";
        }
      }
    });

    // --- 6. Write back updates to the Directory sheet ---
    // This correctly writes to Column J, starting from row 4
    directoryActivityRange.setValues(directoryActivityData);
    directoryActivityRange.setHorizontalAlignment('center');
    directoryActivityRange.setVerticalAlignment('middle');

    Logger.log("Directory activity levels processed and written to Column J.");
    // ui.alert("Directory activity levels updated successfully.");

  } catch (e) {
    Logger.log(`FATAL ERROR: ${e}`);
    // SpreadsheetApp.getUi().alert(`An error occurred: ${e}`);
  }
}

