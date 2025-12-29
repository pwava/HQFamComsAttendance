/**
 * Calculates and synchronizes two types of monthly average weekly attendance
 * (Sunday Service and Other Events) to an external Google Sheet.
 *
 * It matches calculated results to pre-existing monthly dates in Column B of the
 * destination 'Attendance' tab and inputs the averages into the correct columns.
 */

// --- Global Configuration ---
const CONFIG_TAB_NAME = "Config";
const SERVICE_TAB_NAME = "Sunday Service";
const EVENT_ATTENDANCE_TAB_NAME = "Event Attendance"; // New Source Tab Name
const DESTINATION_TAB_NAME = "Attendance";
const EXTERNAL_SHEET_ID_CELL = "B3";
const DATA_START_ROW = 3; // Data in the destination sheet starts after the 2 header rows


/**
 * Master function to run all synchronization calculations.
 * This is the function you should use for your time-based trigger.
 */
function updateAllAttendanceAverages() {
  Logger.log("--- Starting Full Attendance Update ---");
  
  // 1. Update Sunday Service Average (writes to Column C)
  calculateSundayServiceAverage(); 

  // 2. Update Other Event Average (writes to Column H)
  calculateOtherEventsAverage();
  
  Logger.log("--- Full Attendance Update Complete ---");
}


// --- 1. Function for Sunday Service Attendance (Target Column C) ---

/**
 * Calculates the average weekly (Sunday Service) attendance per month
 * and matches the results to existing dates in Column B of the external 'Attendance' tab,
 * inputting the average into Column C.
 */
function calculateSundayServiceAverage() {
  // Source: Dates in Row 2, Counts in Row 3 (Starting Col B)
  // Target: Destination Column C
  
  Logger.log("Starting Sunday Service Average calculation...");
  
  // Get required sheets and external ID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  const serviceSheet = ss.getSheetByName(SERVICE_TAB_NAME);
  
  if (!configSheet) {
    Logger.log(`Configuration Error: Missing tab named '${CONFIG_TAB_NAME}'.`);
    return;
  }
  const externalSheetId = configSheet.getRange(EXTERNAL_SHEET_ID_CELL).getValue();
  if (!externalSheetId) {
    Logger.log("Configuration Error: External Sheet ID is missing from Config!B3.");
    return;
  }
  if (!serviceSheet) {
    Logger.log(`Configuration Error: Missing tab named '${SERVICE_TAB_NAME}'.`);
    return;
  }
  
  try {
    const lastCol = serviceSheet.getLastColumn();
    if (lastCol < 2) {
      Logger.log("No data found in Sunday Service tab (must start in column B).");
      return;
    }

    // Dates are in Row 2, starting from column 2 (B)
    const dates = serviceSheet.getRange(2, 2, 1, lastCol - 1).getValues()[0];
    // Counts are in Row 3, starting from column 2 (B)
    const counts = serviceSheet.getRange(3, 2, 1, lastCol - 1).getValues()[0];

    // Calculate monthly averages (Sunday remains per-event, which is effectively weekly if you have 1 service per week)
    const monthlyData = processAttendanceData(dates, counts, "Sunday Service");
    
    // Write data to external sheet
    writeAveragesToDestination(externalSheetId, monthlyData, 3, "Sunday Service Average (C)");

  } catch (e) {
    Logger.log(`A critical error occurred in calculateSundayServiceAverage: ${e}`);
  }
}


// --- 2. Function for Other Event Attendance (Target Column H) ---

/**
 * Calculates the average weekly (Other Event) attendance per month
 * and matches the results to existing dates in Column B of the external 'Attendance' tab,
 * inputting the average into Column H.
 */
function calculateOtherEventsAverage() {
  // Source: Dates in Row 2, Counts in Row 4 (Starting Col I) on "Event Attendance" tab
  // Target: Destination Column H
  
  Logger.log("Starting Other Events Average calculation...");
  
  // Get required sheets and external ID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  const eventSheet = ss.getSheetByName(EVENT_ATTENDANCE_TAB_NAME); // <<< PULLING DATA FROM THIS NEW TAB
  
  if (!configSheet) {
    Logger.log(`Configuration Error: Missing tab named '${CONFIG_TAB_NAME}'.`);
    return;
  }
  const externalSheetId = configSheet.getRange(EXTERNAL_SHEET_ID_CELL).getValue();
  if (!externalSheetId) {
    Logger.log("Configuration Error: External Sheet ID is missing from Config!B3.");
    return;
  }
  if (!eventSheet) {
    Logger.log(`Configuration Error: Missing tab named '${EVENT_ATTENDANCE_TAB_NAME}'.`);
    return;
  }

  try {
    const startCol = 9; // Column I (for Dates and Counts)
    const dateRow = 2;   // Dates Row
    const countRow = 4;  // Counts Row

    // Find the rightmost column that has a date in the header row for 'Other Events' (Row 2)
    let lastEventCol = startCol;
    const maxCols = eventSheet.getLastColumn();
    
    for (let c = startCol; c <= maxCols; c++) {
      const cellValue = eventSheet.getRange(dateRow, c).getValue();
      if (cellValue instanceof Date) {
        lastEventCol = c;
      }
    }
    
    if (lastEventCol < startCol) {
      Logger.log(`No 'Other Events' data found in ${EVENT_ATTENDANCE_TAB_NAME} tab in Row 2 (must start in column I with a date).`);
      return;
    }

    const numColumns = lastEventCol - startCol + 1;
    
    const dates = eventSheet.getRange(dateRow, startCol, 1, numColumns).getValues()[0];
    const counts = eventSheet.getRange(countRow, startCol, 1, numColumns).getValues()[0];

    // Calculate monthly averages
    // CHANGE: For Other Events, compute WEEKLY average (divide by unique weeks with events)
    const monthlyData = processAttendanceData(dates, counts, "Other Events", "week");
    
    // Write data to external sheet
    writeAveragesToDestination(externalSheetId, monthlyData, 8, "Other Events Average (H)"); // Column 8 is H

  } catch (e) {
    Logger.log(`A critical error occurred in calculateOtherEventsAverage: ${e}`);
  }
}


// --- 3. Helper Functions ---

/**
 * Processes date and attendance counts into monthly averages.
 * @param {Date[]} dates Array of service dates.
 * @param {number[]} counts Array of attendance counts.
 * @param {string} sourceLabel Label for logging (e.g., "Sunday Service", "Other Events").
 * @param {string} mode (optional) "event" (default) or "week" (divide by unique weeks per month).
 * @returns {Object<string, number>} A map of monthly averages {'YYYY-MM': averageCount}.
 */
function processAttendanceData(dates, counts, sourceLabel, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const useWeekMode = (mode === "week");

  // Stores { 'YYYY-MM': { totalAttendance: N, serviceCount: M, weekSet: {...} } }
  const monthlyAggregates = {}; 
  const monthlyAverages = {};
  let validDataFound = false;
  
  for (let i = 0; i < dates.length; i++) {
    const date = dates[i];
    const count = counts[i];

    if (date instanceof Date && typeof count === 'number' && count > 0) {
      const year = date.getFullYear();
      const month = date.getMonth() + 1; 
      const monthKey = `${year}-${String(month).padStart(2, '0')}`;
      validDataFound = true;

      if (!monthlyAggregates[monthKey]) {
        monthlyAggregates[monthKey] = {
          totalAttendance: 0,
          serviceCount: 0,
          weekSet: {}
        };
      }

      monthlyAggregates[monthKey].totalAttendance += count;

      if (useWeekMode) {
        // Week key: ISO-style year-week (based on spreadsheet timezone)
        const weekKey = Utilities.formatDate(date, tz, "yyyy-ww");
        monthlyAggregates[monthKey].weekSet[weekKey] = true;
      } else {
        monthlyAggregates[monthKey].serviceCount += 1;
      }
    }
  }
  
  if (!validDataFound) {
    Logger.log(`DEBUG (${sourceLabel}): Did not process any valid date/count pairs.`);
    return {};
  }
  
  const generatedKeys = [];
  for (const key in monthlyAggregates) {
    const data = monthlyAggregates[key];

    let divisor = 0;
    if (useWeekMode) {
      divisor = Object.keys(data.weekSet).length;
    } else {
      divisor = data.serviceCount;
    }

    const avg = divisor > 0 ? data.totalAttendance / divisor : 0;
    monthlyAverages[key] = Math.round(avg);
    generatedKeys.push(key);
  }
  
  Logger.log(`DEBUG (${sourceLabel}): Successfully generated monthly keys: ${generatedKeys.join(', ')}`);

  return monthlyAverages;
}

/**
 * Reads destination dates, maps calculated data, clears and writes the averages to the target column.
 * Ensures cells are horizontally and vertically centered, and **does not touch row 15** (to preserve its formula).
 * @param {string} externalSheetId ID of the destination sheet.
 * @param {Object<string, number>} monthlyAverages Map of monthly averages {'YYYY-MM': averageCount}.
 * @param {number} targetColumn The column index (1-based) to write the results (e.g., 3 for C, 8 for H).
 * @param {string} label A descriptive label for logging purposes.
 */
function writeAveragesToDestination(externalSheetId, monthlyAverages, targetColumn, label) {
  let destinationSheet;
  try {
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    destinationSheet = externalSS.getSheetByName(DESTINATION_TAB_NAME);

    if (!destinationSheet) {
      destinationSheet = externalSS.insertSheet(DESTINATION_TAB_NAME);
      Logger.log(`Created new destination tab: '${DESTINATION_TAB_NAME}'.`);
    }

  } catch (e) {
    Logger.log(`Destination Sheet Error: Could not open external sheet or find the tab '${DESTINATION_TAB_NAME}'. Error: ${e.message}`);
    return;
  }
  
  const lastRow = destinationSheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    Logger.log(`Destination sheet has no pre-existing dates in Column B to match against (must start at Row ${DATA_START_ROW}).`);
    return;
  }

  const SKIP_ROW = 15;

  const numRows = lastRow - DATA_START_ROW + 1;
  const dateRange = destinationSheet.getRange(DATA_START_ROW, 2, numRows, 1);
  const destinationDates = dateRange.getValues(); 

  const dateToPositionMap = new Map(); 
  let destinationKeys = [];
  
  for (let i = 0; i < destinationDates.length; i++) {
    const sheetRow = DATA_START_ROW + i;
    const dateCell = destinationDates[i][0];
    if (dateCell instanceof Date) {
      const year = dateCell.getFullYear();
      const month = dateCell.getMonth() + 1; 
      const key = `${year}-${String(month).padStart(2, '0')}`;
      dateToPositionMap.set(key, i);
      destinationKeys.push(key);
    }
  }

  Logger.log(`DEBUG (${label}): Destination monthly keys found in Column B: ${destinationKeys.join(', ')}`);
  
  const hasSkipInside = (SKIP_ROW >= DATA_START_ROW && SKIP_ROW <= lastRow);
  const topBlockRows = hasSkipInside ? Math.max(0, SKIP_ROW - DATA_START_ROW) : numRows;
  const botBlockRows = hasSkipInside ? Math.max(0, lastRow - SKIP_ROW) : 0;
  
  const topValues = topBlockRows > 0 ? Array(topBlockRows).fill(['']) : [];
  const botValues = botBlockRows > 0 ? Array(botBlockRows).fill(['']) : [];

  let monthsUpdated = 0;
  for (const key in monthlyAverages) {
    const avg = monthlyAverages[key];
    if (dateToPositionMap.has(key)) {
      const pos = dateToPositionMap.get(key);
      const sheetRow = DATA_START_ROW + pos;
      if (hasSkipInside && sheetRow === SKIP_ROW) {
        continue;
      } else if (!hasSkipInside) {
        topValues[pos] = [avg];
        monthsUpdated++;
      } else if (sheetRow < SKIP_ROW) {
        const topIndex = sheetRow - DATA_START_ROW;
        topValues[topIndex] = [avg];
        monthsUpdated++;
      } else if (sheetRow > SKIP_ROW) {
        const botIndex = sheetRow - SKIP_ROW - 1;
        botValues[botIndex] = [avg];
        monthsUpdated++;
      }
    } else {
      Logger.log(`WARNING (${label}): Calculated key '${key}' did not match any date in destination Column B. Skipping.`);
    }
  }

  if (!hasSkipInside) {
    const rangeAll = destinationSheet.getRange(DATA_START_ROW, targetColumn, numRows, 1);
    rangeAll.setValues(topValues);
    rangeAll.setHorizontalAlignment('center');
    rangeAll.setVerticalAlignment('middle');
  } else {
    if (topBlockRows > 0) {
      const topRange = destinationSheet.getRange(DATA_START_ROW, targetColumn, topBlockRows, 1);
      topRange.setValues(topValues);
      topRange.setHorizontalAlignment('center');
      topRange.setVerticalAlignment('middle');
    }
    if (botBlockRows > 0) {
      const botRange = destinationSheet.getRange(SKIP_ROW + 1, targetColumn, botBlockRows, 1);
      botRange.setValues(botValues);
      botRange.setHorizontalAlignment('center');
      botRange.setVerticalAlignment('middle');
    }
  }

  if (monthsUpdated > 0) {
    Logger.log(`SUCCESS: Successfully updated ${monthsUpdated} month(s) for ${label}.`);
  } else {
    Logger.log(`WARNING: No calculated averages matched the dates found in Column B for ${label}. Target column values were reset to blank (excluding row ${SKIP_ROW}).`);
  }
}
