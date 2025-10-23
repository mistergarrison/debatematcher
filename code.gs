/**
 * Automated Debate Pairing Tool
 * Google Apps Script
 *
 * @fileoverview A tool for automating the pairing of debate matches (TP and LD) within a Google Sheet.
 * All logic is contained in this single file as required by PRD 6.
 */

// =============================================================================
// Configuration & Constants
// =============================================================================

const CONFIG = {
  SHEET_NAMES: {
    AVAILABILITY: 'Availability',
    SUMMARY: 'Match Summary',
    DEBATERS: 'Debaters',
    JUDGES: 'Judges',
    ROOMS: 'Rooms',
    // Hidden sheet for data aggregation (PRD 5.2). The system is designed to only use this specific sheet.
    AGGREGATE_HISTORY: 'AGGREGATE_HISTORY_DO_NOT_EDIT'
  },
  DEBATE_TYPES: {
    TP: 'TP',
    LD: 'LD'
  },
  RSVP_OPTIONS: ["Yes", "No", "Not responded"],
  ROLES: {
    AFF: 'Aff',
    NEG: 'Neg',
    BYE: 'BYE',
    IRONMAN_SUFFIX: '(IRONMAN)'
  },
  STYLES: {
    HEADER_BG: '#4a86e8',
    HEADER_FONT_COLOR: '#ffffff',
    ERROR_HIGHLIGHT: '#f4c7c3', // Light red for critical errors
    WARNING_HIGHLIGHT: '#fff2cc', // Light yellow for duplicates in match sheets
    FONT: 'Arial'
  },
  // Required order for permanent tabs (PRD 5.1)
  PERMANENT_TABS_ORDER: ['Availability', 'Match Summary', 'Debaters', 'Judges', 'Rooms'],
};

// =============================================================================
// Menu & Triggers
// =============================================================================

/**
 * Adds the custom menu when the spreadsheet opens (PRD 5.3).
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.createMenu('Club Admin')
      .addItem('Generate LD Matches (2 Rounds)', 'generateLdMatches')
      .addItem('Generate TP Matches', 'generateTpMatches')
      .addSeparator()
      .addItem('Clear RSVPs for Next Week', 'clearRsvps')
      .addSeparator()
      .addItem('Initialize Sheet (Setup)', 'initializeSheet')
      // Added for maintenance: ensures formulas and formatting are correctly applied.
      .addItem('Refresh Formulas & Formatting', 'forceUpdateSheets')
      .addToUi();
  }

  // Ensure formatting and formulas are applied on open.
  try {
    // We force update on open to clear potential #REF errors if the script logic (formulas) changed or if data expansion is blocked.
    forceUpdateSheets(true);
  } catch (e) {
    // Silently fail if initialization hasn't happened yet
    Logger.log("Could not apply formatting/formulas onOpen (sheet might not be initialized): " + e.message);
  }
}

/**
 * Helper function to forcibly update formulas and formatting. Useful if sheets get corrupted or after script updates.
 * @param {boolean} [isAutomatic=false] - True if called automatically (e.g., onOpen), false if called manually by user.
 */
function forceUpdateSheets(isAutomatic = false) {
  try {
    // Force update formulas to clear potential #REF errors and apply latest logic.
    applyFormulas(true);
    formatAllSheets();
    applyDataValidations();
    applyAllConditionalFormatting();
    // Only show toast if called manually
    if (!isAutomatic) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) ss.toast("Formulas and formatting refreshed.", "Maintenance", 3);
    }
  } catch (e) {
    Logger.log("Error during forced update: " + e.message);
    // Only show alert if called manually and an error occurred
    if (!isAutomatic) {
      const ui = SpreadsheetApp.getUi();
      if (ui) {
        ui.alert("Error during forced update: " + e.message);
      }
    }
  }
}


// =============================================================================
// Initialization Functions
// =============================================================================

/**
 * Initializes the spreadsheet with required tabs, headers, sample data,
 * formulas, and formatting (PRD 4.1, 5.3).
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if already initialized (PRD 4.1)
  // We check for the existence of the Debaters tab as a proxy for initialization.
  if (ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS)) {
    ui.alert('Initialization Error', 'The sheet appears to be already initialized. Setup aborted to prevent data loss.', ui.ButtonSet.OK);
    return;
  }

  // Create permanent tabs
  createTab(CONFIG.SHEET_NAMES.DEBATERS, getDebaterHeaders());
  createTab(CONFIG.SHEET_NAMES.JUDGES, getJudgeHeaders());
  createTab(CONFIG.SHEET_NAMES.ROOMS, getRoomHeaders());
  createTab(CONFIG.SHEET_NAMES.AVAILABILITY, getAvailabilityHeaders());
  createTab(CONFIG.SHEET_NAMES.SUMMARY, getSummaryHeaders());

  // Create and hide the aggregation sheet (PRD 5.2)
  // The createTab function ensures that if this specific sheet already exists, it is reused, not duplicated.
  const aggSheet = createTab(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY, getAggregateHistoryHeaders());
  aggSheet.hideSheet();

  // Insert sample data (PRD 5.6)
  insertSampleData();

  // Apply formulas, formatting and validation (force update during initialization)
  forceUpdateSheets(true);

  // Sort tabs into the correct initial order
  sortSheets();

  // Activate the main tab
  ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY).activate();

  // Clean up default "Sheet1" if it exists
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // Handle potential error if Sheet1 cannot be deleted (e.g. if it's the only sheet)
    }
  }
}

/**
 * Helper to create a tab if it doesn't exist and set its headers.
 * This function prevents the creation of duplicate tabs.
 * @param {string} name - The name of the tab.
 * @param {Array<string>} headers - The header row values.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The created or existing sheet.
 */
function createTab(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  // Only insert the sheet if it does not already exist.
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Ensure headers are set (important if the sheet existed but was corrupted).
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

// Define Headers
function getDebaterHeaders() { return ['Name', 'Debate Type', 'Partner', 'Hard Mode']; }
function getJudgeHeaders() { return ['Name', 'Children\'s Names', 'Debate Type']; }
function getRoomHeaders() { return ['Room Name', 'Debate Type']; }
function getAvailabilityHeaders() { return ['Participant', 'Attending?']; }
function getSummaryHeaders() {
  // Headers reflect tracking the opposing team (TP) or individual (LD).
  return ['Debater Name', 'Debate Type', 'Total Matches', 'Aff Matches', 'Neg Matches', 'BYEs', 'Ironman Matches',
    '#1 Judge', '#2 Judge', '#3 Judge', '#1 Opponent', '#2 Opponent', '#3 Opponent'];
}
// Aggregation Schema: Used for historical tracking and summary calculations.
// This schema supports panels by allowing multiple entries for the same match with different judges.
function getAggregateHistoryHeaders() {
  // Col G (Opponent Team) stores the name of the opposing team, essential for TP tracking.
  return ['Date', 'Type', 'Round', 'Debater Name', 'Is Ironman', 'Role', 'Opponent Team', 'Judge', 'Room'];
}
// Headers updated to reflect potential for multiple judges (panels).
function getTpMatchHeaders() { return ['Aff Team', 'Neg Team', 'Judge(s)', 'Room']; }
function getLdMatchHeaders() { return ['Round', 'Aff Debater', 'Neg Debater', 'Judge(s)', 'Room']; }


/**
 * Inserts the sample data defined in PRD 5.6.
 */
function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // Debaters
  const debatersData = [
    // LD
    ['Abraham Lincoln', 'LD', '', 'No'], ['Stephen A. Douglas', 'LD', '', 'No'],
    ['Clarence Darrow', 'LD', '', 'Yes'], ['William Jennings Bryan', 'LD', '', 'Yes'],
    ['William F. Buckley Jr.', 'LD', '', 'No'], ['Gore Vidal', 'LD', '', 'No'],
    ['Christopher Hitchens', 'LD', '', 'Yes'], ['Tony Blair', 'LD', '', 'Yes'],
    ['Jordan Peterson', 'LD', '', 'No'], ['Slavoj Žižek', 'LD', '', 'No'],
    ['Lloyd Bentsen', 'LD', '', 'No'], ['Dan Quayle', 'LD', '', 'No'],
    ['Richard Dawkins', 'LD', '', 'Yes'], ['Rowan Williams', 'LD', '', 'Yes'],
    ['Diogenes', 'LD', '', 'No'],
    // TP (Including a name with an apostrophe for testing formula safety)
    ['Noam Chomsky', 'TP', 'Michel Foucault', 'Yes'], ['Michel Foucault', 'TP', 'Noam Chomsky', 'Yes'],
    ['Harlow Shapley', 'TP', 'Heber Curtis', 'No'], ['Heber Curtis', 'TP', 'Harlow Shapley', 'No'],
    ['Muhammad Ali', 'TP', "Sandra Day O'Connor", 'Yes'], ["Sandra Day O'Connor", 'TP', 'Muhammad Ali', 'Yes'],
    ['Richard Nixon', 'TP', 'Nikita Khrushchev', 'No'], ['Nikita Khrushchev', 'TP', 'Richard Nixon', 'No'],
    ['Thomas Henry Huxley', 'TP', 'Samuel Wilberforce', 'Yes'], ['Samuel Wilberforce', 'TP', 'Thomas Henry Huxley', 'Yes'],
    ['John F. Kennedy', 'TP', 'David Frost', 'No'], ['David Frost', 'TP', 'John F. Kennedy', 'No'],
    ['Bob Dole', 'TP', 'Bill Clinton', 'No'], ['Bill Clinton', 'TP', 'Bob Dole', 'No'],
  ];
  const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
  if (debatersSheet && debatersSheet.getLastRow() < 2) {
    debatersSheet.getRange(2, 1, debatersData.length, debatersData[0].length).setValues(debatersData);
  }

  // Judges
  const judgesData = [
    ['Howard K. Smith', 'John F. Kennedy', 'TP'],
    ['Fons Elders', 'Noam Chomsky, Michel Foucault', 'TP'],
    ['John Stevens Henslow', 'Samuel Wilberforce', 'TP'],
    ['Jim Lehrer', 'Bill Clinton', 'TP'],
    ['Judy Woodruff', '', 'TP'], ['Tom Brokaw', '', 'TP'], ['Frank McGee', '', 'TP'], ['Quincy Howe', '', 'TP'],
    ['John T. Raulston', 'Clarence Darrow', 'LD'],
    ['Rudyard Griffiths', 'Jordan Peterson', 'LD'],
    ['Stephen J. Blackwood', '', 'LD'], ['Brit Hume', '', 'LD'], ['Jon Margolis', '', 'LD'],
    ['Bill Shadel', '', 'LD'], ['Judge Judy', '', 'LD'],
  ];
  const judgesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JUDGES);
  if (judgesSheet && judgesSheet.getLastRow() < 2) {
    judgesSheet.getRange(2, 1, judgesData.length, judgesData[0].length).setValues(judgesData);
  }

  // Rooms
  const roomsData = [
    ['Room 101', 'LD'], ['Room 102', 'LD'], ['Room 103', 'LD'], ['Room 201', 'LD'],
    ['Sanctuary right', 'LD'], ['Sanctuary left', 'LD'], ['Pantry', 'LD'],
    ['Chapel', 'TP'], ['Library', 'TP'], ['Music lounge', 'TP'],
    ['Cry room', 'TP'], ['Office', 'TP'], ['Office hallway', 'TP'],
  ];
  const roomsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ROOMS);
  if (roomsSheet && roomsSheet.getLastRow() < 2) {
    roomsSheet.getRange(2, 1, roomsData.length, roomsData[0].length).setValues(roomsData);
  }
}

/**
 * Applies formulas to the dynamic tabs (Availability and Match Summary).
 * @param {boolean} [forceUpdate=false] - If true, clears existing formula ranges before reapplying to prevent #REF errors.
 */
function applyFormulas(forceUpdate = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // 1. Availability Tab (PRD 5.2)
  const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (availabilitySheet) {

    if (forceUpdate && availabilitySheet.getMaxRows() > 1) {
      // Clear A2:A to ensure the main array formula can expand cleanly
      availabilitySheet.getRange(2, 1, availabilitySheet.getMaxRows() - 1, 1).clearContent();
    }

    // Formula to combine and sort unique names from Debaters and Judges (PRD 4.1)
    const participantFormula = `=SORT(UNIQUE(FILTER({${CONFIG.SHEET_NAMES.DEBATERS}!A2:A; ${CONFIG.SHEET_NAMES.JUDGES}!A2:A}, {${CONFIG.SHEET_NAMES.DEBATERS}!A2:A; ${CONFIG.SHEET_NAMES.JUDGES}!A2:A}<>"")))`;
    availabilitySheet.getRange('A2').setFormula(participantFormula);

    // Initialize default RSVP status (only during initial setup or if forced/empty)
    if (forceUpdate || availabilitySheet.getRange('B2').getValue() === "") {
      SpreadsheetApp.flush(); // Ensure the participant formula populates
      setRsvpDefaults(availabilitySheet);
    }
  }


  // 2. Match Summary Tab (PRD 5.2) - Must be formula-driven
  const summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  if (!summarySheet) return;

  const HISTORY = CONFIG.SHEET_NAMES.AGGREGATE_HISTORY;
  const maxRows = 500; // Apply formulas to a fixed large range

  // If force updating, clear the entire data range.
  // This prevents #REF! errors caused by array expansion issues when formulas change or conflict.
  if (forceUpdate && summarySheet.getMaxRows() > 1) {
    // Use getMaxColumns() to ensure we clear potential DEBUG columns (like column N) as well.
    const lastCol = Math.max(getSummaryHeaders().length, summarySheet.getMaxColumns());
    summarySheet.getRange(2, 1, summarySheet.getMaxRows() - 1, lastCol).clearContent();
    SpreadsheetApp.flush(); // Ensure clearing is complete before applying new formulas.
  }

  // A2: Debater Name and B2: Debate Type (populated automatically from Debaters tab)
  const summaryFormulaA2 = `=SORT(FILTER({${CONFIG.SHEET_NAMES.DEBATERS}!A2:B}, ${CONFIG.SHEET_NAMES.DEBATERS}!A2:A<>""), 1, TRUE)`;
  summarySheet.getRange('A2').setFormula(summaryFormulaA2);

  SpreadsheetApp.flush();


  // Formulas updated to handle multiple history entries per match (due to judge panels).
  // We count unique combinations of Date(A)/Type(B)/Round(C) to identify unique matches.
  // We use IFERROR(ROWS(UNIQUE(FILTER(...))), 0) for robust counting.

  // C2: Total Matches
  const totalMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!F:F<>"${CONFIG.ROLES.BYE}"))), 0), "")`;
  summarySheet.getRange('C2:C' + maxRows).setFormula(totalMatchesFormula);

  // D2: Aff Matches
  const affMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!F:F="${CONFIG.ROLES.AFF}"))), 0), "")`;
  summarySheet.getRange('D2:D' + maxRows).setFormula(affMatchesFormula);

  // E2: Neg Matches
  const negMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!F:F="${CONFIG.ROLES.NEG}"))), 0), "")`;
  summarySheet.getRange('E2:E' + maxRows).setFormula(negMatchesFormula);

  // F2: BYEs
  const byeMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!F:F="${CONFIG.ROLES.BYE}"))), 0), "")`;
  summarySheet.getRange('F2:F' + maxRows).setFormula(byeMatchesFormula);

  // G2: Ironman Matches
  const ironmanMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!E:E=TRUE))), 0), "")`;
  summarySheet.getRange('G2:G' + maxRows).setFormula(ironmanMatchesFormula);

  // --- Top 3 Formulas (Judges and Opponents) ---
  // We must escape single quotes in names (e.g., O'Connor) before injecting them into the QUERY string.
  // SQL standard requires replacing a single quote (') with two single quotes ('').
  // We use SUBSTITUTE($A2, "'", "''") within the formula construction for robustness.
  const safeNameA2 = `SUBSTITUTE($A2, "'", "''")`;

  // H2: Top 3 Judges (H, I, J)
  // Uses QUERY to find frequency.
  // We use LABEL in the inner query to ensure headers are not returned as data when results are sparse.
  const topJudgesQuery = `"SELECT H, COUNT(H) WHERE D = '"&${safeNameA2}&"' AND H IS NOT NULL AND H <> '' GROUP BY H LABEL H '', COUNT(H) ''"`;
  // The outer QUERY selects only Col1 (the name) while ordering by Col2 (the count).
  const topJudgesFormula = `=IFERROR(IF($A2<>"", TRANSPOSE(QUERY(QUERY('${HISTORY}'!D:H, ${topJudgesQuery}, 0), "Select Col1 ORDER BY Col2 DESC LIMIT 3", 0)), ""), "")`;
  summarySheet.getRange('H2:H' + maxRows).setFormula(topJudgesFormula);

  // K2: Top 3 Opponents (K, L, M)
  // This identifies the top 3 opposing teams (TP) or individuals (LD).
  // We must use QUERY on UNIQUE(A:G) to ensure opponents are counted once per match (handling panels).

  // FIX: The previous implementation used a single QUERY on the result of UNIQUE(). 
  // This is unreliable in Google Sheets when aggregating (GROUP BY/COUNT) on derived arrays with mixed data types.
  // We switch to a robust double-QUERY structure.

  // Inner QUERY data source is the UNIQUE match history (A:G)
  const uniqueHistorySource = `UNIQUE('${HISTORY}'!A:G)`;

  // Construct the inner QUERY string. We must use ColX notation because the source is an array.
  // Col4=Debater Name (D), Col7=Opponent Team (G).
  // This inner query calculates the frequency (COUNT) of each opponent.
  const topOpponentsInnerQuery = `"SELECT Col7, COUNT(Col7) WHERE Col4 = '"&${safeNameA2}&"' AND Col7 IS NOT NULL AND Col7 <> '' AND Col7 <> '${CONFIG.ROLES.BYE}' GROUP BY Col7 LABEL Col7 '', COUNT(Col7) ''"`;

  // Construct the full formula with the outer QUERY for sorting and limiting.
  // We explicitly set the headers parameter to 0 for both queries for clarity and robustness.
  const topOpponentsFormula = `=IFERROR(IF($A2<>"", TRANSPOSE(QUERY(QUERY(${uniqueHistorySource}, ${topOpponentsInnerQuery}, 0), "SELECT Col1 ORDER BY Col2 DESC LIMIT 3", 0)), ""), "")`;

  summarySheet.getRange('K2:K' + maxRows).setFormula(topOpponentsFormula);
}

/**
 * Applies standard formatting to all sheets (PRD 6.1).
 */
function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Check if ss is null (can happen in some execution contexts)
  if (!ss) return;
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Skip hidden sheet formatting
    if (sheetName === CONFIG.SHEET_NAMES.AGGREGATE_HISTORY) return;

    // Handle potentially empty sheets
    if (sheet.getMaxColumns() === 0 || sheet.getMaxRows() === 0) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    // Ensure Match Summary headers are correct if the sheet somehow got corrupted
    if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      createTab(CONFIG.SHEET_NAMES.SUMMARY, getSummaryHeaders());
    }

    const headerRange = sheet.getRange(1, 1, 1, lastCol);

    // Header Styling
    headerRange.setFontWeight('bold')
      .setBackground(CONFIG.STYLES.HEADER_BG)
      .setFontColor(CONFIG.STYLES.HEADER_FONT_COLOR)
      .setHorizontalAlignment('center')
      // Ensure headers do not wrap (e.g., Match Summary), forcing columns wider during auto-resize.
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);


    // Freeze Header (PRD 5.1)
    sheet.setFrozenRows(1);

    // Apply overall font and borders
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const fullRange = sheet.getDataRange();
      fullRange.setFontFamily(CONFIG.STYLES.FONT)
        .setBorder(true, true, true, true, true, true);
    }


    // Apply Banding (Alternating Row Colors) (PRD 6.1)
    if (sheet.getMaxRows() > 1) {
      const dataRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
      // Remove existing banding first
      sheet.getBandings().forEach(banding => banding.remove());
      dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    }

    // Specific alignment adjustments (UX improvement)
    if (sheetName === CONFIG.SHEET_NAMES.DEBATERS || sheetName === CONFIG.SHEET_NAMES.JUDGES || sheetName === CONFIG.SHEET_NAMES.ROOMS || sheetName === CONFIG.SHEET_NAMES.AVAILABILITY) {
      if (sheet.getMaxRows() > 1) {
        sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getLastColumn() - 1).setHorizontalAlignment('center');
      }
    } else if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      if (sheet.getMaxColumns() >= 7) { // Ensure columns B-G exist
        sheet.getRange('B:G').setHorizontalAlignment('center');
      }
    }

    // Auto-sizing Columns (PRD 6.1). This should now resize based on the clipped header text.
    // We limit the auto-resize to the standard columns to avoid resizing potential DEBUG columns.
    let resizeCols = lastCol;
    if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      resizeCols = Math.min(lastCol, getSummaryHeaders().length);
    }

    if (resizeCols > 0) {
      sheet.autoResizeColumns(1, resizeCols);
    }
  });
}

// ... (The rest of the functions remain the same as the provided code, 
// except for the necessary closing braces for functions partially shown above, 
// and the inclusion of getMatchHistory which now includes a defensive check)...

/**
 * Applies dropdown data validations.
 */
function applyDataValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // Availability Tab - Attending? Dropdown (PRD 5.2)
  const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (availabilitySheet) {
    const rsvpRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.RSVP_OPTIONS, true).setAllowInvalid(false).build();
    availabilitySheet.getRange('B2:B').setDataValidation(rsvpRule);
  }

  // Debaters Tab - Type and Hard Mode
  const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
  if (debatersSheet) {
    const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(CONFIG.DEBATE_TYPES), true).build();
    debatersSheet.getRange('B2:B').setDataValidation(typeRule);
    const hardModeRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build();
    debatersSheet.getRange('D2:D').setDataValidation(hardModeRule);
  }

  // Judges and Rooms Tab - Type
  const judgesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JUDGES);
  if (judgesSheet) {
    const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(CONFIG.DEBATE_TYPES), true).build();
    judgesSheet.getRange('C2:C').setDataValidation(typeRule);
  }
  const roomsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ROOMS);
  if (roomsSheet) {
    const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(CONFIG.DEBATE_TYPES), true).build();
    roomsSheet.getRange('B2:B').setDataValidation(typeRule);
  }
}

/**
 * Applies conditional formatting rules for validation highlights (PRD 5.5).
 * Note: This function defines ALL rules for a sheet as setConditionalFormatRules overwrites existing rules.
 */
function applyAllConditionalFormatting() {
  // We use try-catch because this runs onOpen, and sheets might not exist yet.
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    // 1. Availability Tab
    const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
    if (availabilitySheet) {
      const availabilityRules = [];
      const availabilityRange = availabilitySheet.getRange('B2:B');
      // Highlight blank RSVP if Participant exists (PRD 5.2)
      const rsvpMissingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISBLANK(B2), NOT(ISBLANK(A2)))')
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([availabilityRange])
        .build();
      availabilityRules.push(rsvpMissingRule);
      availabilitySheet.setConditionalFormatRules(availabilityRules);
    }

    // 2. Rooms Tab
    const roomsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ROOMS);
    if (roomsSheet && roomsSheet.getMaxRows() > 1) {
      const roomsRules = [];
      const roomsRange = roomsSheet.getRange('A2:B');
      // Room names must be unique (PRD 5.5)
      const uniqueRoomRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND($A2<>"", COUNTIF($A$2:$A, $A2)>1)')
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([roomsRange])
        .build();
      roomsRules.push(uniqueRoomRule);
      roomsSheet.setConditionalFormatRules(roomsRules);
    }

    // 3. Judges Tab
    const judgesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.JUDGES);
    if (judgesSheet && judgesSheet.getMaxRows() > 1) {
      const judgesRules = [];
      const judgesRange = judgesSheet.getRange('A2:C');
      // A person cannot be both a Judge and a Debater (PRD 5.5)
      // Must use INDIRECT() as required by PRD 5.5 implementation note.
      const overlapRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2<>"", COUNTIF(INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A:A"), $A2)>0)`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([judgesRange])
        .build();
      judgesRules.push(overlapRule);

      // Note: Validating comma-delimited "Children's Names" existence is too complex for CF formulas.
      // This is handled by the Apps Script pre-flight validation (validateRosters).

      judgesSheet.setConditionalFormatRules(judgesRules);
    }

    // 4. Debaters Tab
    const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
    if (debatersSheet && debatersSheet.getMaxRows() > 1) {
      const debatersRules = [];
      const debatersRange = debatersSheet.getRange('A2:D');

      // A person cannot be both a Judge and a Debater
      const overlapRuleDebater = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2<>"", COUNTIF(INDIRECT("${CONFIG.SHEET_NAMES.JUDGES}!A:A"), $A2)>0)`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(overlapRuleDebater);

      // Partnership Consistency (PRD 5.5)

      // Partner must exist in the Debaters roster
      const partnerExistsRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", ISNA(MATCH($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A:A"), 0)))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(partnerExistsRule);

      // Partnerships must be reciprocal.
      // Note: We use INDIRECT here as well, although VLOOKUP on the same sheet might work, consistency is better.
      const reciprocalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", IFERROR(VLOOKUP($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A2:C"), 3, FALSE) <> $A2, FALSE))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(reciprocalRule);

      // Partners must have the same "Hard Mode" setting.
      const hardModeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", IFERROR(VLOOKUP($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A2:D"), 4, FALSE) <> $D2, FALSE))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(hardModeRule);

      debatersSheet.setConditionalFormatRules(debatersRules);
    }
  } catch (e) {
    Logger.log("Error applying conditional formatting: " + e.message);
  }
}


// =============================================================================
// Weekly Workflow Functions
// =============================================================================

/**
 * Clears the RSVPs in the Availability tab, setting them back to "Not responded" (PRD 4.2, 5.3).
 */
function clearRsvps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (availabilitySheet) {
    setRsvpDefaults(availabilitySheet);
    ss.toast("RSVPs cleared and reset to 'Not responded'.", "Workflow Complete", 3);
  }
}

/**
 * Helper to set the default RSVP status based on the participant list (PRD 5.2).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Availability sheet.
 */
function setRsvpDefaults(sheet) {
  const lastRow = sheet.getLastRow();
  // Clear the existing RSVP column content first to ensure clean state
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).clearContent();
  }

  if (lastRow < 2) return;

  // Read column A (Participants)
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Map to Column B defaults
  const newRsvps = data.map(row => {
    // If Participant name exists, set to "Not responded", otherwise blank
    return [row[0] ? "Not responded" : ""];
  });

  // Write Column B
  if (newRsvps.length > 0) {
    sheet.getRange(2, 2, newRsvps.length, 1).setValues(newRsvps);
  }
}

/**
 * Main function to generate Team Policy (TP) matches.
 */
function generateTpMatches() {
  generateMatches(CONFIG.DEBATE_TYPES.TP);
}

/**
 * Main function to generate Lincoln-Douglas (LD) matches (2 rounds).
 */
function generateLdMatches() {
  generateMatches(CONFIG.DEBATE_TYPES.LD);
}

/**
 * Core workflow function for generating matches for a given type (PRD 5.3).
 * @param {string} debateType - 'TP' or 'LD'.
 */
function generateMatches(debateType) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Run data integrity checks (PRD 5.3 Step 1)
    if (!validateRosters()) {
      return; // Validation function handles the alert
    }

    // 2. Get available participants and history (PRD 5.3 Step 2)
    const history = getMatchHistory();
    const resources = getAvailableResources(debateType, history);

    if (resources.debaters.length === 0) {
      ss.toast(`No available debaters for ${debateType}. Cannot generate matches.`, 'Notice', 5);
      return;
    }

    // 3. Run the pairing logic (PRD 5.3 Step 3)
    let allPairings = [];

    if (debateType === CONFIG.DEBATE_TYPES.TP) {
      // formTpTeams ensures consistent alphabetical naming for permanent teams.
      const teams = formTpTeams(resources.debaters);
      const pairings = executePairingRound(teams, resources.judges, resources.rooms, history, 1);
      allPairings = pairings;

    } else if (debateType === CONFIG.DEBATE_TYPES.LD) {
      // LD requires sequential 2-round pairing (PRD 5.4.1)

      // Convert LD debaters to simple "teams" of one person
      const teams = resources.debaters.map(d => ({
        name: d.name,
        members: [d.name],
        hardMode: d.hardMode,
        isIronman: false,
        history: d.history
      }));

      // Round 1 (PRD 5.4.1 Step 1)
      const round1Pairings = executePairingRound(teams, resources.judges, resources.rooms, history, 1);
      allPairings.push(...round1Pairings);

      // Identify R1 BYE recipient to prevent repeat BYE in R2
      const r1ByePairing = round1Pairings.find(p => p.isBye);
      const r1ByeTeamName = r1ByePairing ? r1ByePairing.Aff.name : null;

      // Update History In-Memory (PRD 5.4.1 Step 2)
      const historyR2 = updateHistoryInMemory(history, round1Pairings);

      // Round 2 (PRD 5.4.1 Step 3)
      // Resources can be reused, but we pass the updated history and the BYE exclusion.
      const round2Pairings = executePairingRound(teams, resources.judges, resources.rooms, historyR2, 2, r1ByeTeamName);
      allPairings.push(...round2Pairings);
    }

    // 4. Check for existing sheet (PRD 5.3 Step 4)
    const dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const sheetName = `${debateType} ${dateString}`;
    if (ss.getSheetByName(sheetName)) {
      ui.alert('Overwrite Protection', `A sheet named "${sheetName}" already exists. Delete it manually if you wish to regenerate pairings for today.`, ui.ButtonSet.OK);
      return;
    }

    // 5. Write pairings and update history (PRD 5.3 Step 5)
    const newSheet = createMatchSheet(sheetName, debateType, allPairings);
    updateAggregateHistory(allPairings, debateType, dateString);

    // 6. Re-sort all tabs and apply formatting (PRD 5.3 Step 6)
    sortSheets();

    // Re-apply formatting and formulas. We force update formulas to ensure Match Summary reflects the new data correctly (PRD 5.2).
    // This ensures that the array formulas (like Top Opponents) can expand correctly after data updates.
    forceUpdateSheets(true);

    applyMatchSheetValidation(newSheet, debateType);

    // Ensure formulas recalculate after history update and formatting
    SpreadsheetApp.flush();

    newSheet.activate();
    // Success feedback is the visual change, no alert needed (PRD 5.7)
    ss.toast(`${debateType} Matches Generated Successfully.`, 'Success', 5);

  } catch (error) {
    Logger.log(error.stack);
    // Display critical errors using alerts (PRD 5.7)
    if (ui) {
      ui.alert('Error Generating Matches', error.message, ui.ButtonSet.OK);
    } else {
      Logger.log('Could not display UI alert: ' + error.message);
    }
  }
}

// =============================================================================
// Data Retrieval & Modeling
// =============================================================================

/**
 * Reads the roster data (Debaters, Judges, Rooms) from the sheets.
 * @returns {object} An object containing arrays of debaters, judges, and rooms.
 */
function getRosterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const readSheetData = (sheetName, headersLength) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, headersLength).getValues();
  };

  const debatersData = readSheetData(CONFIG.SHEET_NAMES.DEBATERS, 4);
  const judgesData = readSheetData(CONFIG.SHEET_NAMES.JUDGES, 3);
  const roomsData = readSheetData(CONFIG.SHEET_NAMES.ROOMS, 2);

  // Ensure names are trimmed for data consistency.
  const debaters = debatersData.map(row => ({
    name: String(row[0]).trim(),
    type: String(row[1]).trim(),
    partner: String(row[2]).trim(),
    hardMode: String(row[3]).trim() === 'Yes'
  })).filter(d => d.name);

  const judges = judgesData.map(row => ({
    name: String(row[0]).trim(),
    // Parse comma-delimited children names
    children: String(row[1]).split(',').map(name => name.trim()).filter(name => name.length > 0),
    type: String(row[2]).trim()
  })).filter(j => j.name);

  const rooms = roomsData.map(row => ({
    name: String(row[0]).trim(),
    type: String(row[1]).trim()
  })).filter(r => r.name);

  return { debaters, judges, rooms };
}

/**
 * Reads the Availability sheet to determine who is attending.
 * @returns {Set<string>} A set of names of participants attending ('Yes').
 */
function getAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  const attending = new Set();

  if (!sheet || sheet.getLastRow() < 2) return attending;

  // Read Participant (A) and Attending? (B) columns
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(row => {
    if (row[1] === 'Yes') {
      attending.add(String(row[0]).trim());
    }
  });

  return attending;
}

/**
 * Reads the AGGREGATE_HISTORY sheet and structures the data for optimization lookups.
 * Note: This function handles the history structure where multiple entries exist for one match if panels are used.
 * @returns {object} A history object containing detailed stats per debater and judge history.
 */
function getMatchHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY);
  const history = {
    debaters: {},
    judges: {}
  };

  if (!sheet || sheet.getLastRow() < 2) return history;

  // Headers: Date(0), Type(1), Round(2), Debater Name(3), Is Ironman(4), Role(5), Opponent Team(6), Judge(7), Room(8)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  // We use a set to track unique matches processed for each debater to avoid inflating counts due to panels.
  const processedMatches = new Set();

  data.forEach(row => {
    const date = String(row[0]);
    const type = String(row[1]);
    const round = String(row[2]);
    // Ensure names read from history are trimmed for consistency.
    const debaterName = String(row[3]).trim();

    // Defensive check: Skip processing if the row is missing the debater name.
    if (!debaterName) return;

    const role = String(row[5]);
    // Opponent Team name is stored in Col G (6). This correctly handles TP teams.
    const opponentTeam = String(row[6]).trim();
    const judge = String(row[7]).trim();
    const room = String(row[8]).trim();

    // Define a unique key for the match from the perspective of the debater.
    const matchKey = `${debaterName}|${date}|${type}|${round}`;

    if (!history.debaters[debaterName]) {
      history.debaters[debaterName] = { byes: 0, opponents: {}, judges: {}, affCount: 0, negCount: 0 };
    }

    const stats = history.debaters[debaterName];

    // Process match stats (Aff/Neg/BYE/Opponent) only once per match.
    if (!processedMatches.has(matchKey)) {
      if (role === CONFIG.ROLES.BYE) {
        stats.byes++;
      } else {
        if (role === CONFIG.ROLES.AFF) stats.affCount++;
        if (role === CONFIG.ROLES.NEG) stats.negCount++;

        // Track the opposing team name.
        if (opponentTeam && opponentTeam !== CONFIG.ROLES.BYE) {
          stats.opponents[opponentTeam] = (stats.opponents[opponentTeam] || 0) + 1;
        }
      }
      processedMatches.add(matchKey);
    }

    // Process judge/room stats (always process, as each row represents a judge interaction).
    if (role !== CONFIG.ROLES.BYE) {
      if (judge) {
        stats.judges[judge] = (stats.judges[judge] || 0) + 1;

        // Track which rooms judges have used (Soft Constraint 4)
        if (!history.judges[judge]) history.judges[judge] = { rooms: {} };
        if (room) {
          history.judges[judge].rooms[room] = (history.judges[judge].rooms[room] || 0) + 1;
        }
      }
    }
  });

  return history;
}

/**
 * Combines roster data, attendance, and history to provide available resources for a specific debate type.
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {object} history - The match history object.
 * @returns {object} Available debaters, judges, and rooms with enriched data.
 */
function getAvailableResources(debateType, history) {
  const { debaters, judges, rooms } = getRosterData();
  const attending = getAttendance();

  const availableDebaters = debaters
    .filter(item => item.type === debateType && attending.has(item.name))
    .map(item => {
      // Attach history
      const itemHistory = (history.debaters[item.name] || { byes: 0, opponents: {}, judges: {}, affCount: 0, negCount: 0 });
      return { ...item, history: itemHistory };
    });

  const availableJudges = judges
    .filter(item => item.type === debateType && attending.has(item.name))
    .map(judge => {
      // Attach room history (PRD 5.4 Soft Constraint 4)
      judge.roomHistory = (history.judges && history.judges[judge.name]) ? history.judges[judge.name].rooms : {};
      return judge;
    });

  // Rooms must match the debate type (Hard Constraint 4)
  const availableRooms = rooms.filter(item => item.type === debateType);

  return {
    debaters: availableDebaters,
    judges: availableJudges,
    rooms: availableRooms
  };
}

/**
 * Forms TP teams from available debaters, handling the Ironman case (PRD 5.4).
 * Teams are permanent as defined in the Debaters sheet.
 * CRITICAL: Team names must be deterministic for consistent history tracking.
 * @param {Array<object>} debaters - List of available TP debaters (enriched with history).
 * @returns {Array<object>} List of formed teams.
 */
function formTpTeams(debaters) {
  const teams = [];
  const processed = new Set();

  // The team name format ("Name1 / Name2") is critical for opponent tracking in Match Summary.
  debaters.forEach(debater => {
    if (processed.has(debater.name)) return;

    // Find the partner in the available list.
    const partner = debaters.find(d => d.name === debater.partner);

    if (partner) {
      // Full team (both partners present). 
      // Ensure consistent naming convention (alphabetical) so opponent tracking is reliable.
      const members = [debater.name, partner.name].sort();

      teams.push({
        name: `${members[0]} / ${members[1]}`,
        members: members,
        hardMode: debater.hardMode,
        isIronman: false,
        // Use the history of the member with fewer BYEs for fair BYE prioritization
        // We use the history from the original debater/partner objects captured in the loop closure.
        history: debater.history.byes <= partner.history.byes ? debater.history : partner.history
      });
      processed.add(debater.name);
      processed.add(partner.name);
    } else {
      // Ironman team (partner is missing or not attending) (PRD 5.4 Special Cases)
      teams.push({
        name: `${debater.name} ${CONFIG.ROLES.IRONMAN_SUFFIX}`,
        members: [debater.name],
        hardMode: debater.hardMode,
        isIronman: true,
        history: debater.history
      });
      processed.add(debater.name);
    }
  });

  return teams;
}


// =============================================================================
// Validation Functions (Pre-flight)
// =============================================================================

/**
 * Validates the integrity of the rosters before generating matches (PRD 5.5).
 * @returns {boolean} True if valid, false otherwise.
 */
function validateRosters() {
  const { debaters, judges, rooms } = getRosterData();
  const ui = SpreadsheetApp.getUi();
  let errors = [];

  const debaterNames = new Set(debaters.map(d => d.name));
  const judgeNames = new Set(judges.map(j => j.name));

  // 1. Check for Judge/Debater Overlap
  debaterNames.forEach(name => {
    if (judgeNames.has(name)) {
      errors.push(`"${name}" is listed as both a Debater and a Judge.`);
    }
  });

  // 2. Validate Judge's Children
  judges.forEach(judge => {
    judge.children.forEach(childName => {
      if (!debaterNames.has(childName)) {
        errors.push(`Judge "${judge.name}" lists child "${childName}", who is not in the Debaters roster.`);
      }
    });
  });

  // 3. Validate TP Partnerships (Permanent Teams)
  const tpDebaters = debaters.filter(d => d.type === CONFIG.DEBATE_TYPES.TP);
  const debaterMap = new Map(debaters.map(d => [d.name, d]));

  tpDebaters.forEach(debater => {
    if (!debater.partner) {
      return; // Allowed if no partner is listed (intends to be Ironman)
    }

    const partner = debaterMap.get(debater.partner);

    // Partner existence check
    if (!partner) {
      errors.push(`Debater "${debater.name}" lists partner "${debater.partner}", who is not in the Debaters roster.`);
      return;
    }

    // Partner type check
    if (partner.type !== CONFIG.DEBATE_TYPES.TP) {
      errors.push(`Debater "${debater.name}" (TP) lists partner "${debater.partner}" (${partner.type}). Partners must both be TP.`);
    }

    // Reciprocal partnership
    if (partner.partner !== debater.name) {
      errors.push(`Partnership mismatch: "${debater.name}" lists "${debater.partner}", but "${partner.name}" lists "${partner.partner || 'nobody'}".`);
    }

    // Hard Mode consistency
    if (debater.hardMode !== partner.hardMode) {
      errors.push(`Hard Mode mismatch: "${debater.name}" and "${debater.partner}" must have the same Hard Mode setting.`);
    }
  });

  // 4. Room Uniqueness
  const roomNameSet = new Set();
  rooms.forEach(room => {
    if (roomNameSet.has(room.name)) {
      errors.push(`Duplicate Room Name found: "${room.name}".`);
    }
    roomNameSet.add(room.name);
  });

  // Display errors if any
  if (errors.length > 0) {
    const errorMessage = "Roster validation failed. Please correct the following issues (also check conditional formatting highlights):\n\n" + errors.join('\n');
    if (ui) {
      ui.alert('Validation Error', errorMessage, ui.ButtonSet.OK);
    }
    return false;
  }

  return true;
}

// =============================================================================
// Core Pairing Logic
// =============================================================================

/**
 * Executes the pairing logic for a single round of debates.
 * Uses a randomized iterative approach (Monte Carlo method) to optimize based on soft constraints.
 * @param {Array<object>} teams - The teams/debaters participating.
 * @param {Array<object>} judges - Available judges.
 * @param {Array<object>} rooms - Available rooms.
 * @param {object} history - The current match history (potentially updated in-memory).
 * @param {number} roundNum - The round number (1 or 2).
 * @param {string|null} [excludedFromBye=null] - Name of the team excluded from getting a BYE (for LD Round 2).
 * @returns {Array<object>} The generated pairings.
 */
function executePairingRound(teams, judges, rooms, history, roundNum, excludedFromBye = null) {
  // 1. Handle BYE assignment (PRD 5.4 Special Cases)
  let availableTeams = [...teams];
  let pairings = [];
  let byeTeam = null;

  if (availableTeams.length % 2 !== 0) {
    // Sort by fewest historical BYEs to ensure fairness
    availableTeams.sort((a, b) => a.history.byes - b.history.byes);

    // Find the best candidate who is not excluded (for LD R2 fairness)
    let byeIndex = availableTeams.findIndex(team => team.name !== excludedFromBye);

    if (byeIndex === -1) {
      // If the only participant(s) left were excluded (e.g., only 1 participant total),
      // they must take the bye regardless of exclusion.
      byeIndex = 0;
    }

    if (availableTeams.length > 0) {
      byeTeam = availableTeams.splice(byeIndex, 1)[0];

      pairings.push({
        Round: roundNum,
        Aff: byeTeam,
        Neg: { name: CONFIG.ROLES.BYE, members: [] },
        Judges: [], // Initialize Judges array for panels
        Room: null,
        Penalty: 0,
        isBye: true
      });
    }
  }

  // 2. Generate optimized pairings (PRD 5.4 Soft Constraints)

  const NUM_ITERATIONS = 500; // Iterations for the optimization algorithm
  let bestPairingSet = null;
  let lowestPenaltyScore = Infinity;

  if (availableTeams.length === 0) {
    return pairings; // No matches to pair if only 0 or 1 team.
  }

  for (let i = 0; i < NUM_ITERATIONS; i++) {
    // Shuffle remaining teams to create varied initial pairings
    const shuffledTeams = shuffleArray([...availableTeams]);
    const currentPairings = [];
    let currentTotalPenalty = 0;

    for (let j = 0; j < shuffledTeams.length; j += 2) {
      const team1 = shuffledTeams[j];
      const team2 = shuffledTeams[j + 1];

      // Decide Aff/Neg to balance history
      // Calculate Aff Advantage (Higher means they have done Neg more often)
      const team1AffAdvantage = team1.history.negCount - team1.history.affCount;
      const team2AffAdvantage = team2.history.negCount - team2.history.affCount;

      let aff, neg;
      if (team1AffAdvantage > team2AffAdvantage) {
        aff = team1; neg = team2;
      } else if (team2AffAdvantage > team1AffAdvantage) {
        aff = team2; neg = team1;
      } else {
        // Equal history, randomly assign
        [aff, neg] = (Math.random() > 0.5) ? [team1, team2] : [team2, team1];
      }

      const penalty = calculatePairingPenalty(aff, neg, history);
      currentTotalPenalty += penalty;

      currentPairings.push({
        Round: roundNum,
        Aff: aff,
        Neg: neg,
        Judges: [], // Initialize Judges array
        Room: null,
        Penalty: penalty,
        isBye: false
      });
    }

    if (currentTotalPenalty < lowestPenaltyScore) {
      lowestPenaltyScore = currentTotalPenalty;
      bestPairingSet = currentPairings;
    }

    // Optimization: If score is 0 (perfect), stop early
    if (lowestPenaltyScore === 0) break;
  }

  if (bestPairingSet) {
    pairings.push(...bestPairingSet);
  }

  // 3. Assign Judges (including panels if surplus exists) and Rooms
  assignResources(pairings, judges, rooms, history);

  return pairings;
}

/**
 * Calculates the penalty score for a specific pairing based on soft constraints (PRD 5.4).
 * @param {object} team1 - The first team (Aff).
 * @param {object} team2 - The second team (Neg).
 * @param {object} history - Match history.
 * @returns {number} The penalty score (lower is better).
 */
function calculatePairingPenalty(team1, team2, history) {
  let penalty = 0;

  // Constraint 1: Hard Mode mismatch (High Penalty)
  if (team1.hardMode !== team2.hardMode) {
    penalty += 100;
  }

  // Constraint 2: Minimize rematches (Medium Penalty)
  // We check if the individuals have faced the opposing TEAM name before.
  const checkRematches = (team, opponentTeamName) => {
    let rematches = 0;
    // We iterate over members. Although if teams are permanent, checking one member might suffice, 
    // iterating ensures robustness if history somehow differs between partners (e.g. previous Ironman history).
    team.members.forEach(member => {
      // Check the history passed in memory, which might be updated from Round 1
      const globalMemberHistory = (history.debaters && history.debaters[member]) || { opponents: {} };
      // History correctly tracks the opponent team name.
      if (globalMemberHistory.opponents[opponentTeamName]) {
        rematches += globalMemberHistory.opponents[opponentTeamName];
      }
    });
    return rematches;
  };

  // We calculate the total number of previous encounters between the teams.
  // Since team names are deterministic now, this comparison is reliable.
  penalty += (checkRematches(team1, team2.name) + checkRematches(team2, team1.name)) * 15;

  return penalty;
}

/**
 * Helper function to calculate judge penalty, used in assignResources.
 * @param {object} judge - The judge object.
 * @param {Array<string>} participants - List of debater names in the match.
 * @param {object} history - Match history.
 * @returns {number} The penalty score (Infinity if conflict exists).
 */
function calculateJudgePenalty(judge, participants, history) {
  // Hard Constraint 3: Parent-Child conflict
  const conflict = judge.children.some(child => participants.includes(child));
  if (conflict) return Infinity;

  // Soft Constraint 3: Minimize re-judging
  let judgePenalty = 0;
  participants.forEach(member => {
    // Check the history passed in memory, which might be updated from Round 1
    const memberHistory = (history.debaters && history.debaters[member]) || { judges: {} };
    if (memberHistory.judges[judge.name]) {
      // Penalty increases exponentially with repeat judging to strongly encourage variety
      judgePenalty += 10 * Math.pow(memberHistory.judges[judge.name], 2);
    }
  });
  return judgePenalty;
};


/**
 * Assigns judges and rooms to the generated pairings, respecting constraints and utilizing surplus judges.
 * Throws an error if hard constraints cannot be met (PRD 5.4).
 * @param {Array<object>} pairings - The list of pairings (mutated by this function).
 * @param {Array<object>} availableJudges - Available judges.
 * @param {Array<object>} availableRooms - Available rooms.
 * @param {object} history - Match history.
 */
function assignResources(pairings, availableJudges, availableRooms, history) {
  // Create consumable pools, shuffled for randomness when scores are tied
  const judgesPool = shuffleArray([...availableJudges]);
  const roomsPool = shuffleArray([...availableRooms]);
  const assignedJudges = new Set();
  const assignedRooms = new Set();

  // Filter out BYEs and sort by penalty descending (prioritize assigning resources to difficult matchups)
  const matchesToAssign = pairings.filter(p => !p.isBye).sort((a, b) => b.Penalty - a.Penalty);

  // Check Hard Constraint 1: Sufficiency (At least one judge per match)
  if (matchesToAssign.length > judgesPool.length) {
    throw new Error(`Insufficient judges. Required (minimum): ${matchesToAssign.length}, Available: ${judgesPool.length}.`);
  }
  if (matchesToAssign.length > roomsPool.length) {
    throw new Error(`Insufficient rooms. Required: ${matchesToAssign.length}, Available: ${roomsPool.length}.`);
  }

  // --- Pass 1: Assign Primary Judges (Ensure coverage) ---
  for (const pairing of matchesToAssign) {
    const participants = [...pairing.Aff.members, ...pairing.Neg.members];

    let bestJudge = null;
    let lowestJudgePenalty = Infinity;

    for (const judge of judgesPool) {
      // Check if the judge is already assigned in this round
      if (assignedJudges.has(judge.name)) continue;

      const penalty = calculateJudgePenalty(judge, participants, history);

      if (penalty < lowestJudgePenalty) {
        lowestJudgePenalty = penalty;
        bestJudge = judge;
      }
    }

    if (bestJudge && lowestJudgePenalty !== Infinity) {
      pairing.Judges.push(bestJudge);
      assignedJudges.add(bestJudge.name);
      pairing.Penalty += lowestJudgePenalty; // Add judge penalty to total penalty
    } else {
      // CRITICAL FAILURE: A match cannot be judged due to conflicts. (PRD 5.4 Hard Constraints)
      throw new Error(`Could not find a conflict-free primary judge for match: ${pairing.Aff.name} vs ${pairing.Neg.name}. Please review judge availability and conflicts.`);
    }
  }

  // --- Pass 2: Assign Surplus Judges (Paneling) ---
  const surplusJudges = judgesPool.filter(j => !assignedJudges.has(j.name));

  if (surplusJudges.length > 0 && matchesToAssign.length > 0) {
    // Iterate through the surplus judges and assign them one by one to the best possible match.
    for (const judge of surplusJudges) {
      let bestMatch = null;
      let lowestTotalPenalty = Infinity;
      let bestActualJudgePenalty = 0;

      for (const pairing of matchesToAssign) {
        const participants = [...pairing.Aff.members, ...pairing.Neg.members];

        const penalty = calculateJudgePenalty(judge, participants, history);

        // Encourage even distribution by adding a penalty based on the current panel size.
        // This ensures we prioritize adding a 2nd judge before adding a 3rd, etc.
        const distributionPenalty = pairing.Judges.length * 50;
        const totalPenalty = penalty + distributionPenalty;

        if (totalPenalty < lowestTotalPenalty) {
          lowestTotalPenalty = totalPenalty;
          bestMatch = pairing;
          bestActualJudgePenalty = penalty;
        }
      }

      if (bestMatch && lowestTotalPenalty !== Infinity) {
        bestMatch.Judges.push(judge);
        // Update the match penalty with the actual penalty cost of this judge
        bestMatch.Penalty += bestActualJudgePenalty;
      }
      // If a surplus judge cannot be placed anywhere (due to conflicts), they remain unassigned. This is acceptable.
    }
  }


  // --- Pass 3: Assign Rooms ---
  for (const pairing of matchesToAssign) {
    // We assume the first judge in the array (assigned in Pass 1) is the primary/chair judge for room preference (Soft Constraint 4).
    // Check if primaryJudge exists (it should always exist for non-BYE matches after Pass 1).
    const primaryJudge = pairing.Judges.length > 0 ? pairing.Judges[0] : null;

    let bestRoom = null;
    let highestRoomPreferenceScore = -1; // We want to maximize preference here

    for (const room of roomsPool) {
      if (assignedRooms.has(room.name)) continue;

      // Soft Constraint 4: Preferentially assign judges to rooms they've used before
      const preferenceScore = (primaryJudge && primaryJudge.roomHistory[room.name]) || 0;

      if (preferenceScore > highestRoomPreferenceScore) {
        highestRoomPreferenceScore = preferenceScore;
        bestRoom = room;
      }
    }

    if (bestRoom) {
      pairing.Room = bestRoom;
      assignedRooms.add(bestRoom.name);
    }
    // Note: We verified sufficient rooms at the start, so bestRoom should always be found.
  }
}

/**
 * Updates the in-memory history object with the results of a round.
 * Used specifically for LD sequential pairing (PRD 5.4.1).
 * Handles updates correctly when judge panels are used.
 * @param {object} history - The original history object.
 * @param {Array<object>} pairings - The pairings from the round just generated.
 * @returns {object} A new history object updated with the round results.
 */
function updateHistoryInMemory(history, pairings) {
  // Deep copy the history object to avoid mutating the original
  // Note: This simple deep copy works for this specific data structure.
  const newHistory = JSON.parse(JSON.stringify(history));
  if (!newHistory.debaters) newHistory.debaters = {};

  // Helper function to update stats for a team in a match
  const updateMatchStats = (team, opponentTeamName, role, judgesList) => {
    team.members.forEach(memberName => {
      if (!newHistory.debaters[memberName]) {
        // Initialize if missing (e.g., new debater)
        newHistory.debaters[memberName] = { byes: 0, opponents: {}, judges: {}, affCount: 0, negCount: 0 };
      }
      const stats = newHistory.debaters[memberName];

      if (role === CONFIG.ROLES.BYE) {
        stats.byes++;
      } else {
        // Update Match counts (once per match)
        if (role === CONFIG.ROLES.AFF) stats.affCount++;
        if (role === CONFIG.ROLES.NEG) stats.negCount++;

        // Update Opponent Team counts (once per match)
        if (opponentTeamName && opponentTeamName !== CONFIG.ROLES.BYE) {
          stats.opponents[opponentTeamName] = (stats.opponents[opponentTeamName] || 0) + 1;
        }

        // Update Judge counts (for every judge in the panel)
        if (judgesList && judgesList.length > 0) {
          judgesList.forEach(judge => {
            const judgeName = judge.name;
            if (!stats.judges[judgeName]) {
              stats.judges[judgeName] = 0;
            }
            stats.judges[judgeName]++;
          });
        }
      }
    });
  };

  pairings.forEach(pairing => {
    if (pairing.isBye) {
      // Aff got the BYE
      updateMatchStats(pairing.Aff, CONFIG.ROLES.BYE, CONFIG.ROLES.BYE, null);
    } else {
      // Standard match
      updateMatchStats(pairing.Aff, pairing.Neg.name, CONFIG.ROLES.AFF, pairing.Judges);
      updateMatchStats(pairing.Neg, pairing.Aff.name, CONFIG.ROLES.NEG, pairing.Judges);
    }
  });

  // Note: Judge room history is not updated in memory, as optimization relies on historical data.

  return newHistory;
}


// =============================================================================
// Output and History Functions
// =============================================================================

/**
 * Creates the new match sheet and writes the pairings.
 * @param {string} sheetName - The name for the new sheet.
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {Array<object>} pairings - The final pairings.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The newly created sheet.
 */
function createMatchSheet(sheetName, debateType, pairings) {
  const headers = (debateType === CONFIG.DEBATE_TYPES.TP) ? getTpMatchHeaders() : getLdMatchHeaders();
  const sheet = createTab(sheetName, headers);

  const outputData = pairings.map(p => {
    // Handle multiple judges (panels) by joining names with a comma, sorted alphabetically.
    const judgeNames = (p.Judges && p.Judges.length > 0) ? p.Judges.map(j => j.name).sort().join(', ') : '';
    const roomName = p.Room ? p.Room.name : '';

    if (debateType === CONFIG.DEBATE_TYPES.TP) {
      // TP: Aff Team | Neg Team | Judge(s) | Room
      return [p.Aff.name, p.Neg.name, judgeNames, roomName];
    } else {
      // LD: Round | Aff Debater | Neg Debater | Judge(s) | Room
      return [p.Round, p.Aff.name, p.Neg.name, judgeNames, roomName];
    }
  });

  if (outputData.length > 0) {
    sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);

    // Sort the data (PRD 5.5)
    const dataRange = sheet.getRange(2, 1, outputData.length, headers.length);
    if (debateType === CONFIG.DEBATE_TYPES.LD) {
      // Sort by Round (Col 1 asc), then Room (Col 5 asc)
      dataRange.sort([{ column: 1, ascending: true }, { column: 5, ascending: true }]);
    } else {
      // Sort by Room (Col 4 asc)
      dataRange.sort({ column: 4, ascending: true });
    }
  }

  return sheet;
}

/**
 * Applies conditional formatting validation rules to a newly generated match sheet (PRD 5.5).
 * This is critical for supporting manual adjustments (PRD 4.2), especially with panel judging.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The match sheet.
 * @param {string} debateType - 'TP' or 'LD'.
 */
function applyMatchSheetValidation(sheet, debateType) {
  const rules = [];
  // Ensure the sheet has data before applying validation
  if (sheet.getMaxRows() < 2) return;

  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn());

  // Helper function for INDIRECT references (PRD 5.5 Implementation Note)
  const rosterRef = (name) => `INDIRECT("${CONFIG.SHEET_NAMES[name]}!A:A")`;

  // TP Columns: A=Aff, B=Neg, C=Judge(s), D=Room
  // LD Columns: A=Round, B=Aff, C=Neg, D=Judge(s), E=Room

  const judgeCol = debateType === 'TP' ? 'C' : 'D';
  const roomCol = debateType === 'TP' ? 'D' : 'E';

  // --- Referential Integrity (PRD 5.5) ---

  // Judge(s) must exist in roster. Checks if any name in the comma-separated list is missing.
  // This uses complex array formulas (SUMPRODUCT, SPLIT, MATCH) which are supported in CF.
  const judgeRosterFormula = `=AND(LEN(TRIM($${judgeCol}2))>0, IFERROR(SUMPRODUCT(--(ISNA(MATCH(TRIM(SPLIT($${judgeCol}2, ",")), ${rosterRef('JUDGES')}, 0)))), 1)>0)`;

  const judgeExistsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(judgeRosterFormula)
    .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
    .setRanges([dataRange])
    .build();
  rules.push(judgeExistsRule);

  // Room must exist in roster
  const roomExistsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${roomCol}2<>"", ISNA(MATCH($${roomCol}2, ${rosterRef('ROOMS')}, 0)))`)
    .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
    .setRanges([dataRange])
    .build();
  rules.push(roomExistsRule);

  // --- Uniqueness Constraints (PRD 5.5) ---

  if (debateType === 'TP') {
    // TP: Each judge is assigned to only one match.
    // This uses complex array formulas (TEXTJOIN, SPLIT, COUNTIF) to check for individual judge duplication across panels.
    const uniqueJudgeFormula = `=AND(LEN(TRIM($C2))>0, IFERROR(SUMPRODUCT(--(COUNTIF(TRIM(SPLIT(TEXTJOIN(",", TRUE, $C$2:$C), ",")), TRIM(SPLIT($C2, ",")))>1)), 0)>0)`;

    const uniqueJudgeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueJudgeFormula)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueJudgeRule);

    // TP: Rooms must be unique
    const uniqueRoomRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($D2<>"", COUNTIF($D$2:$D, $D2)>1)`)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueRoomRule);

  } else if (debateType === 'LD') {
    // LD: Resources must be unique WITHIN THE ROUND

    // Unique Judge per round (checks individual judges across panels within the round)
    const uniqueJudgeLDFormula = `=AND(LEN(TRIM($D2))>0, IFERROR(SUMPRODUCT(--(COUNTIF(TRIM(SPLIT(TEXTJOIN(",", TRUE, FILTER($D$2:$D, $A$2:$A=$A2)), ",")), TRIM(SPLIT($D2, ",")))>1)), 0)>0)`;

    const uniqueJudgeLDRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueJudgeLDFormula)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueJudgeLDRule);

    // Unique Room per round
    const uniqueRoomLDRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($E2<>"", COUNTIFS($A$2:$A, $A2, $E$2:$E, $E2)>1)`)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueRoomLDRule);


    // LD: Each debater is in one match PER ROUND. (PRD 5.5 Critical Logic Failure check)
    // Check if Debater in B (Aff) appears elsewhere in B or C within the same round (A).
    const uniqueDebaterAffFormula = `=AND($B2<>"${CONFIG.ROLES.BYE}", $B2<>"", OR(COUNTIFS($B$2:$B, $B2, $A$2:$A, $A2)>1, COUNTIFS($C$2:$C, $B2, $A$2:$A, $A2)>0))`;
    const uniqueDebaterAffRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueDebaterAffFormula)
      .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
      .setRanges([sheet.getRange(`B2:B${sheet.getMaxRows()}`)]) // Apply only to Col B
      .build();
    rules.push(uniqueDebaterAffRule);

    // Similar check for the Neg column (C)
    const uniqueDebaterNegFormula = `=AND($C2<>"${CONFIG.ROLES.BYE}", $C2<>"", OR(COUNTIFS($C$2:$C, $C2, $A$2:$A, $A2)>1, COUNTIFS($B$2:$B, $C2, $A$2:$A, $A2)>0))`;
    const uniqueDebaterNegRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueDebaterNegFormula)
      .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
      .setRanges([sheet.getRange(`C2:C${sheet.getMaxRows()}`)]) // Apply only to Col C
      .build();
    rules.push(uniqueDebaterNegRule);
  }

  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}


/**
 * Appends the generated match results to the AGGREGATE_HISTORY sheet.
 * This allows the Match Summary formulas to update automatically (PRD 5.2).
 * Handles panel judging by recording separate entries for each judge.
 * @param {Array<object>} pairings - The final pairings.
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {string} dateString - The date of the matches.
 */
function updateAggregateHistory(pairings, debateType, dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY);

  // If the history sheet somehow doesn't exist, we must recreate it to prevent failure.
  if (!historySheet) {
    const newHistorySheet = createTab(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY, getAggregateHistoryHeaders());
    newHistorySheet.hideSheet();
    return updateAggregateHistory(pairings, debateType, dateString); // Retry
  }

  const historyData = [];

  // Headers: Date, Type, Round, Debater Name, Is Ironman, Role, Opponent Team, Judge, Room

  // Modified to handle multiple judges (paneling) and ensure Opponent Team name is recorded.
  const appendHistory = (team, role, opponentTeamName, pairing) => {
    const roomName = pairing.Room ? pairing.Room.name : '';
    const judges = pairing.Judges || [];

    team.members.forEach(memberName => {
      // If there are judges (standard match), record one entry PER JUDGE for accurate stats.
      if (judges.length > 0) {
        judges.forEach(judge => {
          historyData.push([
            dateString,
            debateType,
            pairing.Round,
            memberName,
            team.isIronman,
            role,
            opponentTeamName, // The name of the opposing team
            judge.name,
            roomName
          ]);
        });
      } else {
        // If there are no judges (e.g., a BYE), record one entry with a blank judge.
        historyData.push([
          dateString,
          debateType,
          pairing.Round,
          memberName,
          team.isIronman,
          role,
          opponentTeamName, // BYE or similar
          '', // Judge
          roomName
        ]);
      }
    });
  };

  pairings.forEach(pairing => {
    if (pairing.isBye) {
      // Aff got the BYE
      appendHistory(pairing.Aff, CONFIG.ROLES.BYE, CONFIG.ROLES.BYE, pairing);
    } else {
      // Standard match. Pass the name of the opposing team.
      appendHistory(pairing.Aff, CONFIG.ROLES.AFF, pairing.Neg.name, pairing);
      appendHistory(pairing.Neg, CONFIG.ROLES.NEG, pairing.Aff.name, pairing);
    }
  });

  if (historyData.length > 0) {
    const nextRow = historySheet.getLastRow() + 1;
    historySheet.getRange(nextRow, 1, historyData.length, historyData[0].length).setValues(historyData);
  }
}

// =============================================================================
// Utility Functions
// =============================================================================

/**
 * Sorts the sheets according to the requirements (PRD 5.1).
 * Permanent tabs first, then generated tabs (newest first, LD before TP), then hidden tabs last.
 */
function sortSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sheets = ss.getSheets();

  const getSheetPriority = (sheetName) => {
    // 1. Permanent Tabs
    const permanentIndex = CONFIG.PERMANENT_TABS_ORDER.indexOf(sheetName);
    if (permanentIndex !== -1) {
      return { type: 'Permanent', order: permanentIndex, date: null };
    }

    // 2. Hidden/Aggregate Tabs (always last)
    if (sheetName === CONFIG.SHEET_NAMES.AGGREGATE_HISTORY) {
      return { type: 'Hidden', order: 9999, date: null };
    }

    // 3. Generated Match Tabs (e.g., TP 2025-07-29)
    const match = sheetName.match(/^(TP|LD) (\d{4}-\d{2}-\d{2})$/);
    if (match) {
      const debateType = match[1];
      const date = match[2];
      // LD (1) appears before TP (2) (PRD 5.1)
      const typeOrder = debateType === 'LD' ? 1 : 2;
      return { type: 'Generated', order: typeOrder, date: date };
    }

    // 4. Other tabs (e.g., manually created, or legacy aggregation sheets)
    return { type: 'Other', order: 9998, date: null };
  };

  const sortedSheets = sheets.sort((a, b) => {
    const prioA = getSheetPriority(a.getName());
    const prioB = getSheetPriority(b.getName());

    if (prioA.type !== prioB.type) {
      const typeOrder = ['Permanent', 'Generated', 'Other', 'Hidden'];
      return typeOrder.indexOf(prioA.type) - typeOrder.indexOf(prioB.type);
    }

    if (prioA.type === 'Generated') {
      // Sort by Date (Newest first - descending)
      if (prioA.date !== prioB.date) {
        return prioB.date.localeCompare(prioA.date);
      }
      // Sort by Type (LD before TP - ascending)
      return prioA.order - prioB.order;
    }

    // Sort Permanent/Other/Hidden by predefined order
    return prioA.order - prioB.order;
  });

  // Move sheets into the sorted order (Apps Script requires moving them one by one)
  // We need to activate the sheet before moving it.
  const activeSheet = ss.getActiveSheet(); // Remember the active sheet
  for (let i = 0; i < sortedSheets.length; i++) {
    // Use try-catch in case the sheet is hidden and cannot be activated
    try {
      ss.setActiveSheet(sortedSheets[i]);
      // moveActiveSheet uses 1-based indexing
      ss.moveActiveSheet(i + 1);
    } catch (e) {
      Logger.log(`Could not move sheet ${sortedSheets[i].getName()}: ${e.message}`);
    }
  }
  // Try restoring the original active sheet if possible
  try {
    // Check if the active sheet reference is valid and the sheet is not hidden
    if (activeSheet && !activeSheet.isSheetHidden()) {
      ss.setActiveSheet(activeSheet);
    }
  } catch (e) {
    // Ignore if the original active sheet cannot be activated
  }
}

/**
 * Shuffles an array in place using the Fisher-Yates algorithm.
 * @param {Array} array - The array to shuffle.
 * @returns {Array} The shuffled array.
 */
function shuffleArray(array) {
  let currentIndex = array.length, randomIndex;

  while (currentIndex !== 0) {
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex--;
    [array[currentIndex], array[randomIndex]] = [
      array[randomIndex], array[currentIndex]];
  }
  return array;
}
