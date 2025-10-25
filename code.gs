/**
 * Automated Debate Pairing Tool
 * Google Apps Script
 *
 * @fileoverview A tool for automating the pairing of debate matches (TP and LD) within a Google Sheet.
 * This script manages rosters, availability, pairing optimization based on hard/soft constraints,
 * historical tracking, and output generation.
 * All logic is contained in this single file as required by PRD 6.
 */

// =============================================================================
// Configuration & Constants
// =============================================================================

// Central configuration object defining sheet names, types, options, and styles.
// Modifying these values changes the behavior and appearance of the tool.
const CONFIG = {
  // Defines the names of the tabs used by the system.
  SHEET_NAMES: {
    AVAILABILITY: 'Availability',
    SUMMARY: 'Match Summary',
    DEBATERS: 'Debaters',
    JUDGES: 'Judges',
    ROOMS: 'Rooms',
    // Hidden sheet for data aggregation (PRD 5.2). This sheet stores the raw data used by the Match Summary formulas.
    // The system is designed to only use this specific sheet name. It must not be manually edited.
    AGGREGATE_HISTORY: 'AGGREGATE_HISTORY_DO_NOT_EDIT'
  },
  // The supported debate formats.
  DEBATE_TYPES: {
    TP: 'TP', // Team Policy
    LD: 'LD'  // Lincoln-Douglas
  },
  // Options for the Availability tab dropdown.
  RSVP_OPTIONS: ["Yes", "No", "Not responded"],
  // Definitions for roles within a debate.
  ROLES: {
    AFF: 'Aff', // Affirmative
    NEG: 'Neg', // Negative
    BYE: 'BYE', // Indicates a participant received a BYE
    IRONMAN_SUFFIX: '(IRONMAN)' // Suffix added to TP debaters competing alone (PRD 5.4)
  },
  // Visual styling constants (PRD 6.1).
  STYLES: {
    HEADER_BG: '#4a86e8', // Blue background for headers
    HEADER_FONT_COLOR: '#ffffff',
    ERROR_HIGHLIGHT: '#f4c7c3', // Light red for critical errors (e.g., roster conflicts, logic failures)
    WARNING_HIGHLIGHT: '#fff2cc', // Light yellow for duplicates in match sheets (e.g., double-booked room/judge)
    FONT: 'Arial'
  },
  // Required order for permanent tabs (PRD 5.1). The sortSheets function enforces this order.
  PERMANENT_TABS_ORDER: ['Availability', 'Match Summary', 'Debaters', 'Judges', 'Rooms'],
};

// =============================================================================
// Menu & Triggers
// =============================================================================

/**
 * A special Apps Script trigger that runs automatically when the spreadsheet is opened.
 * It adds the custom "Club Admin" menu to the Google Sheets UI (PRD 5.3).
 * It also attempts to refresh formulas and formatting on load.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Check if UI context exists (it might not in some automated executions).
  if (ui) {
    ui.createMenu('Club Admin')
      // Main workflow items
      .addItem('Generate LD Matches (2 Rounds)', 'generateLdMatches')
      .addItem('Generate TP Matches', 'generateTpMatches')
      .addSeparator()
      // Weekly Cleanup
      .addItem('Clear RSVPs for Next Week', 'clearRsvps')
      .addSeparator()
      // Setup and Maintenance
      .addItem('Initialize Sheet (Setup)', 'initializeSheet')
      // Added for maintenance: ensures formulas and formatting are correctly applied if they get corrupted.
      .addItem('Refresh Formulas & Formatting', 'forceUpdateSheets')
      .addToUi();
  }

  // Ensure formatting and formulas are applied on open for robustness.
  try {
    // We force update on open to clear potential #REF errors. 
    // This happens if the script logic (formulas) changed, or if array formulas are blocked from expanding by user data.
    // This makes the sheet self-healing upon opening.
    forceUpdateSheets(true);
  } catch (e) {
    // Silently fail if initialization hasn't happened yet (e.g., required sheets don't exist).
    Logger.log("Could not apply formatting/formulas onOpen (sheet might not be initialized): " + e.message);
  }
}

/**
 * Helper function to forcibly update formulas and formatting. Useful if sheets get corrupted or after script updates.
 * This function ensures that the latest definitions for formulas, validation, and formatting are applied to the spreadsheet.
 * @param {boolean} [isAutomatic=false] - True if called automatically (e.g., onOpen), false if called manually by user.
 */
function forceUpdateSheets(isAutomatic = false) {
  try {
    // Force update formulas (true parameter) clears the ranges first to prevent #REF errors and apply latest logic.
    applyFormulas(true, isAutomatic);
    formatAllSheets();
    applyDataValidations();
    applyAllConditionalFormatting();

    // Only show toast (small notification) if called manually by the user for feedback.
    if (!isAutomatic) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) ss.toast("Formulas and formatting refreshed.", "Maintenance", 3);
    }
  } catch (e) {
    // Log the error for debugging.
    Logger.log("Error during forced update: " + e.message);
    // Only show alert (popup dialog) if called manually and an error occurred, avoiding alerts onOpen.
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
 * Initializes the spreadsheet structure from scratch.
 * Sets up required tabs, headers, sample data, formulas, and formatting (PRD 4.1, 5.3).
 * This function is designed to be run once when the spreadsheet is first created.
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if already initialized (PRD 4.1). Prevents accidental data loss.
  // We check for the existence of the Debaters tab as a proxy for initialization status.
  if (ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS)) {
    ui.alert('Initialization Error', 'The sheet appears to be already initialized. Setup aborted to prevent data loss.', ui.ButtonSet.OK);
    return;
  }

  // Create permanent tabs (PRD 5.1) using the helper function to ensure headers are set.
  createTab(CONFIG.SHEET_NAMES.DEBATERS, getDebaterHeaders());
  createTab(CONFIG.SHEET_NAMES.JUDGES, getJudgeHeaders());
  createTab(CONFIG.SHEET_NAMES.ROOMS, getRoomHeaders());
  createTab(CONFIG.SHEET_NAMES.AVAILABILITY, getAvailabilityHeaders());
  createTab(CONFIG.SHEET_NAMES.SUMMARY, getSummaryHeaders());

  // Create and hide the aggregation sheet (PRD 5.2). This is the system database.
  // The createTab function ensures that if this specific sheet already exists (e.g., from a failed previous attempt), it is reused.
  const aggSheet = createTab(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY, getAggregateHistoryHeaders());
  aggSheet.hideSheet(); // Hide from users as it's infrastructure.

  // Insert sample data (PRD 5.6) to demonstrate functionality.
  insertSampleData();

  // Apply formulas, formatting and validation. We use force update (true) during initialization to ensure a clean start.
  forceUpdateSheets(true);

  // Sort tabs into the correct initial order (PRD 5.1).
  sortSheets();

  // Activate the main landing tab (Availability) for the user.
  ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY).activate();

  // Clean up default "Sheet1" if it exists, which Google Sheets often creates by default.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // Handle potential error if Sheet1 cannot be deleted (e.g. if it's the only visible sheet).
      Logger.log("Could not delete Sheet1 during initialization: " + e.message);
    }
  }
}

/**
 * Helper to create a tab if it doesn't exist and set its headers.
 * This function is idempotent: it prevents the creation of duplicate tabs and ensures header integrity.
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
  // Ensure headers are set (important if the sheet existed but was corrupted or empty).
  // This writes the headers to the first row.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

// Define Headers for all sheets (Schemas defined in PRD 5.2)
function getDebaterHeaders() { return ['Name', 'Debate Type', 'Partner', 'Hard Mode']; }
function getJudgeHeaders() { return ['Name', 'Children\'s Names', 'Debate Type']; }
function getRoomHeaders() { return ['Room Name', 'Debate Type']; }
function getAvailabilityHeaders() { return ['Participant', 'Attending?']; }
function getSummaryHeaders() {
  // Headers reflect tracking statistics required by PRD 5.2, including top opponents and judges.
  return ['Debater Name', 'Debate Type', 'Total Matches', 'Aff Matches', 'Neg Matches', 'BYEs', 'Ironman Matches',
    '#1 Judge', '#2 Judge', '#3 Judge', '#1 Opponent', '#2 Opponent', '#3 Opponent'];
}
// Aggregation Schema: Used for historical tracking and summary calculations.
// This schema is denormalized (one row per debater per judge per match) to support accurate statistics when panels are used (PRD 5.2 Note).
function getAggregateHistoryHeaders() {
  // Schema Columns:
  // A-C: Match identifier (Date, Type, Round)
  // D-F: Debater info (Name, Is Ironman, Role)
  // G: Opponent Team (Crucial for TP tracking and opponent frequency stats)
  // H-I: Resources (Judge, Room)
  return ['Date', 'Type', 'Round', 'Debater Name', 'Is Ironman', 'Role', 'Opponent Team', 'Judge', 'Room'];
}
// Headers for generated match sheets. Judge(s) column supports comma-delimited panels.
function getTpMatchHeaders() { return ['Aff Team', 'Neg Team', 'Judge(s)', 'Room']; }
function getLdMatchHeaders() { return ['Round', 'Aff Debater', 'Neg Debater', 'Judge(s)', 'Room']; }


/**
 * Inserts the sample data defined in PRD 5.6.
 * This provides initial data for testing and demonstration purposes.
 */
function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // Debaters Data (LD and TP)
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
    // TP (Including a name with an apostrophe ("Sandra Day O'Connor") specifically to test formula robustness (SQL escaping in QUERY)).
    ['Noam Chomsky', 'TP', 'Michel Foucault', 'Yes'], ['Michel Foucault', 'TP', 'Noam Chomsky', 'Yes'],
    ['Harlow Shapley', 'TP', 'Heber Curtis', 'No'], ['Heber Curtis', 'TP', 'Harlow Shapley', 'No'],
    ['Muhammad Ali', 'TP', "Sandra Day O'Connor", 'Yes'], ["Sandra Day O'Connor", 'TP', 'Muhammad Ali', 'Yes'],
    ['Richard Nixon', 'TP', 'Nikita Khrushchev', 'No'], ['Nikita Khrushchev', 'TP', 'Richard Nixon', 'No'],
    ['Thomas Henry Huxley', 'TP', 'Samuel Wilberforce', 'Yes'], ['Samuel Wilberforce', 'TP', 'Thomas Henry Huxley', 'Yes'],
    ['John F. Kennedy', 'TP', 'David Frost', 'No'], ['David Frost', 'TP', 'John F. Kennedy', 'No'],
    ['Bob Dole', 'TP', 'Bill Clinton', 'No'], ['Bill Clinton', 'TP', 'Bob Dole', 'No'],
  ];
  const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
  // Only insert if the sheet is empty (getLastRow() < 2 means only the header exists).
  if (debatersSheet && debatersSheet.getLastRow() < 2) {
    debatersSheet.getRange(2, 1, debatersData.length, debatersData[0].length).setValues(debatersData);
  }

  // Judges Data (includes parent-child relationships for conflict testing)
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

  // Rooms Data
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
 * Applies Google Sheets formulas to the dynamic tabs (Availability and Match Summary).
 * This function ensures the spreadsheet logic remains intact and updates automatically based on data changes (PRD 5.2).
 * @param {boolean} [forceUpdate=false] - If true, clears existing formula ranges before reapplying. 
 *                                        This is crucial for preventing #REF errors when array formulas cannot expand.
 */
function applyFormulas(forceUpdate = false, isAutomatic = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // 1. Availability Tab (PRD 5.2)
  const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (availabilitySheet) {

    // CRITICAL: When using ARRAYFORMULA or similar expanding formulas (like UNIQUE/SORT/FILTER),
    // we must ensure the range below the formula is clear, otherwise it results in a #REF! error.
    if (forceUpdate && availabilitySheet.getMaxRows() > 1) {
      // Clear A2:A to ensure the main array formula in A2 can expand cleanly.
      availabilitySheet.getRange(2, 1, availabilitySheet.getMaxRows() - 1, 1).clearContent();
    }

    // Formula to combine and sort unique names from Debaters and Judges (PRD 4.1).
    // Uses array literals {} to stack the ranges vertically (semicolon separator).
    // FILTER removes blanks, UNIQUE removes duplicates, SORT alphabetizes.
    const participantFormula = `=SORT(UNIQUE(FILTER({${CONFIG.SHEET_NAMES.DEBATERS}!A2:A; ${CONFIG.SHEET_NAMES.JUDGES}!A2:A}, {${CONFIG.SHEET_NAMES.DEBATERS}!A2:A; ${CONFIG.SHEET_NAMES.JUDGES}!A2:A}<>"")))`;
    availabilitySheet.getRange('A2').setFormula(participantFormula);

    // Initialize default RSVP status. This should only run on a MANUAL refresh or if the column is empty.
    // It should NOT run during the automatic onOpen trigger if data already exists.
    if ((forceUpdate && !isAutomatic) || availabilitySheet.getRange('B2').getValue() === "") {
      SpreadsheetApp.flush(); // Ensure the participant formula populates before setting defaults.
      setRsvpDefaults(availabilitySheet);
    }
  }


  // 2. Match Summary Tab (PRD 5.2) - Must be formula-driven, summarizing data from AGGREGATE_HISTORY.
  const summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  if (!summarySheet) return;

  const HISTORY = CONFIG.SHEET_NAMES.AGGREGATE_HISTORY;
  const maxRows = 500; // Apply formulas down to a fixed large range to accommodate growth.

  // If force updating, clear the entire data range (A2:M...).
  // This prevents #REF! errors caused by array expansion issues when formulas change or conflict with existing data.
  if (forceUpdate && summarySheet.getMaxRows() > 1) {
    // Use getMaxColumns() to ensure we clear potential DEBUG columns (like column N) as well, not just the standard headers.
    const lastCol = Math.max(getSummaryHeaders().length, summarySheet.getMaxColumns());
    summarySheet.getRange(2, 1, summarySheet.getMaxRows() - 1, lastCol).clearContent();
    SpreadsheetApp.flush(); // Ensure clearing is complete before applying new formulas.
  }

  // A2: Debater Name and B2: Debate Type (populated automatically from Debaters tab).
  // This is an array formula that populates columns A and B simultaneously.
  const summaryFormulaA2 = `=SORT(FILTER({${CONFIG.SHEET_NAMES.DEBATERS}!A2:B}, ${CONFIG.SHEET_NAMES.DEBATERS}!A2:A<>""), 1, TRUE)`;
  summarySheet.getRange('A2').setFormula(summaryFormulaA2);

  // Flush to ensure A2:B populates before subsequent formulas rely on $A2.
  SpreadsheetApp.flush();


  // Formulas updated to handle multiple history entries per match (due to judge panels, PRD 5.2 Note).
  // The history sheet is denormalized (one row per debater per judge).
  // To count actual matches, we must count unique combinations of Date(A)/Type(B)/Round(C).
  // We use IFERROR(ROWS(UNIQUE(FILTER(...))), 0) for robust counting that returns 0 if no matches are found.

  // C2: Total Matches (Excluding BYEs)
  const totalMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!F:F<>"${CONFIG.ROLES.BYE}"))), 0), "")`;
  // Apply the formula to the entire column range C2:C500.
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

  // G2: Ironman Matches (Column E in History is TRUE/FALSE boolean)
  const ironmanMatchesFormula = `=IF($A2<>"", IFERROR(ROWS(UNIQUE(FILTER('${HISTORY}'!A:C, '${HISTORY}'!D:D=$A2, '${HISTORY}'!E:E=TRUE))), 0), "")`;
  summarySheet.getRange('G2:G' + maxRows).setFormula(ironmanMatchesFormula);

  // --- Top 3 Formulas (Judges and Opponents) ---
  // These use the QUERY function (Google Visualization API Query Language) for aggregation and sorting.

  // CRITICAL: We must escape single quotes in names (e.g., O'Connor) before injecting them into the QUERY string.
  // The Query Language standard requires replacing a single quote (') with two single quotes ('').
  // We use SUBSTITUTE($A2, "'", "''") within the formula construction for robustness and security.
  const safeNameA2 = `SUBSTITUTE($A2, "'", "''")`;

  // H2: Top 3 Judges (H, I, J)
  // Uses a double QUERY structure for aggregation, sorting, and limiting.

  // Inner QUERY: Calculates the frequency (COUNT) of each judge (H) for the current debater (D).
  // We use LABEL ... '' to ensure the aggregation columns (COUNT(H)) do not return headers, which can disrupt the outer QUERY if results are sparse.
  const topJudgesQuery = `"SELECT H, COUNT(H) WHERE D = '"&${safeNameA2}&"' AND H IS NOT NULL AND H <> '' GROUP BY H LABEL H '', COUNT(H) ''"`;

  // Outer QUERY: Takes the result of the inner query, sorts by frequency (Col2 DESC), limits to 3, and selects only the name (Col1).
  // TRANSPOSE converts the resulting column into a row (populating H, I, J).
  const topJudgesFormula = `=IFERROR(IF($A2<>"", TRANSPOSE(QUERY(QUERY('${HISTORY}'!D:H, ${topJudgesQuery}, 0), "Select Col1 ORDER BY Col2 DESC LIMIT 3", 0)), ""), "")`;
  summarySheet.getRange('H2:H' + maxRows).setFormula(topJudgesFormula);

  // K2: Top 3 Opponents (K, L, M)
  // This identifies the top 3 opposing teams (TP) or individuals (LD).

  // Challenge: Because the history sheet is denormalized for panels, a simple QUERY on the raw history would count opponents multiple times per match.
  // Solution: We must first use UNIQUE(A:G) to get a list of distinct matches before running the frequency analysis.

  // We use a robust double-QUERY structure because aggregating (GROUP BY/COUNT) on derived arrays (like UNIQUE) with mixed data types can be unreliable in Google Sheets if done in a single QUERY.

  // Inner QUERY data source is the UNIQUE match history (A:G). We only need up to G (Opponent Team).
  const uniqueHistorySource = `UNIQUE('${HISTORY}'!A:G)`;

  // Construct the inner QUERY string. We must use ColX notation because the source is an array (result of UNIQUE), not a sheet range.
  // Col4=Debater Name (D), Col7=Opponent Team (G).
  // This inner query calculates the frequency (COUNT) of each opponent team, excluding BYEs.
  const topOpponentsInnerQuery = `"SELECT Col7, COUNT(Col7) WHERE Col4 = '"&${safeNameA2}&"' AND Col7 IS NOT NULL AND Col7 <> '' AND Col7 <> '${CONFIG.ROLES.BYE}' GROUP BY Col7 LABEL Col7 '', COUNT(Col7) ''"`;

  // Construct the full formula with the outer QUERY for sorting, limiting, and transposing.
  // We explicitly set the headers parameter to 0 for both queries for clarity and robustness.
  const topOpponentsFormula = `=IFERROR(IF($A2<>"", TRANSPOSE(QUERY(QUERY(${uniqueHistorySource}, ${topOpponentsInnerQuery}, 0), "SELECT Col1 ORDER BY Col2 DESC LIMIT 3", 0)), ""), "")`;

  summarySheet.getRange('K2:K' + maxRows).setFormula(topOpponentsFormula);
}

/**
 * Applies standard formatting to all sheets (PRD 6.1).
 * Ensures a consistent look and feel across the application.
 */
function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Check if ss is null (can happen in some execution contexts)
  if (!ss) return;
  const sheets = ss.getSheets();

  // Iterate over every sheet in the spreadsheet.
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Skip hidden sheet formatting (it's infrastructure, formatting is irrelevant).
    if (sheetName === CONFIG.SHEET_NAMES.AGGREGATE_HISTORY) return;

    // Handle potentially empty or corrupted sheets gracefully.
    if (sheet.getMaxColumns() === 0 || sheet.getMaxRows() === 0) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    // Defensive check: Ensure Match Summary headers are correct if the sheet somehow got corrupted.
    if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      createTab(CONFIG.SHEET_NAMES.SUMMARY, getSummaryHeaders());
    }

    // Define the header range (Row 1).
    const headerRange = sheet.getRange(1, 1, 1, lastCol);

    // Header Styling (PRD 6.1)
    headerRange.setFontWeight('bold')
      .setBackground(CONFIG.STYLES.HEADER_BG)
      .setFontColor(CONFIG.STYLES.HEADER_FONT_COLOR)
      .setHorizontalAlignment('center')
      // Ensure headers do not wrap (e.g., "Ironman Matches" in Match Summary).
      // Using CLIP strategy forces columns wider during auto-resize, improving readability.
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);


    // Freeze Header Row (PRD 5.1) - Keeps headers visible during scrolling.
    sheet.setFrozenRows(1);

    // Apply overall font and borders to the entire data range.
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const fullRange = sheet.getDataRange();
      fullRange.setFontFamily(CONFIG.STYLES.FONT)
        .setBorder(true, true, true, true, true, true);
    }


    // Apply Banding (Alternating Row Colors) (PRD 6.1) for readability.
    if (sheet.getMaxRows() > 1) {
      const dataRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
      // Remove existing banding first to ensure consistency.
      sheet.getBandings().forEach(banding => banding.remove());
      // Apply new banding, including the header row.
      dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    }

    // Specific alignment adjustments (UX improvement). Center-aligning non-name columns.
    if (sheetName === CONFIG.SHEET_NAMES.DEBATERS || sheetName === CONFIG.SHEET_NAMES.JUDGES || sheetName === CONFIG.SHEET_NAMES.ROOMS || sheetName === CONFIG.SHEET_NAMES.AVAILABILITY) {
      if (sheet.getMaxRows() > 1) {
        // Align columns from the second column onwards.
        sheet.getRange(2, 2, sheet.getMaxRows() - 1, sheet.getLastColumn() - 1).setHorizontalAlignment('center');
      }
    } else if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      // Center-align the statistical columns (B-G) in the Summary tab.
      if (sheet.getMaxColumns() >= 7) { // Ensure columns B-G exist
        sheet.getRange('B:G').setHorizontalAlignment('center');
      }
    }

    // Auto-sizing Columns (PRD 6.1). This should now resize based on the clipped header text.
    // We limit the auto-resize to the standard columns to avoid resizing potential DEBUG columns that might be very wide.
    let resizeCols = lastCol;
    if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
      // Ensure we don't try to resize beyond the actual columns if the standard header count is larger.
      resizeCols = Math.min(lastCol, getSummaryHeaders().length);
    }

    if (resizeCols > 0) {
      sheet.autoResizeColumns(1, resizeCols);
    }
  });
}

/**
 * Applies dropdown data validations to input fields.
 * This restricts user input to valid options defined in the CONFIG.
 */
function applyDataValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // Availability Tab - Attending? Dropdown (PRD 5.2)
  const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (availabilitySheet) {
    // Create a validation rule requiring values from CONFIG.RSVP_OPTIONS. setAllowInvalid(false) strictly enforces the list.
    const rsvpRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.RSVP_OPTIONS, true).setAllowInvalid(false).build();
    // Apply to the entire 'Attending?' column (B).
    availabilitySheet.getRange('B2:B').setDataValidation(rsvpRule);
  }

  // Debaters Tab - Type (B) and Hard Mode (D)
  const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
  if (debatersSheet) {
    const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(CONFIG.DEBATE_TYPES), true).build();
    debatersSheet.getRange('B2:B').setDataValidation(typeRule);
    const hardModeRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build();
    debatersSheet.getRange('D2:D').setDataValidation(hardModeRule);
  }

  // Judges (C) and Rooms (B) Tab - Type
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
 * These rules provide continuous visual feedback on data integrity issues.
 * Note: This function defines ALL rules for a sheet as setConditionalFormatRules overwrites existing rules.
 */
function applyAllConditionalFormatting() {
  // We use try-catch because this runs onOpen, and sheets might not exist yet (e.g., before initialization).
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    // 1. Availability Tab
    const availabilitySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
    if (availabilitySheet) {
      const availabilityRules = [];
      const availabilityRange = availabilitySheet.getRange('B2:B');
      // Highlight blank RSVP (B2) if Participant (A2) exists (PRD 5.2).
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
      // Room names must be unique (PRD 5.5). Check if the count of the current room name (A2) in the whole column (A:A) is > 1.
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
      // A person cannot be both a Judge and a Debater (PRD 5.5).
      // Must use INDIRECT() as required by PRD 5.5 implementation note (necessary when referencing other sheets in CF).
      // Formula: Check if the Judge name ($A2) exists in the Debaters roster (A:A).
      const overlapRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2<>"", COUNTIF(INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A:A"), $A2)>0)`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([judgesRange])
        .build();
      judgesRules.push(overlapRule);

      // Note: Validating comma-delimited "Children's Names" existence is too complex for CF formulas.
      // This specific validation (ensuring children exist in the roster) is handled by the Apps Script pre-flight validation (validateRosters).

      judgesSheet.setConditionalFormatRules(judgesRules);
    }

    // 4. Debaters Tab
    const debatersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBATERS);
    if (debatersSheet && debatersSheet.getMaxRows() > 1) {
      const debatersRules = [];
      const debatersRange = debatersSheet.getRange('A2:D');

      // A person cannot be both a Judge and a Debater (Mirror check from the Judges tab).
      const overlapRuleDebater = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2<>"", COUNTIF(INDIRECT("${CONFIG.SHEET_NAMES.JUDGES}!A:A"), $A2)>0)`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(overlapRuleDebater);

      // Partnership Consistency (PRD 5.5). These rules apply only to TP debaters ($B2="TP").

      // Partner (C2) must exist in the Debaters roster (A:A).
      const partnerExistsRule = SpreadsheetApp.newConditionalFormatRule()
        // Formula: Check if Type is TP, Partner is listed, and MATCH fails to find the partner (ISNA).
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", ISNA(MATCH($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A:A"), 0)))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(partnerExistsRule);

      // Partnerships must be reciprocal.
      // Formula: Look up the partner ($C2) and check if their partner (Column 3 of the VLOOKUP) matches the original debater ($A2).
      // Note: We use INDIRECT here as well, although VLOOKUP on the same sheet might work, consistency is better for CF reliability.
      const reciprocalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", IFERROR(VLOOKUP($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A2:C"), 3, FALSE) <> $A2, FALSE))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(reciprocalRule);

      // Partners must have the same "Hard Mode" setting.
      // Formula: Look up the partner ($C2) and check if their Hard Mode (Column 4) matches the original debater's Hard Mode ($D2).
      const hardModeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($B2="TP", $C2<>"", IFERROR(VLOOKUP($C2, INDIRECT("${CONFIG.SHEET_NAMES.DEBATERS}!A2:D"), 4, FALSE) <> $D2, FALSE))`)
        .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
        .setRanges([debatersRange])
        .build();
      debatersRules.push(hardModeRule);

      debatersSheet.setConditionalFormatRules(debatersRules);
    }
  } catch (e) {
    // Log errors during CF application for debugging, but do not halt execution.
    Logger.log("Error applying conditional formatting: " + e.message);
  }
}


// =============================================================================
// Weekly Workflow Functions
// =============================================================================

/**
 * Clears the RSVPs in the Availability tab, setting them back to "Not responded" (PRD 4.2, 5.3).
 * This function is run after a meeting to prepare for the next week.
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
 * Ensures that the RSVP column (B) correctly reflects the dynamically generated Participant list (A).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Availability sheet.
 */
function setRsvpDefaults(sheet) {
  const lastRow = sheet.getLastRow();
  // Clear the existing RSVP column content (B2:B) first to ensure a clean state.
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).clearContent();
  }

  // If there are no participants, return early.
  if (lastRow < 2) return;

  // Read column A (Participants). This list is generated by a formula.
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  // Map the data to Column B defaults. We process this in memory for performance.
  const newRsvps = data.map(row => {
    // If Participant name (row[0]) exists, set RSVP to "Not responded", otherwise keep the cell blank.
    return [row[0] ? "Not responded" : ""];
  });

  // Write Column B in a single batch operation.
  if (newRsvps.length > 0) {
    sheet.getRange(2, 2, newRsvps.length, 1).setValues(newRsvps);
  }
}

/**
 * Main function to generate Team Policy (TP) matches.
 * Menu trigger entry point for TP.
 */
function generateTpMatches() {
  generateMatches(CONFIG.DEBATE_TYPES.TP);
}

/**
 * Main function to generate Lincoln-Douglas (LD) matches (2 rounds).
 * Menu trigger entry point for LD.
 */
function generateLdMatches() {
  generateMatches(CONFIG.DEBATE_TYPES.LD);
}

/**
 * Core workflow function for generating matches for a given type (PRD 5.3).
 * This orchestrates the entire process: validation, data retrieval, pairing optimization, and output generation.
 * @param {string} debateType - 'TP' or 'LD'.
 */
function generateMatches(debateType) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Use a try-catch block to handle unexpected errors gracefully (PRD 5.7).
  try {
    // 1. Run data integrity checks (PRD 5.3 Step 1).
    // Validates rosters (partnerships, conflicts, etc.) before attempting pairing.
    if (!validateRosters()) {
      return; // Validation function handles the alert UI. Execution stops here if invalid.
    }

    // 2. Get available participants and history (PRD 5.3 Step 2).
    // Read the entire match history from the hidden sheet.
    const history = getMatchHistory();
    // Determine who is available (RSVP='Yes') and filter resources (Debaters, Judges, Rooms) for the specific debate type.
    const resources = getAvailableResources(debateType, history);

    // Check if there are any debaters available before proceeding.
    if (resources.debaters.length === 0) {
      ss.toast(`No available debaters for ${debateType}. Cannot generate matches.`, 'Notice', 5);
      return;
    }

    // 3. Run the pairing logic (PRD 5.3 Step 3).
    let allPairings = [];

    if (debateType === CONFIG.DEBATE_TYPES.TP) {
      // TP specific logic: Form teams based on partnerships and handle Ironman cases (PRD 5.4).
      // formTpTeams ensures consistent alphabetical naming for permanent teams, which is crucial for history tracking.
      const teams = formTpTeams(resources.debaters);
      // Execute a single round of TP pairings.
      const pairings = executePairingRound(teams, resources.judges, resources.rooms, history, 1);
      allPairings = pairings;

    } else if (debateType === CONFIG.DEBATE_TYPES.LD) {
      // LD specific logic: Requires sequential 2-round pairing (PRD 5.4.1).

      // Convert LD debaters into a standardized "team" structure (of one person) for the pairing algorithm.
      const teams = resources.debaters.map(d => ({
        name: d.name,
        members: [d.name],
        hardMode: d.hardMode,
        isIronman: false, // Ironman concept only applies to TP.
        history: d.history
      }));

      // Round 1 (PRD 5.4.1 Step 1). Generate pairings based on historical data.
      const round1Pairings = executePairingRound(teams, resources.judges, resources.rooms, history, 1);
      allPairings.push(...round1Pairings);

      // Identify R1 BYE recipient. This person must be excluded from the R2 BYE to ensure fairness (PRD 5.4.1 Constraint).
      const r1ByePairing = round1Pairings.find(p => p.isBye);
      const r1ByeTeamName = r1ByePairing ? r1ByePairing.Aff.name : null;

      // Update History In-Memory (PRD 5.4.1 Step 2).
      // Simulate the results of Round 1 so Round 2 optimization considers these matchups (e.g., minimizing rematches).
      const historyR2 = updateHistoryInMemory(history, round1Pairings);

      // Round 2 (PRD 5.4.1 Step 3).
      // Resources (Judges/Rooms) can be reused (reshuffled), but we pass the updated history and the BYE exclusion.
      const round2Pairings = executePairingRound(teams, resources.judges, resources.rooms, historyR2, 2, r1ByeTeamName);
      allPairings.push(...round2Pairings);
    }

    // 4. Check for existing sheet (PRD 5.3 Step 4). Prevents accidental overwriting of today's results.
    const dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const sheetName = `${debateType} ${dateString}`;
    if (ss.getSheetByName(sheetName)) {
      ui.alert('Overwrite Protection', `A sheet named "${sheetName}" already exists. Delete it manually if you wish to regenerate pairings for today.`, ui.ButtonSet.OK);
      return;
    }

    // 5. Write pairings and update history (PRD 5.3 Step 5).
    // Create the new sheet and populate it with the generated pairings.
    const newSheet = createMatchSheet(sheetName, debateType, allPairings);
    // Append the results to the persistent history log.
    updateAggregateHistory(allPairings, debateType, dateString);

    // 6. Re-sort all tabs and apply formatting (PRD 5.3 Step 6).
    sortSheets(); // Ensures correct tab order (newest generated sheet first).

    // Re-apply formatting and formulas. We force update formulas to ensure Match Summary reflects the new data correctly (PRD 5.2).
    // This is crucial because the Match Summary relies on array formulas that need to recalculate and expand after the history data changes.
    forceUpdateSheets(true);

    // Apply validation rules (Conditional Formatting) to the newly generated match sheet (PRD 5.5).
    applyMatchSheetValidation(newSheet, debateType);

    // Ensure all pending spreadsheet operations (formulas recalculation, formatting) are completed.
    SpreadsheetApp.flush();

    // Activate the new sheet for the user.
    newSheet.activate();
    // Success feedback (PRD 5.7).
    ss.toast(`${debateType} Matches Generated Successfully.`, 'Success', 5);

  } catch (error) {
    // Catch any unexpected errors during the process.
    Logger.log(error.stack); // Log the full stack trace for debugging.
    // Display critical errors using alerts (PRD 5.7) as they indicate a failure in the pairing process (e.g., insufficient judges).
    if (ui) {
      ui.alert('Error Generating Matches', error.message, ui.ButtonSet.OK);
    } else {
      // Fallback if UI context is unavailable (e.g., triggered execution).
      Logger.log('Could not display UI alert: ' + error.message);
    }
  }
}

// =============================================================================
// Data Retrieval & Modeling
// =============================================================================

/**
 * Reads the roster data (Debaters, Judges, Rooms) from the respective sheets.
 * This function handles the raw data input and structures it into JavaScript objects.
 * @returns {object} An object containing arrays of debaters, judges, and rooms.
 */
function getRosterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Helper function to read data from a sheet, starting from the second row.
  const readSheetData = (sheetName, headersLength) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    // Read the data range in a single call for efficiency.
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, headersLength).getValues();
  };

  // Read raw data from the sheets.
  const debatersData = readSheetData(CONFIG.SHEET_NAMES.DEBATERS, 4);
  const judgesData = readSheetData(CONFIG.SHEET_NAMES.JUDGES, 3);
  const roomsData = readSheetData(CONFIG.SHEET_NAMES.ROOMS, 2);

  // Process and structure the data.
  // CRITICAL: Ensure names are trimmed (using .trim()) for data consistency, as trailing spaces can break lookups.
  const debaters = debatersData.map(row => ({
    name: String(row[0]).trim(),
    type: String(row[1]).trim(),
    partner: String(row[2]).trim(),
    // Convert "Yes"/"No" to boolean for easier logic processing.
    hardMode: String(row[3]).trim() === 'Yes'
  })).filter(d => d.name); // Filter out empty rows.

  const judges = judgesData.map(row => ({
    name: String(row[0]).trim(),
    // Parse comma-delimited children names (PRD 5.2). Split by comma, trim spaces, and filter out empty strings.
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
 * @returns {Set<string>} A set of names of participants (Debaters and Judges) marked as 'Yes'.
 */
function getAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  const attending = new Set(); // Use a Set for efficient lookups (O(1) average time complexity).

  if (!sheet || sheet.getLastRow() < 2) return attending;

  // Read Participant (A) and Attending? (B) columns.
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(row => {
    // Check if the RSVP status is exactly 'Yes'.
    if (row[1] === 'Yes') {
      attending.add(String(row[0]).trim());
    }
  });

  return attending;
}

/**
 * Reads the AGGREGATE_HISTORY sheet and structures the data for optimization lookups.
 * This function transforms the raw history log into structured statistics used by the pairing algorithm.
 * Note: This function handles the history structure where multiple entries exist for one match if panels are used (denormalized data).
 * @returns {object} A history object containing detailed stats per debater and judge history.
 * Structure: { debaters: { [name]: { byes, opponents: {}, judges: {}, affCount, negCount } }, judges: { [name]: { rooms: {} } } }
 */
function getMatchHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY);
  // Initialize the history structure.
  const history = {
    debaters: {},
    judges: {}
  };

  // If the history sheet is empty or doesn't exist, return the empty structure.
  if (!sheet || sheet.getLastRow() < 2) return history;

  // Read the history data.
  // Headers: Date(0), Type(1), Round(2), Debater Name(3), Is Ironman(4), Role(5), Opponent Team(6), Judge(7), Room(8)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  // We use a set to track unique matches processed for each debater.
  // This is CRITICAL because the history sheet is denormalized (multiple rows per match due to panels).
  // We must avoid inflating match counts (Aff/Neg/BYE/Opponents) by processing the same match multiple times.
  const processedMatches = new Set();

  data.forEach(row => {
    const date = String(row[0]);
    const type = String(row[1]);
    const round = String(row[2]);
    // Ensure names read from history are trimmed for consistency.
    const debaterName = String(row[3]).trim();

    // Defensive check: Skip processing if the row is missing the debater name (corrupted data).
    if (!debaterName) return;

    const role = String(row[5]);
    // Opponent Team name is stored in Col G (6). This correctly handles TP teams and LD individuals.
    const opponentTeam = String(row[6]).trim();
    const judge = String(row[7]).trim();
    const room = String(row[8]).trim();

    // Define a unique key for the match from the perspective of the debater.
    const matchKey = `${debaterName}|${date}|${type}|${round}`;

    // Initialize debater history structure if not present.
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

        // Track the opposing team name (used for Soft Constraint 2: Minimize rematches).
        if (opponentTeam && opponentTeam !== CONFIG.ROLES.BYE) {
          stats.opponents[opponentTeam] = (stats.opponents[opponentTeam] || 0) + 1;
        }
      }
      // Mark this match as processed for this debater.
      processedMatches.add(matchKey);
    }

    // Process judge/room stats. We ALWAYS process these, as each row represents a unique judge interaction (Soft Constraint 3).
    if (role !== CONFIG.ROLES.BYE) {
      if (judge) {
        // Increment the count of times this debater has been judged by this judge.
        stats.judges[judge] = (stats.judges[judge] || 0) + 1;

        // Track which rooms judges have used (Soft Constraint 4).
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
 * This function prepares the input data for the pairing algorithm.
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {object} history - The match history object generated by getMatchHistory().
 * @returns {object} Available debaters, judges, and rooms, enriched with their historical data.
 */
function getAvailableResources(debateType, history) {
  const { debaters, judges, rooms } = getRosterData();
  const attending = getAttendance();

  // Filter available debaters by type and attendance, then enrich with history.
  const availableDebaters = debaters
    .filter(item => item.type === debateType && attending.has(item.name))
    .map(item => {
      // Attach history object. Provide default empty history if the debater is new.
      const itemHistory = (history.debaters[item.name] || { byes: 0, opponents: {}, judges: {}, affCount: 0, negCount: 0 });
      return { ...item, history: itemHistory };
    });

  // Filter available judges by type and attendance, then enrich with room history.
  const availableJudges = judges
    .filter(item => item.type === debateType && attending.has(item.name))
    .map(judge => {
      // Attach room history (PRD 5.4 Soft Constraint 4).
      judge.roomHistory = (history.judges && history.judges[judge.name]) ? history.judges[judge.name].rooms : {};
      return judge;
    });

  // Filter available rooms. Rooms must match the debate type (Hard Constraint 4).
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
 * @returns {Array<object>} List of formed teams ready for pairing.
 */
function formTpTeams(debaters) {
  const teams = [];
  const processed = new Set(); // Tracks debaters already assigned to a team.

  // Iterate through available debaters to form teams.
  debaters.forEach(debater => {
    if (processed.has(debater.name)) return;

    // Find the partner in the available list (must also be attending).
    const partner = debaters.find(d => d.name === debater.partner);

    if (partner) {
      // Full team (both partners present). 

      // Ensure consistent naming convention (alphabetical sort).
      // This is VITAL so that "A / B" is always the same team name, regardless of which partner the loop encountered first. This ensures reliable opponent tracking.
      const members = [debater.name, partner.name].sort();
      // The team name format ("Name1 / Name2") is critical for opponent tracking in Match Summary.

      teams.push({
        name: `${members[0]} / ${members[1]}`,
        members: members,
        hardMode: debater.hardMode, // Hard Mode is consistent due to validation.
        isIronman: false,
        // Assign team history. We use the history of the member with fewer BYEs for fair BYE prioritization.
        // This ensures that if one partner has significantly more BYEs, the team inherits the history of the less-burdened partner for this calculation.
        history: debater.history.byes <= partner.history.byes ? debater.history : partner.history
      });
      processed.add(debater.name);
      processed.add(partner.name);
    } else {
      // Ironman team (partner is missing or not attending) (PRD 5.4 Special Cases).
      teams.push({
        name: `${debater.name} ${CONFIG.ROLES.IRONMAN_SUFFIX}`,
        members: [debater.name],
        hardMode: debater.hardMode,
        isIronman: true,
        history: debater.history // Uses the individual's history.
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
 * This is a pre-flight check that stops the pairing process if critical data issues are found.
 * It complements the Conditional Formatting by providing a detailed error report.
 * @returns {boolean} True if valid, false otherwise.
 */
function validateRosters() {
  const { debaters, judges, rooms } = getRosterData();
  const ui = SpreadsheetApp.getUi();
  let errors = [];

  // Create sets for efficient name lookups.
  const debaterNames = new Set(debaters.map(d => d.name));
  const judgeNames = new Set(judges.map(j => j.name));

  // 1. Check for Judge/Debater Overlap (PRD 5.5).
  debaterNames.forEach(name => {
    if (judgeNames.has(name)) {
      errors.push(`"${name}" is listed as both a Debater and a Judge.`);
    }
  });

  // 2. Validate Judge's Children (PRD 5.5).
  // Ensures that all names listed in "Children's Names" exist in the Debaters roster.
  judges.forEach(judge => {
    judge.children.forEach(childName => {
      if (!debaterNames.has(childName)) {
        errors.push(`Judge "${judge.name}" lists child "${childName}", who is not in the Debaters roster.`);
      }
    });
  });

  // 3. Validate TP Partnerships (Permanent Teams) (PRD 5.5).
  const tpDebaters = debaters.filter(d => d.type === CONFIG.DEBATE_TYPES.TP);
  // Create a map for quick access to debater details by name.
  const debaterMap = new Map(debaters.map(d => [d.name, d]));

  tpDebaters.forEach(debater => {
    if (!debater.partner) {
      return; // Allowed if no partner is listed (intends to be Ironman or roster is incomplete).
    }

    const partner = debaterMap.get(debater.partner);

    // Partner existence check.
    if (!partner) {
      errors.push(`Debater "${debater.name}" lists partner "${debater.partner}", who is not in the Debaters roster.`);
      return; // Stop further checks for this debater if partner doesn't exist.
    }

    // Partner type check (Must both be TP).
    if (partner.type !== CONFIG.DEBATE_TYPES.TP) {
      errors.push(`Debater "${debater.name}" (TP) lists partner "${debater.partner}" (${partner.type}). Partners must both be TP.`);
    }

    // Reciprocal partnership check.
    if (partner.partner !== debater.name) {
      errors.push(`Partnership mismatch: "${debater.name}" lists "${debater.partner}", but "${partner.name}" lists "${partner.partner || 'nobody'}".`);
    }

    // Hard Mode consistency check.
    if (debater.hardMode !== partner.hardMode) {
      errors.push(`Hard Mode mismatch: "${debater.name}" and "${debater.partner}" must have the same Hard Mode setting.`);
    }
  });

  // 4. Room Uniqueness (PRD 5.5).
  const roomNameSet = new Set();
  rooms.forEach(room => {
    if (roomNameSet.has(room.name)) {
      errors.push(`Duplicate Room Name found: "${room.name}".`);
    }
    roomNameSet.add(room.name);
  });

  // Display errors if any exist.
  if (errors.length > 0) {
    const errorMessage = "Roster validation failed. Please correct the following issues (also check conditional formatting highlights):\n\n" + errors.join('\n');
    if (ui) {
      ui.alert('Validation Error', errorMessage, ui.ButtonSet.OK);
    }
    return false; // Halt execution.
  }

  return true; // Proceed with pairing.
}

// =============================================================================
// Core Pairing Logic
// =============================================================================

/**
 * Executes the pairing logic for a single round of debates.
 * Uses a randomized iterative approach (Monte Carlo method / randomized hill climbing) to optimize based on soft constraints.
 * This approach balances optimization quality with execution speed, suitable for Google Apps Script limitations.
 * @param {Array<object>} teams - The teams/debaters participating in this round.
 * @param {Array<object>} judges - Available judges for this round.
 * @param {Array<object>} rooms - Available rooms for this round.
 * @param {object} history - The current match history (potentially updated in-memory for LD R2).
 * @param {number} roundNum - The round number (1 or 2).
 * @param {string|null} [excludedFromBye=null] - Name of the team excluded from getting a BYE (used for LD Round 2 fairness).
 * @returns {Array<object>} The generated pairings for the round.
 */
function executePairingRound(teams, judges, rooms, history, roundNum, excludedFromBye = null) {
  // 1. Handle BYE assignment (PRD 5.4 Special Cases).
  let availableTeams = [...teams]; // Create a working copy of the teams array.
  let pairings = [];
  let byeTeam = null;

  // Check if there is an odd number of teams.
  if (availableTeams.length % 2 !== 0) {
    // Sort by fewest historical BYEs to ensure fairness (PRD 5.4).
    availableTeams.sort((a, b) => a.history.byes - b.history.byes);

    // Find the best candidate (fewest BYEs) who is not excluded (for LD R2 fairness, PRD 5.4.1).
    let byeIndex = availableTeams.findIndex(team => team.name !== excludedFromBye);

    if (byeIndex === -1) {
      // If the only participant(s) left were excluded (e.g., only 1 participant total in the event),
      // they must take the bye regardless of the exclusion rule. This ensures the system always generates a result (PRD 4.2).
      byeIndex = 0;
    }

    if (availableTeams.length > 0) {
      // Remove the selected team from the pool and assign them the BYE.
      byeTeam = availableTeams.splice(byeIndex, 1)[0];

      pairings.push({
        Round: roundNum,
        Aff: byeTeam, // The team receiving the BYE is conventionally listed as Aff.
        Neg: { name: CONFIG.ROLES.BYE, members: [] },
        Judges: [], // Initialize Judges array (BYEs have no judges).
        Room: null,
        Penalty: 0,
        isBye: true
      });
    }
  }

  // 2. Generate optimized pairings (PRD 5.4 Soft Constraints).

  const NUM_ITERATIONS = 500; // Number of attempts for the optimization algorithm. Higher values increase quality but take longer.
  let bestPairingSet = null;
  let lowestPenaltyScore = Infinity;

  // If no teams remain (0 or 1 team total), return the pairings (which might only contain the BYE).
  if (availableTeams.length === 0) {
    return pairings;
  }

  // Optimization Loop: Try different randomized pairings and keep the best one found.
  for (let i = 0; i < NUM_ITERATIONS; i++) {
    // Shuffle the remaining teams to create a varied initial pairing configuration for this iteration.
    const shuffledTeams = shuffleArray([...availableTeams]);
    const currentPairings = [];
    let currentTotalPenalty = 0;

    // Create pairings from the shuffled list.
    for (let j = 0; j < shuffledTeams.length; j += 2) {
      const team1 = shuffledTeams[j];
      const team2 = shuffledTeams[j + 1];

      // Decide Aff/Neg assignment to balance historical roles.
      // Calculate Aff Advantage: (NegCount - AffCount). A higher value means they have done Neg more often and should prefer Aff.
      const team1AffAdvantage = team1.history.negCount - team1.history.affCount;
      const team2AffAdvantage = team2.history.negCount - team2.history.affCount;

      let aff, neg;
      if (team1AffAdvantage > team2AffAdvantage) {
        // Team 1 strongly prefers Aff.
        aff = team1; neg = team2;
      } else if (team2AffAdvantage > team1AffAdvantage) {
        // Team 2 strongly prefers Aff.
        aff = team2; neg = team1;
      } else {
        // Equal history, randomly assign roles.
        [aff, neg] = (Math.random() > 0.5) ? [team1, team2] : [team2, team1];
      }

      // Calculate the penalty score for this specific pairing based on soft constraints.
      const penalty = calculatePairingPenalty(aff, neg, history);
      currentTotalPenalty += penalty;

      currentPairings.push({
        Round: roundNum,
        Aff: aff,
        Neg: neg,
        Judges: [], // Initialize Judges array, to be assigned later.
        Room: null,
        Penalty: penalty,
        isBye: false
      });
    }

    // Check if this iteration produced a better result (lower total penalty).
    if (currentTotalPenalty < lowestPenaltyScore) {
      lowestPenaltyScore = currentTotalPenalty;
      bestPairingSet = currentPairings;
    }

    // Optimization: If score is 0 (perfect pairing with no soft constraint violations), stop early.
    if (lowestPenaltyScore === 0) break;
  }

  // Add the best found pairings to the overall list.
  if (bestPairingSet) {
    pairings.push(...bestPairingSet);
  }

  // 3. Assign Judges (including panels if surplus exists) and Rooms.
  // This function mutates the 'pairings' array by assigning resources and updating penalties.
  assignResources(pairings, judges, rooms, history);

  return pairings;
}

/**
 * Calculates the penalty score for a specific pairing based on soft constraints (PRD 5.4).
 * This score guides the optimization algorithm.
 * @param {object} team1 - The first team (Aff).
 * @param {object} team2 - The second team (Neg).
 * @param {object} history - Match history (potentially in-memory).
 * @returns {number} The penalty score (lower is better).
 */
function calculatePairingPenalty(team1, team2, history) {
  let penalty = 0;

  // Constraint 1: Hard Mode mismatch (High Penalty). This is the most important soft constraint.
  if (team1.hardMode !== team2.hardMode) {
    penalty += 100;
  }

  // Constraint 2: Minimize rematches (Medium Penalty).
  // We check if the individuals have faced the opposing TEAM name before.
  const checkRematches = (team, opponentTeamName) => {
    let rematches = 0;
    // We iterate over members. Although if teams are permanent (TP), checking one member might suffice, 
    // iterating ensures robustness if history somehow differs between partners (e.g. previous Ironman history) or for LD (where team=member).
    team.members.forEach(member => {
      // Check the history passed in memory, which might be updated from Round 1 (for LD).
      const globalMemberHistory = (history.debaters && history.debaters[member]) || { opponents: {} };
      // History correctly tracks the opponent team name.
      if (globalMemberHistory.opponents[opponentTeamName]) {
        rematches += globalMemberHistory.opponents[opponentTeamName];
      }
    });
    return rematches;
  };

  // We calculate the total number of previous encounters between the teams.
  // Since team names are deterministic (alphabetical sort), this comparison is reliable.
  // Penalty is weighted by 15 per previous encounter.
  penalty += (checkRematches(team1, team2.name) + checkRematches(team2, team1.name)) * 15;

  return penalty;
}

/**
 * Helper function to calculate judge penalty, used in assignResources.
 * Determines how suitable a judge is for a specific match.
 * @param {object} judge - The judge object.
 * @param {Array<string>} participants - List of all debater names in the match.
 * @param {object} history - Match history (potentially in-memory).
 * @returns {number} The penalty score (Infinity if a hard constraint conflict exists).
 */
function calculateJudgePenalty(judge, participants, history) {
  // Hard Constraint 3: Parent-Child conflict.
  const conflict = judge.children.some(child => participants.includes(child));
  if (conflict) {
    return Infinity; // Judge cannot be assigned if conflict exists.
  }

  // Soft Constraint 3: Minimize re-judging.
  let judgePenalty = 0;
  participants.forEach(member => {
    // Check the history passed in memory, which might be updated from Round 1 (for LD).
    const memberHistory = (history.debaters && history.debaters[member]) || { judges: {} };
    if (memberHistory.judges[judge.name]) {
      // Penalty increases exponentially (squared) with repeat judging.
      // This strongly encourages variety. (e.g., 1st repeat=10, 2nd repeat=40, 3rd repeat=90).
      judgePenalty += 10 * Math.pow(memberHistory.judges[judge.name], 2);
    }
  });
  return judgePenalty;
};


/**
 * Assigns judges and rooms to the generated pairings, respecting constraints and utilizing surplus judges for panels.
 * Throws an error if hard constraints cannot be met (PRD 5.4).
 * This function uses a multi-pass greedy algorithm to ensure coverage and optimize placement.
 * @param {Array<object>} pairings - The list of pairings (mutated by this function).
 * @param {Array<object>} availableJudges - Available judges for the round.
 * @param {Array<object>} availableRooms - Available rooms for the round.
 * @param {object} history - Match history (potentially in-memory).
 */
function assignResources(pairings, availableJudges, availableRooms, history) {
  // Create consumable pools. Shuffled for randomness when scores/preferences are tied.
  const judgesPool = shuffleArray([...availableJudges]);
  const roomsPool = shuffleArray([...availableRooms]);
  const assignedJudges = new Set();
  const assignedRooms = new Set();

  // Filter out BYEs (which don't need resources) and sort matches by penalty descending.
  // Sorting prioritizes assigning resources to the most difficult matchups (highest penalty) first, increasing the chance of finding suitable judges/rooms.
  const matchesToAssign = pairings.filter(p => !p.isBye).sort((a, b) => b.Penalty - a.Penalty);

  // Check Hard Constraint 1: Sufficiency (At least one judge and one room per match).
  if (matchesToAssign.length > judgesPool.length) {
    throw new Error(`Insufficient judges. Required (minimum): ${matchesToAssign.length}, Available: ${judgesPool.length}.`);
  }
  if (matchesToAssign.length > roomsPool.length) {
    throw new Error(`Insufficient rooms. Required: ${matchesToAssign.length}, Available: ${roomsPool.length}.`);
  }

  // --- Pass 1: Assign Primary Judges (Ensure coverage) ---
  // Goal: Assign one suitable judge to every match, prioritizing the best fit (lowest penalty).
  for (const pairing of matchesToAssign) {
    const participants = [...pairing.Aff.members, ...pairing.Neg.members];

    let bestJudge = null;
    let lowestJudgePenalty = Infinity;

    // Iterate through the available judges pool to find the best fit.
    for (const judge of judgesPool) {
      // Check if the judge is already assigned in this round (Hard Constraint 2).
      if (assignedJudges.has(judge.name)) continue;

      const penalty = calculateJudgePenalty(judge, participants, history);

      if (penalty < lowestJudgePenalty) {
        lowestJudgePenalty = penalty;
        bestJudge = judge;
      }
    }

    // Assign the best judge found.
    if (bestJudge && lowestJudgePenalty !== Infinity) {
      pairing.Judges.push(bestJudge);
      assignedJudges.add(bestJudge.name);
      pairing.Penalty += lowestJudgePenalty; // Add judge penalty to the total match penalty score.
    } else {
      // CRITICAL FAILURE: A match cannot be judged due to conflicts (e.g., all available judges are parents of the debaters).
      // This violates PRD 5.4 Hard Constraints and requires manual intervention.
      throw new Error(`Could not find a conflict-free primary judge for match: ${pairing.Aff.name} vs ${pairing.Neg.name}. Please review judge availability and conflicts.`);
    }
  }

  // --- Pass 2: Assign Surplus Judges (Paneling) (PRD 5.4 Special Cases) ---
  // Goal: Assign remaining judges to panels, prioritizing good fit and even distribution.
  const surplusJudges = judgesPool.filter(j => !assignedJudges.has(j.name));

  if (surplusJudges.length > 0 && matchesToAssign.length > 0) {
    // Iterate through the surplus judges and assign them one by one to the best possible match.
    for (const judge of surplusJudges) {
      let bestMatch = null;
      let lowestTotalPenalty = Infinity;
      let bestActualJudgePenalty = 0; // Tracks the penalty contribution of this specific judge.

      // Find the best match for this surplus judge.
      for (const pairing of matchesToAssign) {
        const participants = [...pairing.Aff.members, ...pairing.Neg.members];

        const penalty = calculateJudgePenalty(judge, participants, history);

        // Encourage even distribution by adding a penalty based on the current panel size.
        // This ensures we prioritize adding a 2nd judge to all matches before adding a 3rd judge to any match.
        const distributionPenalty = pairing.Judges.length * 50;
        const totalPenalty = penalty + distributionPenalty;

        if (totalPenalty < lowestTotalPenalty) {
          lowestTotalPenalty = totalPenalty;
          bestMatch = pairing;
          bestActualJudgePenalty = penalty;
        }
      }

      // Assign the judge to the best match found, if no conflicts exist.
      if (bestMatch && lowestTotalPenalty !== Infinity) {
        bestMatch.Judges.push(judge);
        // Update the match penalty with the actual penalty cost of this judge (excluding the distribution penalty).
        bestMatch.Penalty += bestActualJudgePenalty;
      }
      // If a surplus judge cannot be placed anywhere (due to conflicts with all matches), they remain unassigned. This is acceptable.
    }
  }


  // --- Pass 3: Assign Rooms ---
  // Goal: Assign rooms, respecting Soft Constraint 4 (Judge preference).
  for (const pairing of matchesToAssign) {
    // We assume the first judge in the array (assigned in Pass 1) is the primary/chair judge for room preference.
    // Check if primaryJudge exists (it should always exist for non-BYE matches after Pass 1).
    const primaryJudge = pairing.Judges.length > 0 ? pairing.Judges[0] : null;

    let bestRoom = null;
    let highestRoomPreferenceScore = -1; // We want to maximize preference score here (higher is better).

    // Iterate through available rooms to find the best fit based on judge history.
    for (const room of roomsPool) {
      if (assignedRooms.has(room.name)) continue;

      // Soft Constraint 4: Preferentially assign judges to rooms they've used before.
      // Score is the number of times the primary judge has used this room.
      const preferenceScore = (primaryJudge && primaryJudge.roomHistory[room.name]) || 0;

      if (preferenceScore > highestRoomPreferenceScore) {
        highestRoomPreferenceScore = preferenceScore;
        bestRoom = room;
      }
    }

    // Assign the best room found.
    if (bestRoom) {
      pairing.Room = bestRoom;
      assignedRooms.add(bestRoom.name);
    }
    // Note: We verified sufficient rooms at the start (Hard Constraint 1), so bestRoom should always be found.
  }
}

/**
 * Updates the in-memory history object with the results of a round.
 * Used specifically for LD sequential pairing (PRD 5.4.1) to simulate Round 1 results before generating Round 2.
 * Handles updates correctly when judge panels are used.
 * @param {object} history - The original history object (from getMatchHistory).
 * @param {Array<object>} pairings - The pairings from the round just generated.
 * @returns {object} A new history object updated with the round results.
 */
function updateHistoryInMemory(history, pairings) {
  // Deep copy the history object to avoid mutating the original history object.
  // This is crucial as the original history object represents the persistent state.
  // Note: This simple JSON deep copy works for this specific data structure (no Dates, functions, etc.).
  const newHistory = JSON.parse(JSON.stringify(history));
  if (!newHistory.debaters) newHistory.debaters = {};

  // Helper function to update stats for a team (and its members) in a match.
  const updateMatchStats = (team, opponentTeamName, role, judgesList) => {
    team.members.forEach(memberName => {
      // Initialize history structure if the member is missing (e.g., new debater added recently).
      if (!newHistory.debaters[memberName]) {
        newHistory.debaters[memberName] = { byes: 0, opponents: {}, judges: {}, affCount: 0, negCount: 0 };
      }
      const stats = newHistory.debaters[memberName];

      if (role === CONFIG.ROLES.BYE) {
        stats.byes++;
      } else {
        // Update Match counts (Aff/Neg).
        if (role === CONFIG.ROLES.AFF) stats.affCount++;
        if (role === CONFIG.ROLES.NEG) stats.negCount++;

        // Update Opponent Team counts. This ensures R2 minimizes R1 rematches.
        if (opponentTeamName && opponentTeamName !== CONFIG.ROLES.BYE) {
          stats.opponents[opponentTeamName] = (stats.opponents[opponentTeamName] || 0) + 1;
        }

        // Update Judge counts (for every judge in the panel). This ensures R2 minimizes R1 re-judging.
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

  // Process all pairings from the round.
  pairings.forEach(pairing => {
    if (pairing.isBye) {
      // Aff got the BYE.
      updateMatchStats(pairing.Aff, CONFIG.ROLES.BYE, CONFIG.ROLES.BYE, null);
    } else {
      // Standard match. Update both Aff and Neg teams.
      updateMatchStats(pairing.Aff, pairing.Neg.name, CONFIG.ROLES.AFF, pairing.Judges);
      updateMatchStats(pairing.Neg, pairing.Aff.name, CONFIG.ROLES.NEG, pairing.Judges);
    }
  });

  // Note: Judge room history (history.judges) is not updated in memory, as room optimization relies on long-term historical data, not intra-day data.

  return newHistory;
}


// =============================================================================
// Output and History Functions
// =============================================================================

/**
 * Creates the new match sheet and writes the pairings (PRD 5.3 Step 5).
 * Formats the data according to the specific debate type schema.
 * @param {string} sheetName - The name for the new sheet (e.g., "TP 2025-07-29").
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {Array<object>} pairings - The final optimized pairings.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The newly created sheet.
 */
function createMatchSheet(sheetName, debateType, pairings) {
  const headers = (debateType === CONFIG.DEBATE_TYPES.TP) ? getTpMatchHeaders() : getLdMatchHeaders();
  // Use createTab to ensure the sheet exists and has correct headers.
  const sheet = createTab(sheetName, headers);

  // Map the pairing objects into a 2D array suitable for spreadsheet output.
  const outputData = pairings.map(p => {
    // Handle multiple judges (panels) by joining names with a comma, sorted alphabetically for consistency (PRD 5.2 Note).
    const judgeNames = (p.Judges && p.Judges.length > 0) ? p.Judges.map(j => j.name).sort().join(', ') : '';
    const roomName = p.Room ? p.Room.name : '';

    if (debateType === CONFIG.DEBATE_TYPES.TP) {
      // TP Schema: Aff Team | Neg Team | Judge(s) | Room
      return [p.Aff.name, p.Neg.name, judgeNames, roomName];
    } else {
      // LD Schema: Round | Aff Debater | Neg Debater | Judge(s) | Room
      return [p.Round, p.Aff.name, p.Neg.name, judgeNames, roomName];
    }
  });

  // Write the data to the sheet in a single batch operation.
  if (outputData.length > 0) {
    sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);

    // Sort the data for usability (PRD 5.5).
    const dataRange = sheet.getRange(2, 1, outputData.length, headers.length);
    if (debateType === CONFIG.DEBATE_TYPES.LD) {
      // Sort LD by Round (Col 1 asc), then Room (Col 5 asc).
      dataRange.sort([{ column: 1, ascending: true }, { column: 5, ascending: true }]);
    } else {
      // Sort TP by Room (Col 4 asc).
      dataRange.sort({ column: 4, ascending: true });
    }
  }

  return sheet;
}

/**
 * Applies conditional formatting validation rules to a newly generated match sheet (PRD 5.5).
 * This is critical for supporting manual adjustments (PRD 4.2), providing immediate feedback if changes introduce conflicts.
 * This validation handles the complexities of panel judging (comma-separated lists).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The match sheet.
 * @param {string} debateType - 'TP' or 'LD'.
 */
function applyMatchSheetValidation(sheet, debateType) {
  const rules = [];
  // Ensure the sheet has data before applying validation.
  if (sheet.getMaxRows() < 2) return;

  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn());

  // Helper function for INDIRECT references, required as CF cannot reference other sheets directly.
  const rosterRef = (name) => `INDIRECT("${CONFIG.SHEET_NAMES[name]}!A:A")`;

  // Define column letters based on debate type schema.
  // TP Columns: A=Aff, B=Neg, C=Judge(s), D=Room
  // LD Columns: A=Round, B=Aff, C=Neg, D=Judge(s), E=Room

  const judgeCol = debateType === 'TP' ? 'C' : 'D';
  const roomCol = debateType === 'TP' ? 'D' : 'E';

  // --- Referential Integrity (PRD 5.5) ---

  // Rule: Judge(s) must exist in the Judges roster.
  // Challenge: The Judge(s) cell contains a comma-separated list (panels). We must check if ANY name in the list is missing.
  // Solution: Use complex array formulas supported in CF.
  // Breakdown of judgeRosterFormula:
  // 1. SPLIT($${judgeCol}2, ",") -> Splits the cell content into an array of names.
  // 2. TRIM(...) -> Removes leading/trailing spaces from each name.
  // 3. MATCH(..., ${rosterRef('JUDGES')}, 0) -> Tries to find each name in the Judges roster. Returns #N/A if not found.
  // 4. ISNA(...) -> Returns TRUE if MATCH failed (name not found).
  // 5. --(...) -> Converts TRUE/FALSE to 1/0.
  // 6. SUMPRODUCT(...) -> Sums the results. If the sum > 0, at least one judge is missing.
  const judgeRosterFormula = `=AND(LEN(TRIM($${judgeCol}2))>0, IFERROR(SUMPRODUCT(--(ISNA(MATCH(TRIM(SPLIT($${judgeCol}2, ",")), ${rosterRef('JUDGES')}, 0)))), 1)>0)`;

  const judgeExistsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(judgeRosterFormula)
    .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT) // Critical error if roster reference fails.
    .setRanges([dataRange])
    .build();
  rules.push(judgeExistsRule);

  // Rule: Room must exist in the Rooms roster.
  const roomExistsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${roomCol}2<>"", ISNA(MATCH($${roomCol}2, ${rosterRef('ROOMS')}, 0)))`)
    .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
    .setRanges([dataRange])
    .build();
  rules.push(roomExistsRule);

  // --- Uniqueness Constraints (PRD 5.5) ---
  // These constraints check for resource duplication (e.g., judge assigned twice).

  if (debateType === 'TP') {
    // TP Rule: Each judge is assigned to only one match.
    // Challenge: We need to check individual judge duplication across all panels in the sheet.
    // Breakdown of uniqueJudgeFormula:
    // 1. TEXTJOIN(",", TRUE, $C$2:$C) -> Combines all judge lists into one giant comma-separated string.
    // 2. SPLIT(..., ",") -> Splits the giant string into an array of all individual judge assignments.
    // 3. TRIM(SPLIT($C2, ",")) -> Gets the list of judges in the current row.
    // 4. COUNTIF(AllAssignmentsArray, CurrentRowJudges) -> Counts how many times each judge in the current row appears in the total assignments.
    // 5. --(...)>1 -> Checks if the count is greater than 1 (duplicate).
    // 6. SUMPRODUCT(...) -> Sums the duplicates. If > 0, there is a conflict.
    const uniqueJudgeFormula = `=AND(LEN(TRIM($C2))>0, IFERROR(SUMPRODUCT(--(COUNTIF(TRIM(SPLIT(TEXTJOIN(",", TRUE, $C$2:$C), ",")), TRIM(SPLIT($C2, ",")))>1)), 0)>0)`;

    const uniqueJudgeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueJudgeFormula)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT) // Warning color as this might be manually overridden temporarily.
      .setRanges([dataRange])
      .build();
    rules.push(uniqueJudgeRule);

    // TP Rule: Rooms must be unique.
    const uniqueRoomRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($D2<>"", COUNTIF($D$2:$D, $D2)>1)`)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueRoomRule);

  } else if (debateType === 'LD') {
    // LD Rule: Resources must be unique WITHIN THE ROUND (PRD 5.5).

    // LD Rule: Unique Judge per round (checks individual judges across panels within the round).
    // This uses the same logic as the TP unique judge formula, but adds a FILTER to restrict the TEXTJOIN source to the current round ($A$2:$A=$A2).
    const uniqueJudgeLDFormula = `=AND(LEN(TRIM($D2))>0, IFERROR(SUMPRODUCT(--(COUNTIF(TRIM(SPLIT(TEXTJOIN(",", TRUE, FILTER($D$2:$D, $A$2:$A=$A2)), ",")), TRIM(SPLIT($D2, ",")))>1)), 0)>0)`;

    const uniqueJudgeLDRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueJudgeLDFormula)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueJudgeLDRule);

    // LD Rule: Unique Room per round. Uses COUNTIFS to check for duplicates matching both Round (A) and Room (E).
    const uniqueRoomLDRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($E2<>"", COUNTIFS($A$2:$A, $A2, $E$2:$E, $E2)>1)`)
      .setBackground(CONFIG.STYLES.WARNING_HIGHLIGHT)
      .setRanges([dataRange])
      .build();
    rules.push(uniqueRoomLDRule);


    // LD Rule: Each debater is in exactly one match PER ROUND. (PRD 5.5 Critical Logic Failure check).

    // Check if Debater in B (Aff) appears elsewhere in B or C within the same round (A).
    // COUNTIFS($B$2:$B, $B2, $A$2:$A, $A2)>1 -> Checks for duplicate Aff assignment in the round.
    // COUNTIFS($C$2:$C, $B2, $A$2:$A, $A2)>0 -> Checks if the Aff debater is assigned as Neg in the same round.
    const uniqueDebaterAffFormula = `=AND($B2<>"${CONFIG.ROLES.BYE}", $B2<>"", OR(COUNTIFS($B$2:$B, $B2, $A$2:$A, $A2)>1, COUNTIFS($C$2:$C, $B2, $A$2:$A, $A2)>0))`;
    const uniqueDebaterAffRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueDebaterAffFormula)
      .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT) // Critical error if a debater is double-booked.
      .setRanges([sheet.getRange(`B2:B${sheet.getMaxRows()}`)]) // Apply only to Col B (Aff).
      .build();
    rules.push(uniqueDebaterAffRule);

    // Similar check for the Neg column (C).
    const uniqueDebaterNegFormula = `=AND($C2<>"${CONFIG.ROLES.BYE}", $C2<>"", OR(COUNTIFS($C$2:$C, $C2, $A$2:$A, $A2)>1, COUNTIFS($B$2:$B, $C2, $A$2:$A, $A2)>0))`;
    const uniqueDebaterNegRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(uniqueDebaterNegFormula)
      .setBackground(CONFIG.STYLES.ERROR_HIGHLIGHT)
      .setRanges([sheet.getRange(`C2:C${sheet.getMaxRows()}`)]) // Apply only to Col C (Neg).
      .build();
    rules.push(uniqueDebaterNegRule);
  }

  // Apply all collected rules to the sheet.
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}


/**
 * Appends the generated match results to the AGGREGATE_HISTORY sheet (PRD 5.3 Step 6).
 * This allows the Match Summary formulas to update automatically (PRD 5.2).
 * Handles panel judging by recording separate entries for each judge (denormalization).
 * @param {Array<object>} pairings - The final pairings generated.
 * @param {string} debateType - 'TP' or 'LD'.
 * @param {string} dateString - The date of the matches (YYYY-MM-DD).
 */
function updateAggregateHistory(pairings, debateType, dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY);

  // If the history sheet somehow doesn't exist (e.g., accidentally deleted), we must recreate it to prevent failure.
  if (!historySheet) {
    Logger.log("WARNING: AGGREGATE_HISTORY sheet missing. Recreating.");
    const newHistorySheet = createTab(CONFIG.SHEET_NAMES.AGGREGATE_HISTORY, getAggregateHistoryHeaders());
    newHistorySheet.hideSheet();
    return updateAggregateHistory(pairings, debateType, dateString); // Retry the update operation.
  }

  const historyData = []; // Array to accumulate data for bulk writing.

  // History Schema Headers: Date, Type, Round, Debater Name, Is Ironman, Role, Opponent Team, Judge, Room

  // Helper function to append history entries for a team.
  // Modified to handle multiple judges (paneling) and ensure Opponent Team name is recorded correctly.
  const appendHistory = (team, role, opponentTeamName, pairing) => {
    const roomName = pairing.Room ? pairing.Room.name : '';
    const judges = pairing.Judges || [];

    // Iterate through each member of the team (1 for LD/Ironman, 2 for TP).
    team.members.forEach(memberName => {
      // If there are judges (standard match), record one entry PER JUDGE.
      // This denormalization is crucial for accurate judge statistics (PRD 5.2 Note).
      if (judges.length > 0) {
        judges.forEach(judge => {
          historyData.push([
            dateString,
            debateType,
            pairing.Round,
            memberName,
            team.isIronman,
            role,
            opponentTeamName, // The name of the opposing team/debater.
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
          opponentTeamName, // BYE or similar.
          '', // Judge
          roomName
        ]);
      }
    });
  };

  // Process all pairings and generate the history entries.
  pairings.forEach(pairing => {
    if (pairing.isBye) {
      // Aff got the BYE.
      appendHistory(pairing.Aff, CONFIG.ROLES.BYE, CONFIG.ROLES.BYE, pairing);
    } else {
      // Standard match. Pass the name of the opposing team for accurate tracking.
      appendHistory(pairing.Aff, CONFIG.ROLES.AFF, pairing.Neg.name, pairing);
      appendHistory(pairing.Neg, CONFIG.ROLES.NEG, pairing.Aff.name, pairing);
    }
  });

  // Write the collected history data to the sheet in a single batch operation.
  if (historyData.length > 0) {
    const nextRow = historySheet.getLastRow() + 1;
    historySheet.getRange(nextRow, 1, historyData.length, historyData[0].length).setValues(historyData);
  }
}

// =============================================================================
// Utility Functions
// =============================================================================

/**
 * Sorts the sheets (tabs) according to the requirements (PRD 5.1).
 * Order: Permanent tabs first (fixed order), then Generated tabs (newest first, LD before TP), then Other tabs, then Hidden tabs last.
 */
function sortSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sheets = ss.getSheets();

  // Helper function to determine the priority and categorization of a sheet based on its name.
  const getSheetPriority = (sheetName) => {
    // 1. Permanent Tabs
    const permanentIndex = CONFIG.PERMANENT_TABS_ORDER.indexOf(sheetName);
    if (permanentIndex !== -1) {
      return { type: 'Permanent', order: permanentIndex, date: null };
    }

    // 2. Hidden/Aggregate Tabs (always last).
    if (sheetName === CONFIG.SHEET_NAMES.AGGREGATE_HISTORY) {
      return { type: 'Hidden', order: 9999, date: null };
    }

    // 3. Generated Match Tabs (e.g., "TP 2025-07-29").
    // Uses regex to identify the format and extract the type and date.
    const match = sheetName.match(/^(TP|LD) (\d{4}-\d{2}-\d{2})$/);
    if (match) {
      const debateType = match[1];
      const date = match[2];
      // LD (1) appears before TP (2) (PRD 5.1).
      const typeOrder = debateType === 'LD' ? 1 : 2;
      return { type: 'Generated', order: typeOrder, date: date };
    }

    // 4. Other tabs (e.g., manually created by the user, or legacy sheets).
    return { type: 'Other', order: 9998, date: null };
  };

  // Sort the array of sheets based on the defined priorities.
  const sortedSheets = sheets.sort((a, b) => {
    const prioA = getSheetPriority(a.getName());
    const prioB = getSheetPriority(b.getName());

    // Primary Sort: By Type category.
    if (prioA.type !== prioB.type) {
      const typeOrder = ['Permanent', 'Generated', 'Other', 'Hidden'];
      return typeOrder.indexOf(prioA.type) - typeOrder.indexOf(prioB.type);
    }

    // Secondary Sort: Specific rules for 'Generated' tabs.
    if (prioA.type === 'Generated') {
      // Sort by Date (Newest first - descending).
      if (prioA.date !== prioB.date) {
        return prioB.date.localeCompare(prioA.date);
      }
      // Sort by Type (LD before TP - ascending).
      return prioA.order - prioB.order;
    }

    // Secondary Sort: For Permanent/Other/Hidden, use the predefined order index.
    return prioA.order - prioB.order;
  });

  // Move sheets into the sorted order. Google Apps Script requires moving them one by one.
  // We need to activate the sheet before moving it.
  const activeSheet = ss.getActiveSheet(); // Remember the currently active sheet to restore it later.
  for (let i = 0; i < sortedSheets.length; i++) {
    // Use try-catch because hidden sheets cannot be activated, which prevents moving them if we don't handle the error.
    try {
      ss.setActiveSheet(sortedSheets[i]);
      // moveActiveSheet uses 1-based indexing.
      ss.moveActiveSheet(i + 1);
    } catch (e) {
      // Log the issue but continue sorting the rest of the sheets.
      Logger.log(`Could not move sheet ${sortedSheets[i].getName()} (likely hidden): ${e.message}`);
    }
  }
  // Try restoring the original active sheet.
  try {
    // Check if the active sheet reference is valid and the sheet is not hidden.
    if (activeSheet && !activeSheet.isSheetHidden()) {
      ss.setActiveSheet(activeSheet);
    }
  } catch (e) {
    // Ignore if the original active sheet cannot be activated (e.g., if it was deleted during initialization).
    Logger.log("Could not restore active sheet after sorting: " + e.message);
  }
}

/**
 * Shuffles an array in place using the Fisher-Yates (aka Knuth) algorithm.
 * Used for randomization in the pairing optimization (Monte Carlo) and resource assignment.
 * @param {Array} array - The array to shuffle.
 * @returns {Array} The shuffled array (the same array instance).
 */
function shuffleArray(array) {
  let currentIndex = array.length, randomIndex;

  // While there remain elements to shuffle.
  while (currentIndex !== 0) {
    // Pick a remaining element.
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex--;

    // And swap it with the current element.
    [array[currentIndex], array[randomIndex]] = [
      array[randomIndex], array[currentIndex]];
  }
  return array;
}
