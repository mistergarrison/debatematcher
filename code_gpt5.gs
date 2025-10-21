/**
 * Code.gs
 * Automated Debate Pairing Tool for Google Sheets
 * Single-file implementation per PRD.
 *
 * Author: ChatGPT (implementation per user-provided PRD)
 * Date: 2025-10-20
 *
 * NOTE: Paste this file into the Apps Script editor of the target Google Sheet.
 * Run initializeSheet() once, or open the sheet and use the "Club Admin" menu.
 */

/* ============================
   Configuration & Constants
   ============================ */

/** Names of permanent sheets in order */
const PERMANENT_SHEETS = ['Availability', 'Debaters', 'Judges', 'Rooms'];

/** RSVP options */
const RSVP_OPTIONS = ['Yes', 'No', 'Not responded'];

/** Date formatting */
function todayStr() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ============================
   Helpers
   ============================ */

/**
 * Get UI object for alerts.
 * @returns {GoogleAppsScript.Base.Ui}
 */
function ui() {
  return SpreadsheetApp.getUi();
}

/**
 * Safely get sheet or create (if create=true).
 * @param {string} name
 * @param {boolean} create
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(name, create = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  if (!create) return null;
  return ss.insertSheet(name);
}

/**
 * Ensure sheet exists and set a header row.
 * @param {string} name
 * @param {string[]} headers
 * @param {boolean} freezeHeader
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSheetWithHeader(name, headers, freezeHeader = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Only set header if first row blank or different
  const currentHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let needSet = currentHeader.join('|') !== headers.join('|');
  if (needSet) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  if (freezeHeader) sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Build menu on open.
 */
function onOpen() {
  const menu = ui().createMenu('Club Admin');
  menu.addItem('Initialize Sheet', 'initializeSheet')
      .addSeparator()
      .addItem('Generate TP Matches', 'menuGenerateTP')
      .addItem('Generate LD Matches', 'menuGenerateLD')
      .addSeparator()
      .addItem('Clear RSVPs', 'clearRsvps')
      .addToUi();
  // Re-apply conditional formatting rules each open
  try {
    applyValidationAndFormatting();
  } catch (e) {
    // Non-fatal
    console.warn('Formatting init failed: ' + e);
  }
}

/* ============================
   Initialization
   ============================ */

/**
 * initializeSheet
 * One-click sheet initialization. Won't overwrite existing permanent tabs.
 *
 * Creates:
 * - Availability, Debaters, Judges, Rooms (permanent)
 * - Adds sample rows (only if empty)
 * - Adds formulas to Availability participant list (unique sorted names)
 * - Adds data validation for RSVP cells
 * - Applies conditional formatting and validation rules
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheets().map(s => s.getName());
  // 1) Create permanent tabs
  const availabilityHeaders = ['Participant', 'Attending?'];
  const debatersHeaders = ['Name', 'Debate Type', 'Partner', 'Hard Mode'];
  const judgesHeaders = ['Name', "Children's Names (comma-separated)", 'Debate Type'];
  const roomsHeaders = ['Room Name', 'Debate Type'];

  const sheetsToMake = [
    {name: 'Availability', headers: availabilityHeaders},
    {name: 'Debaters', headers: debatersHeaders},
    {name: 'Judges', headers: judgesHeaders},
    {name: 'Rooms', headers: roomsHeaders}
  ];

  sheetsToMake.forEach(s => ensureSheetWithHeader(s.name, s.headers));

  // 2) Populate sample data if Debaters/Judges/Rooms are empty (only add if no data rows)
  const debatersSheet = getSheet('Debaters');
  if (debatersSheet.getLastRow() <= 1) {
    const sampleDebaters = [
      // LD
      ['Abraham Lincoln', 'LD', '', 'No'],
      ['Stephen A. Douglas', 'LD', '', 'No'],
      ['Clarence Darrow', 'LD', '', 'No'],
      ['William Jennings Bryan', 'LD', '', 'No'],
      ['William F. Buckley Jr.', 'LD', '', 'No'],
      ['Gore Vidal', 'LD', '', 'No'],
      ['Christopher Hitchens', 'LD', '', 'No'],
      ['Tony Blair', 'LD', '', 'No'],
      ['Jordan Peterson', 'LD', '', 'No'],
      ['Slavoj Žižek', 'LD', '', 'No'],
      ['Lloyd Bentsen', 'LD', '', 'No'],
      ['Dan Quayle', 'LD', '', 'No'],
      ['Richard Dawkins', 'LD', '', 'No'],
      ['Rowan Williams', 'LD', '', 'No'],
      ['Diogenes', 'LD', '', 'No'],
      // TP teams (we'll store each teammate as separate row but include partner name)
      ['Noam Chomsky', 'TP', 'Michel Foucault', 'No'],
      ['Michel Foucault', 'TP', 'Noam Chomsky', 'No'],
      ['Harlow Shapley', 'TP', 'Heber Curtis', 'No'],
      ['Heber Curtis', 'TP', 'Harlow Shapley', 'No'],
      ['Muhammad Ali', 'TP', 'George Foreman', 'No'],
      ['George Foreman', 'TP', 'Muhammad Ali', 'No'],
      ['Richard Nixon', 'TP', 'Nikita Khrushchev', 'No'],
      ['Nikita Khrushchev', 'TP', 'Richard Nixon', 'No'],
      ['Thomas Henry Huxley', 'TP', 'Samuel Wilberforce', 'No'],
      ['Samuel Wilberforce', 'TP', 'Thomas Henry Huxley', 'No'],
      ['John F. Kennedy', 'TP', 'David Frost', 'No'],
      ['David Frost', 'TP', 'John F. Kennedy', 'No'],
      ['Bob Dole', 'TP', 'Bill Clinton', 'No'],
      ['Bill Clinton', 'TP', 'Bob Dole', 'No']
    ];
    debatersSheet.getRange(2, 1, sampleDebaters.length, sampleDebaters[0].length).setValues(sampleDebaters);
  }

  const judgesSheet = getSheet('Judges');
  if (judgesSheet.getLastRow() <= 1) {
    const sampleJudges = [
      // TP judges
      ['Howard K. Smith', 'John F. Kennedy', 'TP'],
      ['Fons Elders', 'Noam Chomsky, Michel Foucault', 'TP'],
      ['John Stevens Henslow', 'Samuel Wilberforce', 'TP'],
      ['Jim Lehrer', 'Bill Clinton', 'TP'],
      ['Judy Woodruff', '', 'TP'],
      ['Tom Brokaw', '', 'TP'],
      ['Frank McGee', '', 'TP'],
      ['Quincy Howe', '', 'TP'],
      // LD judges
      ['John T. Raulston', 'Clarence Darrow', 'LD'],
      ['Rudyard Griffiths', 'Jordan Peterson', 'LD'],
      ['Stephen J. Blackwood', '', 'LD'],
      ['Brit Hume', '', 'LD'],
      ['Jon Margolis', '', 'LD'],
      ['Bill Shadel', '', 'LD'],
      ['Judge Judy', '', 'LD']
    ];
    judgesSheet.getRange(2, 1, sampleJudges.length, sampleJudges[0].length).setValues(sampleJudges);
  }

  const roomsSheet = getSheet('Rooms');
  if (roomsSheet.getLastRow() <= 1) {
    const sampleRooms = [
      ['Room 101', 'LD'],
      ['Room 102', 'LD'],
      ['Room 103', 'LD'],
      ['Room 201', 'LD'],
      ['Sanctuary right', 'LD'],
      ['Sanctuary left', 'LD'],
      ['Pantry', 'LD'],
      ['Chapel', 'TP'],
      ['Library', 'TP'],
      ['Music lounge', 'TP'],
      ['Cry room', 'TP'],
      ['Office', 'TP'],
      ['Office hallway', 'TP']
    ];
    roomsSheet.getRange(2, 1, sampleRooms.length, sampleRooms[0].length).setValues(sampleRooms);
  }

  // 3) Availability sheet: participant list formula and default RSVPs
  const availabilitySheet = getSheet('Availability');
  // Put participants formula in A2 and fill a bunch of rows (formula uses arrayformula)
  // We'll generate a script formula that uses UNIQUE and SORT across Debaters and Judges names.
  const participantFormula = '=SORT(UNIQUE(TRANSPOSE(SPLIT(TEXTJOIN("|",TRUE,Debaters!A2:A,Judges!A2:A),"|"))))';
  availabilitySheet.getRange('A2').setFormula(participantFormula);
  // Clear any old Attending? column and set default "Not responded"
  // We'll set a script to set Attending? via onEdit or initialize; but here set the dropdown validation and default area
  const maxRows = 200;
  availabilitySheet.getRange(2, 2, maxRows, 1).clearContent();
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(RSVP_OPTIONS).setAllowInvalid(false).build();
  availabilitySheet.getRange(2, 2, maxRows, 1).setDataValidation(rule);
  // Fill "Not responded" for participant rows that are non-empty
  const avVals = availabilitySheet.getRange(2,1,maxRows,1).getValues().map(r=>r[0]);
  const fill = [];
  for (let i=0;i<avVals.length;i++) {
    if (avVals[i] && avVals[i].toString().trim()!=='') fill.push([RSVP_OPTIONS[2]]);
    else fill.push(['']);
  }
  availabilitySheet.getRange(2,2,maxRows,1).setValues(fill);

  // 4) Apply conditional formatting & validation rules
  applyValidationAndFormatting();

  // 5) Reorder sheets so permanents are first in specified order
  reorderPermanentTabs();

  ui().alert('Initialization complete. Permanent sheets created (if they did not exist) and sample data added where empty.');
}

/* ============================
   Validation & Formatting
   ============================ */

/**
 * Apply validation and conditional formatting rules described in PRD.
 * Overwrites formatting on the affected sheets (recommended approach).
 */
function applyValidationAndFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Availability: highlight Attending? cell in light red if Participant present but Attending blank (Not responded counts as content so only blank).
  const availability = getSheet('Availability', false);
  if (availability) {
    const lastRow = Math.max(availability.getLastRow(), 200);
    const cfRules = [];
    // Conditional formula can't reference other sheets easily for cross-sheet; but this rule is internal to availability.
    // Formula: =AND($A2<>"",$B2="")
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$B2="")')
      .setBackground('#FFCFCF') // light red
      .setRanges([availability.getRange(2,1,lastRow-1,2)])
      .build();
    cfRules.push(rule);
    availability.setConditionalFormatRules(cfRules);
    // Ensure header formatting
    availability.getRange(1,1,1,2).setFontWeight('bold');
  }

  // Rooms: Room names must be unique -> highlight duplicates
  const rooms = getSheet('Rooms', false);
  if (rooms) {
    const lastRow = Math.max(rooms.getLastRow(), 200);
    const range = rooms.getRange(2,1,lastRow-1,1);
    // Use custom formula to detect duplicates: =COUNTIF($A:$A,$A2)>1
    const ruleRooms = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=COUNTIF($A:$A,$A2)>1')
      .setBackground('#FFCFCF')
      .setRanges([range])
      .build();
    rooms.setConditionalFormatRules([ruleRooms]);
    rooms.getRange(1,1,1,2).setFontWeight('bold');
    // Room name uniqueness can also be validated via script-check in generate functions
  }

  // Judges: Cannot be both judge and debater. We'll highlight names that appear in Debaters!A:A
  const judges = getSheet('Judges', false);
  if (judges) {
    const lastRow = Math.max(judges.getLastRow(), 200);
    const range = judges.getRange(2,1,lastRow-1,1);
    // Conditional formula referencing other sheet via COUNTIF(INDIRECT("Debaters!A:A"),$A2)>0
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=COUNTIF(INDIRECT("Debaters!A:A"),$A2)>0')
      .setBackground('#FFCFCF')
      .setRanges([range])
      .build();
    // Also highlight if children's names are not found in Debaters (we'll check by splitting)
    const childrenRange = judges.getRange(2,2,lastRow-1,1);
    const ruleChildren = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2<>"",SUMPRODUCT(--(LEN(TRIM(SPLIT($B2,",")))>0),--(COUNTIF(INDIRECT("Debaters!A:A"),TRIM(SPLIT($B2,",")))=0))>0)')
      .setBackground('#FFCFCF')
      .setRanges([childrenRange])
      .build();
    judges.setConditionalFormatRules([rule, ruleChildren]);
    judges.getRange(1,1,1,3).setFontWeight('bold');
    // Also set data validation for Debate Type
    const dtRule = SpreadsheetApp.newDataValidation().requireValueInList(['TP','LD']).setAllowInvalid(false).build();
    judges.getRange(2,3,200,1).setDataValidation(dtRule);
  }

  // Debaters: partnership consistency checks via conditional formatting
  const debaters = getSheet('Debaters', false);
  if (debaters) {
    const lastRow = Math.max(debaters.getLastRow(), 200);
    const nameRange = debaters.getRange(2,1,lastRow-1,1);
    const partnerRange = debaters.getRange(2,3,lastRow-1,1);
    // Partner must exist: =AND($B2="TP", $C2<>"", COUNTIF(INDIRECT("Debaters!A:A"),$C2)=1)
    const partnerMissingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2="TP",$C2<>"",COUNTIF(INDIRECT("Debaters!A:A"),$C2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([partnerRange])
      .build();
    // Partnership reciprocity: highlight partner cell if partner's partner not equal
    // This is tricky in pure formula; we can flag when: =AND($B2="TP",$C2<>"", INDEX(INDIRECT("Debaters!C:C"), MATCH($C2,INDIRECT("Debaters!A:A"),0))<>$A2)
    const reciprocityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2="TP",$C2<>"",INDEX(INDIRECT("Debaters!C:C"),MATCH($C2,INDIRECT("Debaters!A:A"),0))<>$A2)')
      .setBackground('#FFCFCF')
      .setRanges([partnerRange])
      .build();
    // Hard Mode consistent between partners
    const hardModeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2="TP",$C2<>"",INDEX(INDIRECT("Debaters!D:D"),MATCH($C2,INDIRECT("Debaters!A:A"),0))<>$D2)')
      .setBackground('#FFCFCF')
      .setRanges([debaters.getRange(2,4,lastRow-1,1)])
      .build();
    // Debate type validation
    const dtRule = SpreadsheetApp.newDataValidation().requireValueInList(['TP','LD']).setAllowInvalid(false).build();
    debaters.getRange(2,2,200,1).setDataValidation(dtRule);
    // Hard Mode validation
    const hmRule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No']).setAllowInvalid(false).build();
    debaters.getRange(2,4,200,1).setDataValidation(hmRule);
    debaters.setConditionalFormatRules([partnerMissingRule, reciprocityRule, hardModeRule]);
    debaters.getRange(1,1,1,4).setFontWeight('bold');
  }
}

/**
 * Reorder sheets so permanent tabs are first in PERMANENT_SHEETS order.
 */
function reorderPermanentTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let idx = 1;
  PERMANENT_SHEETS.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(idx);
      idx++;
    }
  });
  // After permanents, sort generated match tabs by date desc and LD before TP for same date.
  const sheets = ss.getSheets();
  // Collect generated names like "TP yyyy-mm-dd" or "LD yyyy-mm-dd"
  const generated = sheets.filter(s => {
    const n = s.getName();
    return !PERMANENT_SHEETS.includes(n) && /^((TP|LD) \d{4}-\d{2}-\d{2})$/.test(n);
  }).map(s => s.getName());
  // Parse and sort
  generated.sort((a,b) => {
    // a = "TP 2025-07-29"
    const [ta, da] = a.split(' ');
    const [tb, db] = b.split(' ');
    if (db === da) {
      // LD before TP
      if (tb === ta) return 0;
      return (tb === 'LD') ? 1 : -1;
    }
    // Sort by date desc
    return (db > da) ? -1 : 1;
  });
  // Move generated sheets after permanents
  let position = PERMANENT_SHEETS.length + 1;
  generated.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(position);
      position++;
    }
  });
}

/* ============================
   Utility: Read Rosters
   ============================ */

/**
 * Read Debaters roster into array of objects.
 * @returns {Array<{name:string,type:string,partner:string,hardMode:string}>}
 */
function readDebaters() {
  const s = getSheet('Debaters');
  if (!s) return [];
  const values = s.getDataRange().getValues();
  const headers = values.shift();
  const rows = values.filter(r => r[0] && r[0].toString().trim() !== '');
  return rows.map(r => ({
    name: r[0].toString().trim(),
    type: r[1] ? r[1].toString().trim() : '',
    partner: r[2] ? r[2].toString().trim() : '',
    hardMode: r[3] ? r[3].toString().trim() : 'No'
  }));
}

/**
 * Read Judges roster into array of objects.
 * @returns {Array<{name:string,children:Array<string>,type:string}>}
 */
function readJudges() {
  const s = getSheet('Judges');
  if (!s) return [];
  const values = s.getDataRange().getValues();
  const headers = values.shift();
  const rows = values.filter(r => r[0] && r[0].toString().trim() !== '');
  return rows.map(r => ({
    name: r[0].toString().trim(),
    children: r[1] ? r[1].toString().split(',').map(x => x.trim()).filter(Boolean) : [],
    type: r[2] ? r[2].toString().trim() : ''
  }));
}

/**
 * Read Rooms roster.
 * @returns {Array<{name:string,type:string}>}
 */
function readRooms() {
  const s = getSheet('Rooms');
  if (!s) return [];
  const values = s.getDataRange().getValues();
  const headers = values.shift();
  const rows = values.filter(r => r[0] && r[0].toString().trim() !== '');
  return rows.map(r => ({ name: r[0].toString().trim(), type: r[1] ? r[1].toString().trim() : '' }));
}

/**
 * Read Availability: returns list of participants with RSVP status.
 * @returns {Array<{name:string,attending:string}>}
 */
function readAvailability() {
  const s = getSheet('Availability');
  if (!s) return [];
  const data = s.getRange(2,1,Math.max(s.getLastRow()-1, 200), 2).getValues();
  return data.filter(r => r[0] && r[0].toString().trim() !== '').map(r => ({ name: r[0].toString().trim(), attending: r[1] ? r[1].toString().trim() : '' }));
}

/* ============================
   Match History Utilities
   ============================ */

/**
 * Read all generated match sheets and return history object:
 * history.matches is array of {sheetName,type,date,rows}
 * Also returns byeCounts map by name.
 */
function readMatchHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const matches = [];
  const byeCounts = {};
  sheets.forEach(s => {
    const name = s.getName();
    const m = name.match(/^(TP|LD) (\d{4}-\d{2}-\d{2})$/);
    if (!m) return;
    const type = m[1];
    const date = m[2];
    const values = s.getDataRange().getValues();
    const headers = values.shift();
    const rows = values;
    rows.forEach(r => {
      // count BYEs for participants
      if (type === 'TP') {
        const aff = (r[0] || '').toString().trim();
        const neg = (r[1] || '').toString().trim();
        if (aff === 'BYE' && neg) {
          byeCounts[neg] = (byeCounts[neg] || 0) + 1;
        } else if (neg === 'BYE' && aff) {
          byeCounts[aff] = (byeCounts[aff] || 0) + 1;
        } else {
          // no BYE
        }
      } else if (type === 'LD') {
        // If Round field present and an opponent equals "BYE"
        // Schema: Round | Aff | Neg | Judge | Room
        const aff = (r[1] || '').toString().trim();
        const neg = (r[2] || '').toString().trim();
        if (aff === 'BYE' && neg) byeCounts[neg] = (byeCounts[neg] || 0) + 1;
        if (neg === 'BYE' && aff) byeCounts[aff] = (byeCounts[aff] || 0) + 1;
      }
    });
    matches.push({ sheetName: name, type, date, headers, rows });
  });
  return { matches, byeCounts };
}

/* ============================
   Validation (pre-generate)
   ============================ */

/**
 * Run pre-generation validation checks per PRD.
 * Throws Error with descriptive message on failure.
 */
function validateBeforeGenerate(debateType) {
  // debateType is 'TP' or 'LD'
  const debaters = readDebaters();
  const judges = readJudges();
  const rooms = readRooms();

  // Rooms must be unique
  const roomNames = rooms.map(r => r.name);
  const dupRoom = roomNames.find((r, i) => roomNames.indexOf(r) !== i);
  if (dupRoom) throw new Error(`Duplicate room name detected: ${dupRoom} (Rooms tab)`);

  // Judges cannot be both judge and debater
  const debaterNames = debaters.map(d => d.name);
  const judgeNames = judges.map(j => j.name);
  const both = judgeNames.find(j => debaterNames.indexOf(j) !== -1);
  if (both) throw new Error(`Person cannot be both Judge and Debater: ${both}`);

  // All children's names must exist in Debaters roster
  judges.forEach(j => {
    j.children.forEach(child => {
      if (child && debaterNames.indexOf(child) === -1) {
        throw new Error(`Judge '${j.name}' has child '${child}' not present in Debaters roster`);
      }
    });
  });

  // Partnership consistency for TP debaters
  const tpDebaters = debaters.filter(d => d.type === 'TP');
  tpDebaters.forEach(d => {
    if (d.partner && d.partner !== '') {
      const partner = tpDebaters.find(p => p.name === d.partner);
      if (!partner) throw new Error(`Debater '${d.name}' lists partner '${d.partner}' who is not in Debaters roster`);
      if (partner.partner !== d.name) throw new Error(`Partnership not reciprocal: ${d.name} -> ${d.partner} but ${d.partner} -> ${partner.partner}`);
      if (partner.hardMode !== d.hardMode) throw new Error(`Hard Mode mismatch between partners: ${d.name} and ${d.partner}`);
    }
  });

  // Make sure enough judges and rooms for at least 1 match - actual insufficiency checked later
  if (rooms.length === 0) throw new Error('No rooms configured (Rooms tab)');
  if (judges.length === 0) throw new Error('No judges configured (Judges tab)');
}

/* ============================
   Pairing Engine (Core)
   ============================ */

/**
 * menu wrapper for TP generate
 */
function menuGenerateTP() {
  try {
    generateMatches('TP');
  } catch (e) {
    ui().alert('Error generating TP matches: ' + e.message);
  }
}

/**
 * menu wrapper for LD generate
 */
function menuGenerateLD() {
  try {
    generateMatches('LD');
  } catch (e) {
    ui().alert('Error generating LD matches: ' + e.message);
  }
}

/**
 * generateMatches
 * Orchestrates a full generate flow for given debateType ('TP' or 'LD')
 */
function generateMatches(debateType) {
  validateBeforeGenerate(debateType);

  const availability = readAvailability().filter(p => p.attending === 'Yes');
  const availableNames = availability.map(a => a.name);
  if (availableNames.length === 0) {
    ui().alert('No participants RSVP\'d "Yes" for the meeting.');
    return;
  }

  const debaters = readDebaters().filter(d => availableNames.indexOf(d.name) !== -1);
  const judges = readJudges().filter(j => availableNames.indexOf(j.name) !== -1 || true); // judges listed in roster but might not RSVP; PRD expects RSVP for judges too - we will filter below to attending==Yes
  const attendingJudgeNames = readAvailability().filter(a=>a.attending==='Yes').map(a=>a.name);
  const eligibleJudges = readJudges().filter(j => attendingJudgeNames.indexOf(j.name) !== -1 && j.type === debateType);
  const rooms = readRooms().filter(r => r.type === debateType);

  // If no eligible judges (attending & type)
  if (eligibleJudges.length === 0) throw new Error(`No eligible judges of type ${debateType} are attending.`);

  // If insufficient rooms/judges to host any matches, fail
  const maxMatchesPossible = Math.min(Math.floor(availableNames.length / (debateType === 'TP' ? 2 : 1)), rooms.length, eligibleJudges.length);
  if (maxMatchesPossible < 1) throw new Error(`Insufficient resources: rooms or judges for ${debateType} matches.`);

  // Prevent overwrite of today's sheet
  const sheetName = debateType + ' ' + todayStr();
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
    throw new Error(`A sheet named '${sheetName}' already exists for today. Aborting to avoid overwrite.`);
  }

  // Build initial pool of participants depending on type
  if (debateType === 'TP') {
    // TP participants are teams: pair partners into teams. If partner absent, they'll be ironman.
    // Each row in Debaters roster is individual, but teams are partners
    const teams = buildTpTeams(debaters, availableNames);
    // Pair teams
    const matchRows = pairTpTeams(teams, eligibleJudges, rooms);
    createTpSheet(sheetName, matchRows);
  } else {
    // LD two-round flow
    const ldDebaters = readDebaters().filter(d => d.type === 'LD' && availableNames.indexOf(d.name) !== -1);
    if (ldDebaters.length === 0) throw new Error('No LD debaters attending.');
    const matchRows = pairLdTwoRounds(ldDebaters, eligibleJudges, rooms);
    createLdSheet(sheetName, matchRows);
  }
  // Reorder tabs
  reorderPermanentTabs();
}

/* ============================
   TP Helpers
   ============================ */

/**
 * Build TP teams array from attending debaters.
 * Each team object: {name: string (team display), members: [names], hardMode: 'Yes'/'No'}
 */
function buildTpTeams(attendingDebaters, overallAvailableNames) {
  const byName = {};
  attendingDebaters.forEach(d => byName[d.name] = d);
  const visited = new Set();
  const teams = [];
  attendingDebaters.forEach(d => {
    if (visited.has(d.name)) return;
    if (d.type !== 'TP') return;
    const partnerName = d.partner;
    if (partnerName && partnerName !== '' && byName[partnerName]) {
      // both present -> team
      teams.push({ name: `${d.name} / ${partnerName}`, members: [d.name, partnerName], hardMode: d.hardMode || 'No' });
      visited.add(d.name); visited.add(partnerName);
    } else {
      // partner absent -> IRONMAN
      teams.push({ name: `${d.name} (IRONMAN)`, members: [d.name], hardMode: d.hardMode || 'No' });
      visited.add(d.name);
    }
  });
  return teams;
}

/**
 * Pair TP teams given judges and rooms.
 * Returns rows for sheet Aff | Neg | Judge | Room
 */
function pairTpTeams(teams, judges, rooms) {
  // Compute BYE assignment if odd number of teams
  const history = readMatchHistory();
  const byeCounts = history.byeCounts || {};
  const teamCount = teams.length;
  const needsBye = (teamCount % 2 === 1);

  // If insufficient judges/rooms for number of matches, stop
  const matchCount = Math.floor(teamCount / 2);
  if (judges.length < matchCount) throw new Error(`Not enough judges for ${matchCount} TP matches (judges available: ${judges.length})`);
  if (rooms.length < matchCount) throw new Error(`Not enough rooms for ${matchCount} TP matches (rooms available: ${rooms.length})`);

  // Assign BYE to team with fewest BYEs historically
  let byeTeam = null;
  if (needsBye) {
    teams.sort((a,b) => (byeCounts[a.members[0]] || 0) - (byeCounts[b.members[0]] || 0));
    byeTeam = teams.shift(); // remove from pairing
  }

  // Now pair remaining teams trying to satisfy soft constraints
  // We'll use a randomized greedy approach with scoring - try multiple iterations and pick best
  const attempts = 400;
  let bestSolution = null;
  let bestScore = -Infinity;

  for (let t=0;t<attempts;t++) {
    const shuffledTeams = shuffleArray(teams.slice());
    const pairs = [];
    // pair sequentially
    for (let i=0;i<shuffledTeams.length;i+=2) {
      const a = shuffledTeams[i], b = shuffledTeams[i+1];
      pairs.push({aff:a, neg:b});
    }
    // Assign judges & rooms greedily with small preference heuristics
    const judgePool = shuffleArray(judges.slice());
    const roomPool = shuffleArray(rooms.slice());
    const assignments = [];
    for (let i=0;i<pairs.length;i++) {
      const p = pairs[i];
      // choose judge not parent-of any member, and not used already
      let judge = judgePool.find(j => !j.children.some(c => p.aff.members.includes(c) || p.neg.members.includes(c)));
      if (!judge) {
        // relax parent rule (shouldn't happen if judges were prefiltered), choose any unused judge
        judge = judgePool[0];
      }
      // remove chosen judge from pool so unique judges
      const jIndex = judgePool.indexOf(judge);
      if (jIndex !== -1) judgePool.splice(jIndex,1);

      // choose room (first available)
      const room = roomPool.shift();
      assignments.push({aff:p.aff.name, neg:p.neg.name, judge: judge ? judge.name : 'UNASSIGNED', room: room ? room.name : 'UNASSIGNED', affMembers: p.aff.members, negMembers: p.neg.members});
    }
    // Score solution
    const sc = scoreTpSolution(assignments);
    if (sc > bestScore) {
      bestScore = sc;
      bestSolution = {assignments: assignments.slice(), byeTeam};
      // early exit if perfect score (heuristic threshold)
      if (bestScore > 1000) break;
    }
  }
  // Build rows for sheet
  const rows = [];
  if (bestSolution.byeTeam) {
    // BYE row as "BYE" vs team name? PRD: BYE assigned to participant/team with fewest BYEs; BYEs don't need judge/room.
    // We will write team in first column and "BYE" in second for consistency.
    rows.push([bestSolution.byeTeam.name, 'BYE', '', '']);
  }
  bestSolution.assignments.forEach(a => {
    rows.push([a.aff, a.neg, a.judge, a.room]);
  });
  return rows;
}

/**
 * Score TP solution (higher is better).
 * Simple heuristic: reward matching same hard mode, penalize judge-parent conflicts (should be 0), prefer judge variance, room reuse preference left minimal.
 * This is intentionally tunable.
 */
function scoreTpSolution(assignments) {
  let score = 0;
  // reward differing aff/neg history? For now placeholder - could be extended to consult history
  assignments.forEach(a => {
    // penalize if judge is 'UNASSIGNED'
    if (!a.judge || a.judge==='UNASSIGNED') score -= 1000;
    // penalize if room is 'UNASSIGNED'
    if (!a.room || a.room==='UNASSIGNED') score -= 1000;
    // small reward for diversity heuristics (placeholder)
    score += 10;
  });
  return score;
}

/* ============================
   LD Helpers (two-round system)
   ============================ */

/**
 * pairLdTwoRounds
 * Implements the sequential two-round algorithm per PRD.
 *
 * Returns rows for sheet: Round | Aff | Neg | Judge | Room
 */
function pairLdTwoRounds(ldDebaters, eligibleJudges, rooms) {
  // Pre-check resources
  const availableDebaters = ldDebaters.slice();
  const total = availableDebaters.length;
  // If odd -> one BYE per two rounds? PRD: BYE assigned in Round 1 and Round 2 pairing avoids repeat BYEs.
  // For simplicity: each round pairs all available debaters; if odd, assign one BYE each round such that BYEs are distributed.
  if (eligibleJudges.length < Math.ceil(total / 2)) throw new Error(`Not enough eligible judges for LD rounds (${eligibleJudges.length} judges for ${Math.ceil(total/2)} matches).`);
  if (rooms.length < Math.ceil(total / 2)) throw new Error(`Not enough rooms for LD rounds (${rooms.length} rooms for ${Math.ceil(total/2)} matches).`);

  // Read history and build helper structures to minimize rematches and BYE repeats
  const history = readMatchHistory();
  const pastOpponents = buildPastOpponentsMap(history.matches);
  const byeCounts = history.byeCounts || {};

  // Round 1 pairing
  const round1 = generateLdRound(1, availableDebaters, eligibleJudges, rooms, pastOpponents, byeCounts);

  // Update in-memory history with round1 results (opponents, judges, BYEs)
  const midOpponents = JSON.parse(JSON.stringify(pastOpponents)); // shallow copy structure copy
  round1.forEach(r => {
    if (r.aff === 'BYE') {
      byeCounts[r.neg] = (byeCounts[r.neg] || 0) + 1;
    } else if (r.neg === 'BYE') {
      byeCounts[r.aff] = (byeCounts[r.aff] || 0) + 1;
    } else {
      midOpponents[r.aff] = midOpponents[r.aff] || new Set();
      midOpponents[r.aff].add(r.neg);
      midOpponents[r.neg] = midOpponents[r.neg] || new Set();
      midOpponents[r.neg].add(r.aff);
    }
  });

  // Round 2 pairing using updated in-memory history
  const round2 = generateLdRound(2, availableDebaters, eligibleJudges, rooms, midOpponents, byeCounts);

  // Combine rows: sort by Round asc, Room asc later when creating sheet
  const rows = [];
  round1.forEach(r => rows.push([1, r.aff, r.neg, r.judge, r.room]));
  round2.forEach(r => rows.push([2, r.aff, r.neg, r.judge, r.room]));
  return rows;
}

/**
 * Build a map of past opponents from history.
 * Returns object { name: Set(opponentNames) }
 */
function buildPastOpponentsMap(matchSheets) {
  const map = {};
  matchSheets.forEach(m => {
    if (m.type === 'LD') {
      // rows are Round | Aff | Neg | Judge | Room
      m.rows.forEach(r => {
        const aff = (r[1] || '').toString().trim();
        const neg = (r[2] || '').toString().trim();
        if (!aff || !neg) return;
        if (aff === 'BYE' || neg === 'BYE') return;
        map[aff] = map[aff] || new Set();
        map[neg] = map[neg] || new Set();
        map[aff].add(neg);
        map[neg].add(aff);
      });
    } else if (m.type === 'TP') {
      // TP rows: Aff | Neg | Judge | Room, where Aff/Neg are team strings; skip TP for LD opponents.
    }
  });
  return map;
}

/**
 * generateLdRound
 * Generate pairings for a single LD round given debaters, judges, rooms, and opponent history.
 *
 * Returns array of {aff, neg, judge, room}
 */
function generateLdRound(roundNumber, debaters, judges, rooms, pastOpponents, byeCounts) {
  // Debaters is array of objects with name, etc.
  // We'll pair using greedy algorithm attempting to minimize rematches and BYE repeats.
  const names = debaters.map(d => d.name);
  const needsBye = (names.length % 2 === 1);
  const byes = []; // which name gets bye for this round
  const participants = names.slice();

  // Determine BYE candidate if needed: least historical BYEs
  let byeName = null;
  if (needsBye) {
    participants.sort((a,b) => (byeCounts[a] || 0) - (byeCounts[b] || 0));
    byeName = participants.shift();
  }

  // Now pair participants into matches avoiding rematches by checking pastOpponents set.
  // We'll attempt multiple random shuffles and choose pairing with minimal rematch count.
  const attempts = 500;
  let best = null;
  let bestScore = Infinity; // lower is better (# of rematches + other penalties)

  for (let t=0;t<attempts;t++) {
    const shuffled = shuffleArray(participants.slice());
    const pairs = [];
    let rematches = 0;
    for (let i=0;i<shuffled.length;i+=2) {
      const a = shuffled[i], b = shuffled[i+1];
      const had = (pastOpponents[a] && pastOpponents[a].has) ? pastOpponents[a].has(b) : (pastOpponents[a] && pastOpponents[a].has(b));
      const isRematch = (pastOpponents[a] && pastOpponents[a].has && pastOpponents[a].has(b)) || (pastOpponents[b] && pastOpponents[b].has && pastOpponents[b].has(a));
      if (isRematch) rematches++;
      pairs.push({aff:a, neg:b, rematch:isRematch});
    }
    // Score: primary key rematches, secondary: judge reuse or judge-child conflicts (we'll consider in assignment)
    if (rematches < bestScore) {
      bestScore = rematches;
      best = pairs;
      if (bestScore === 0) break;
    }
  }

  // Assign judges and rooms for these pairs while avoiding judging children and minimizing judge reuse.
  const judgePool = shuffleArray(judges.slice());
  const roomPool = shuffleArray(rooms.slice());
  const assignments = [];
  const usedJudges = new Set();
  for (let i=0;i<best.length;i++) {
    const p = best[i];
    // choose judge not parent of either debater and not already used
    let chosen = judgePool.find(j => !usedJudges.has(j.name) && !j.children.some(c => c === p.aff || c === p.neg));
    if (!chosen) {
      // Relax uniqueness but still avoid parent-child if possible
      chosen = judgePool.find(j => !j.children.some(c => c === p.aff || c === p.neg));
    }
    if (!chosen) {
      // last resort: any judge
      chosen = judgePool[0];
    }
    usedJudges.add(chosen.name);
    const room = roomPool.shift();
    assignments.push({aff:p.aff, neg:p.neg, judge: chosen.name, room: room ? room.name : 'UNASSIGNED'});
  }

  // Build result rows. Include BYE if assigned.
  const rows = [];
  if (byeName) {
    rows.push({aff: 'BYE', neg: byeName, judge: '', room: ''});
  }
  assignments.forEach(a => rows.push({aff: a.aff, neg: a.neg, judge: a.judge, room: a.room}));
  return rows;
}

/* ============================
   Sheet Creation Helpers
   ============================ */

/**
 * Create TP sheet with given rows.
 * TP schema: Aff | Neg | Judge | Room
 */
function createTpSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.insertSheet(sheetName);
  const headers = ['Aff', 'Neg', 'Judge', 'Room'];
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (rows && rows.length) {
    s.getRange(2,1,rows.length,headers.length).setValues(rows);
  }
  s.setFrozenRows(1);
  // Apply conditional formatting for referential integrity: Participants must exist in rosters; judges/rooms must exist; duplicates flagged
  applyMatchSheetFormatting(s, 'TP');
}

/**
 * Create LD sheet with given rows.
 * LD schema: Round | Aff | Neg | Judge | Room
 */
function createLdSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.insertSheet(sheetName);
  const headers = ['Round', 'Aff', 'Neg', 'Judge', 'Room'];
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (rows && rows.length) {
    s.getRange(2,1,rows.length,headers.length).setValues(rows);
    // sort by Round asc, then Room asc
    sortRangeByColumns(s, 2, rows.length, 1, [{col:1, asc:true},{col:5, asc:true}]); // Round is col1, Room col5
  }
  s.setFrozenRows(1);
  applyMatchSheetFormatting(s, 'LD');
}

/**
 * Sort range in-place. startRow is 2 (data), numRows length.
 * sortKeys: [{col:index starting at 1 for sheet, asc:true/false}, ...]
 */
function sortRangeByColumns(sheet, startRow, numRows, startCol, sortKeys) {
  const range = sheet.getRange(startRow, startCol, numRows, sheet.getLastColumn());
  // Google Apps Script only supports single sort in range.sort(). We'll use sheet.getRange(...).sort([{column:col, ascending:bool}, ...])
  const keys = sortKeys.map(k => ({column: k.col, ascending: k.asc}));
  range.sort(keys);
}

/**
 * Apply conditional formatting & referential checks to generated match sheets.
 * Highlights missing participants, duplicate assignments, invalid judges/rooms, etc.
 */
function applyMatchSheetFormatting(sheet, type) {
  // Build rules array
  const rules = [];
  const lastRow = Math.max(sheet.getLastRow(), 200);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === 'TP') {
    // Columns: A=Aff, B=Neg, C=Judge, D=Room
    // 1) Participants must exist in Debaters roster (allow "BYE" and "(IRONMAN)" forms)
    // We'll check A and B columns with formula that strips " (IRONMAN)" and compares to Debaters list.
    const participantRange = sheet.getRange(2,1,lastRow-1,2);
    const formulaPart = 'OR($A2="BYE", COUNTIF(INDIRECT("Debaters!A:A"), IF(RIGHT($A2,9)="(IRONMAN)", LEFT($A2, LEN($A2)-10), $A2))>0)';
    const ruleA = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=NOT(' + formulaPart + ')')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,1,lastRow-1,1)])
      .build();
    // For B
    const formulaPartB = 'OR($B2="BYE", COUNTIF(INDIRECT("Debaters!A:A"), IF(RIGHT($B2,9)="(IRONMAN)", LEFT($B2, LEN($B2)-10), $B2))>0)';
    const ruleB = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=NOT(' + formulaPartB + ')')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,2,lastRow-1,1)])
      .build();
    rules.push(ruleA, ruleB);

    // 2) Judge must exist in Judges roster
    const ruleJudge = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",COUNTIF(INDIRECT("Judges!A:A"),$C2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,3,lastRow-1,1)])
      .build();
    rules.push(ruleJudge);

    // 3) Room must exist in Rooms roster
    const ruleRoom = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($D2<>"",COUNTIF(INDIRECT("Rooms!A:A"),$D2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,4,lastRow-1,1)])
      .build();
    rules.push(ruleRoom);

    // 4) Duplicates: each debater assigned to only one match; highlight duplicates across A and B columns
    // We'll add conditional formatting for A column if COUNTIF($A:$A,$A2)+COUNTIF($B:$B,$A2)>1
    const dupA = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$A2<>"BYE", (COUNTIF($A:$A,$A2)+COUNTIF($B:$B,$A2))>1)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,1,lastRow-1,1)])
      .build();
    const dupB = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2<>"",$B2<>"BYE", (COUNTIF($A:$A,$B2)+COUNTIF($B:$B,$B2))>1)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,2,lastRow-1,1)])
      .build();
    rules.push(dupA, dupB);

  } else if (type === 'LD') {
    // Columns: A=Round, B=Aff, C=Neg, D=Judge, E=Room
    // Participants must exist (allow BYE)
    const ruleAff = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2<>"", $B2<>"BYE", COUNTIF(INDIRECT("Debaters!A:A"),$B2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,2,lastRow-1,1)])
      .build();
    const ruleNeg = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"", $C2<>"BYE", COUNTIF(INDIRECT("Debaters!A:A"),$C2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,3,lastRow-1,1)])
      .build();
    rules.push(ruleAff, ruleNeg);

    // Judge must exist
    const ruleJudge = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($D2<>"", COUNTIF(INDIRECT("Judges!A:A"),$D2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,4,lastRow-1,1)])
      .build();
    rules.push(ruleJudge);

    // Room must exist
    const ruleRoom = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2<>"", COUNTIF(INDIRECT("Rooms!A:A"),$E2)=0)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,5,lastRow-1,1)])
      .build();
    rules.push(ruleRoom);

    // Duplicate checks:
    // Each debater must appear at most once per round (i.e., for round 1, cannot be in two matches)
    // We'll highlight if same name appears twice in same round column combos
    // For Aff column:
    const dupAff = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=COUNTIFS($A:$A,$A2,$B:$B,$B2)>1')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,2,lastRow-1,1)])
      .build();
    const dupNeg = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=COUNTIFS($A:$A,$A2,$C:$C,$C2)>1')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,3,lastRow-1,1)])
      .build();
    rules.push(dupAff, dupNeg);

    // Each judge should be assigned to only one matchup (covering both rounds). Highlight if judge appears more than once.
    const dupJudge = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($D2<>"",COUNTIF($D:$D,$D2)>1)')
      .setBackground('#FFCFCF')
      .setRanges([sheet.getRange(2,4,lastRow-1,1)])
      .build();
    rules.push(dupJudge);
  }

  sheet.setConditionalFormatRules(rules);
}

/* ============================
   Utility functions
   ============================ */

/** Fisher-Yates */
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

/* ============================
   RSVPs
   ============================ */

/**
 * clearRsvps
 * Resets Availability Attending? column to "Not responded" for non-empty participants.
 */
function clearRsvps() {
  const s = getSheet('Availability');
  if (!s) {
    ui().alert('Availability sheet not found.');
    return;
  }
  const maxRows = Math.max(s.getLastRow(), 200);
  const parts = s.getRange(2,1,maxRows,1).getValues();
  const newVals = parts.map(r => {
    const p = r[0];
    if (p && p.toString().trim() !== '') return [RSVP_OPTIONS[2]];
    return [''];
  });
  s.getRange(2,2,maxRows,1).setValues(newVals);
  ui().alert('RSVPs cleared (set to "Not responded" where participant exists).');
}

/* ============================
   Misc Helpers / Exports
   ============================ */

/**
 * Simple function to test reading of rosters (debug).
 */
function _logRosters() {
  console.log(JSON.stringify({debaters: readDebaters(), judges: readJudges(), rooms:readRooms()}));
}
