/**
 * ======================================================================
 * AUTOMATED DEBATE PAIRING TOOL
 * ======================================================================
 * All logic contained within this single file as per requirements.
 * 
 * SECTIONS:
 * 1. Global Constants & Menu
 * 2. Initialization & Admin
 * 3. Validation & Integrity
 * 4. Data Loading & History
 * 5. Pairing Logic Engines (TP & LD)
 * 6. Output & Formatting
 */

/* =========================================
   1. GLOBAL CONSTANTS & MENU
   ========================================= */

const TABS = {
  AVAILABILITY: 'Availability',
  DEBATERS: 'Debaters',
  JUDGES: 'Judges',
  ROOMS: 'Rooms'
};

const TYPES = {
  TP: 'TP',
  LD: 'LD'
};

/**
 * Creates the custom menu when the sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛡️ Club Admin')
    .addItem('Initialize Sheet (First Time Only)', 'initializeSheet')
    .addSeparator()
    .addItem('Generate TP Matches', 'generateTPMatches')
    .addItem('Generate LD Matches', 'generateLDMatches')
    .addSeparator()
    .addItem('Clear RSVPs', 'clearRSVPs')
    .addToUi();

  // Re-apply conditional formatting rules on open to ensure integrity highlights work
  applyGlobalValidationRules_();
}

/* =========================================
   2. INITIALIZATION & ADMIN
   ========================================= */

/**
 * Sets up the four permanent tabs with headers, sample data, and basic formatting.
 * Will not overwrite existing tabs of the same name.
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. DEBATERS TAB
  if (!ss.getSheetByName(TABS.DEBATERS)) {
    const s = ss.insertSheet(TABS.DEBATERS);
    s.appendRow(['Name', 'Debate Type', 'Partner', 'Hard Mode']);
    s.setFrozenRows(1);
    s.getRange("A1:D1").setFontWeight("bold");

    const sampleDebaters = [
      // LD
      ['Abraham Lincoln', 'LD', '', 'Yes'], ['Stephen A. Douglas', 'LD', '', 'Yes'],
      ['Clarence Darrow', 'LD', '', 'No'], ['William Jennings Bryan', 'LD', '', 'No'],
      ['William F. Buckley Jr.', 'LD', '', 'Yes'], ['Gore Vidal', 'LD', '', 'Yes'],
      ['Christopher Hitchens', 'LD', '', 'Yes'], ['Tony Blair', 'LD', '', 'No'],
      ['Jordan Peterson', 'LD', '', 'Yes'], ['Slavoj Žižek', 'LD', '', 'Yes'],
      ['Lloyd Bentsen', 'LD', '', 'No'], ['Dan Quayle', 'LD', '', 'No'],
      ['Richard Dawkins', 'LD', '', 'Yes'], ['Rowan Williams', 'LD', '', 'Yes'],
      ['Diogenes', 'LD', '', 'Yes'],
      // TP
      ['Noam Chomsky', 'TP', 'Michel Foucault', 'Yes'], ['Michel Foucault', 'TP', 'Noam Chomsky', 'Yes'],
      ['Harlow Shapley', 'TP', 'Heber Curtis', 'No'], ['Heber Curtis', 'TP', 'Harlow Shapley', 'No'],
      ['Muhammad Ali', 'TP', 'George Foreman', 'Yes'], ['George Foreman', 'TP', 'Muhammad Ali', 'Yes'],
      ['Richard Nixon', 'TP', 'Nikita Khrushchev', 'No'], ['Nikita Khrushchev', 'TP', 'Richard Nixon', 'No'],
      ['Thomas Henry Huxley', 'TP', 'Samuel Wilberforce', 'Yes'], ['Samuel Wilberforce', 'TP', 'Thomas Henry Huxley', 'Yes'],
      ['John F. Kennedy', 'TP', 'David Frost', 'No'], ['David Frost', 'TP', 'John F. Kennedy', 'No'],
      ['Bob Dole', 'TP', 'Bill Clinton', 'No'], ['Bill Clinton', 'TP', 'Bob Dole', 'No']
    ];
    if (sampleDebaters.length > 0) {
      s.getRange(2, 1, sampleDebaters.length, 4).setValues(sampleDebaters);
    }
    // Data validation for Type and Hard Mode
    s.getRange("B2:B").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TP', 'LD']).build());
    s.getRange("D2:D").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build());
  }

  // 2. JUDGES TAB
  if (!ss.getSheetByName(TABS.JUDGES)) {
    const s = ss.insertSheet(TABS.JUDGES);
    s.appendRow(['Name', 'Children\'s Names (comma-separated)', 'Debate Type']);
    s.setFrozenRows(1);
    s.getRange("A1:C1").setFontWeight("bold");

    const sampleJudges = [
      ['Howard K. Smith', 'John F. Kennedy', 'TP'],
      ['Fons Elders', 'Noam Chomsky, Michel Foucault', 'TP'],
      ['John Stevens Henslow', 'Samuel Wilberforce', 'TP'],
      ['Jim Lehrer', 'Bill Clinton', 'TP'],
      ['Judy Woodruff', '', 'TP'], ['Tom Brokaw', '', 'TP'],
      ['Frank McGee', '', 'TP'], ['Quincy Howe', '', 'TP'],
      ['John T. Raulston', 'Clarence Darrow', 'LD'],
      ['Rudyard Griffiths', 'Jordan Peterson', 'LD'],
      ['Stephen J. Blackwood', '', 'LD'], ['Brit Hume', '', 'LD'],
      ['Jon Margolis', '', 'LD'], ['Bill Shadel', '', 'LD'],
      ['Judge Judy', '', 'LD']
    ];
    if (sampleJudges.length > 0) {
      s.getRange(2, 1, sampleJudges.length, 3).setValues(sampleJudges);
    }
    s.getRange("C2:C").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TP', 'LD']).build());
  }

  // 3. ROOMS TAB
  if (!ss.getSheetByName(TABS.ROOMS)) {
    const s = ss.insertSheet(TABS.ROOMS);
    s.appendRow(['Room Name', 'Debate Type']);
    s.setFrozenRows(1);
    s.getRange("A1:B1").setFontWeight("bold");

    const sampleRooms = [
      ['Room 101', 'LD'], ['Room 102', 'LD'], ['Room 103', 'LD'], ['Room 201', 'LD'],
      ['Sanctuary right', 'LD'], ['Sanctuary left', 'LD'], ['Pantry', 'LD'],
      ['Chapel', 'TP'], ['Library', 'TP'], ['Music lounge', 'TP'],
      ['Cry room', 'TP'], ['Office', 'TP'], ['Office hallway', 'TP']
    ];
    if (sampleRooms.length > 0) {
      s.getRange(2, 1, sampleRooms.length, 2).setValues(sampleRooms);
    }
    s.getRange("B2:B").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TP', 'LD']).build());
  }

  // 4. AVAILABILITY TAB
  if (!ss.getSheetByName(TABS.AVAILABILITY)) {
    const s = ss.insertSheet(TABS.AVAILABILITY);
    s.appendRow(['Participant', 'Attending?']);
    s.setFrozenRows(1);
    s.getRange("A1:B1").setFontWeight("bold");

    // Formula to aggregate UNIQUE names from Debaters and Judges, sorted, excluding blanks.
    // Uses FILTER to ignore blank rows before UNIQUE to be safe.
    const formula = `=SORT(UNIQUE({FILTER(Debaters!A2:A, Debaters!A2:A<>""); FILTER(Judges!A2:A, Judges!A2:A<>"")}))`;
    s.getRange("A2").setFormula(formula);

    // Data validation for Attending?
    s.getRange("B2:B").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No', 'Not responded'], true).build());
  }

  // Final cleanup
  deleteSheet1_(ss);
  sortTabs_(ss);
  applyGlobalValidationRules_();
  ui.alert("Initialization Complete", "Roster tabs created. Please fill in your actual club data.", ui.Button.OK);
}

/**
 * Clears all RSVPs in the Availability tab to 'Not responded'.
 */
function clearRSVPs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.AVAILABILITY);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // Reset all to 'Not responded' where a participant exists
    const participantRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rsvpRange = sheet.getRange(2, 2, lastRow - 1, 1);
    const newValues = participantRange.map(row => row[0] ? ['Not responded'] : ['']);
    rsvpRange.setValues(newValues);
  }
  SpreadsheetApp.getActive().toast("RSVPs have been reset.", "Success");
}

/* =========================================
   3. VALIDATION & INTEGRITY
   ========================================= */

/**
 * Main entry point for TP generation.
 */
function generateTPMatches() {
  if (!preFlightCheck_(TYPES.TP)) return;
  generateMatches_(TYPES.TP);
}

/**
 * Main entry point for LD generation.
 */
function generateLDMatches() {
  if (!preFlightCheck_(TYPES.LD)) return;
  generateMatches_(TYPES.LD);
}

/**
 * Runs critical integrity checks before attempting generation.
 * Halts execution and alerts user if data is corrupt.
 * @param {string} type - TP or LD
 * @return {boolean} True if passed, false if failed.
 */
function preFlightCheck_(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Check for today's sheet to prevent overwrite
  const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const sheetName = `${type} ${dateStr}`;
  if (ss.getSheetByName(sheetName)) {
    ui.alert("Error: Sheet Exists", `A pairing sheet for today (${sheetName}) already exists. Please delete it if you wish to regenerate.`, ui.Button.OK);
    return false;
  }

  // 2. TP Specific: Partner Reciprocity & Hard Mode Consistency
  if (type === TYPES.TP) {
    const debaters = getData_(ss, TABS.DEBATERS); // [Name, Type, Partner, HardMode]
    const debaterMap = new Map();
    debaters.forEach(d => debaterMap.set(d[0], { type: d[1], partner: d[2], hm: d[3] }));

    for (let i = 0; i < debaters.length; i++) {
      const [name, dType, partner, hm] = debaters[i];
      if (dType !== TYPES.TP) continue;

      if (!partner) {
        ui.alert("Data Error", `TP Debater '${name}' has no partner listed.`, ui.Button.OK);
        return false;
      }
      const pData = debaterMap.get(partner);
      if (!pData) {
        ui.alert("Data Error", `Partner '${partner}' for '${name}' does not exist in the Debaters roster.`, ui.Button.OK);
        return false;
      }
      if (pData.partner !== name) {
        ui.alert("Data Error", `Partnership mismatch: '${name}' lists '${partner}', but '${partner}' lists '${pData.partner}'.`, ui.Button.OK);
        return false;
      }
      if (pData.hm !== hm) {
        ui.alert("Data Error", `Hard Mode mismatch for team '${name}' & '${partner}'. Both must have the same setting.`, ui.Button.OK);
        return false;
      }
    }
  }

  // 3. Judge Children Existence Check
  const judges = getData_(ss, TABS.JUDGES);
  const allDebaters = new Set(getData_(ss, TABS.DEBATERS).map(d => d[0]));

  for (let i = 0; i < judges.length; i++) {
    const childrenStr = judges[i][1];
    if (childrenStr) {
      const children = childrenStr.split(',').map(s => s.trim());
      for (const child of children) {
        if (child && !allDebaters.has(child)) {
           ui.alert("Data Error", `Judge '${judges[i][0]}' lists unknown child '${child}'. Please check Debaters roster.`, ui.Button.OK);
           return false;
        }
      }
    }
  }

  return true;
}

/**
 * Applies conditional formatting rules to permanent tabs for visual validation.
 * Should be called onOpen and after initialization.
 */
function applyGlobalValidationRules_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Availability Tab Rules
  const availSheet = ss.getSheetByName(TABS.AVAILABILITY);
  if (availSheet) {
    const rules = [];
    // Red if name exists but Attending is blank
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(NOT(ISBLANK($A2)), ISBLANK($B2))')
      .setBackground('#F4CCCC') // Light red
      .setRanges([availSheet.getRange("B2:B")])
      .build());
    availSheet.setConditionalFormatRules(rules);
  }
  
  // Note: Cross-sheet validation (e.g. checking if a partner exists in Debaters tab)
  // requires INDIRECT which can be slow and complex to maintain in pure CF.
  // Relying mostly on preFlightCheck_ for hard data integrity.
}

/* =========================================
   4. DATA LOADING & HISTORY
   ========================================= */

/**
 * Loads all necessary data into standardized objects.
 * @param {string} type - TP or LD
 */
function loadEnvironment_(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Availability Map
  const availRaw = getData_(ss, TABS.AVAILABILITY);
  const availableSet = new Set(
    availRaw.filter(r => r[1] === 'Yes').map(r => r[0])
  );

  // 2. Load Debaters
  const allDebaters = getData_(ss, TABS.DEBATERS);
  const typeDebaters = allDebaters.filter(d => d[1] === type && availableSet.has(d[0]));
  
  // 3. Load Judges
  const allJudges = getData_(ss, TABS.JUDGES);
  const typeJudges = allJudges.filter(j => j[2] === type && availableSet.has(j[0]));

  // 4. Load Rooms
  const allRooms = getData_(ss, TABS.ROOMS);
  const typeRooms = allRooms.filter(r => r[1] === type).map(r => r[0]);

  // 5. Build History
  const history = buildHistory_(ss, type);

  return {
    debaters: typeDebaters.map(d => ({
      name: d[0], 
      partner: d[2], 
      hm: d[3] === 'Yes',
      history: history.debaters[d[0]] || {opponents: [], judges: [], byes: 0}
    })),
    judges: typeJudges.map(j => ({
      name: j[0],
      children: j[1] ? j[1].split(',').map(s => s.trim()) : [],
      history: history.judges[j[0]] || {rooms: [], debaters: []}
    })),
    rooms: typeRooms, // simple array of strings
    fullDebaterRoster: new Map(allDebaters.map(d => [d[0], d])) // For Ironman lookups if needed
  };
}

/**
 * Scans past sheets to build match history.
 */
function buildHistory_(ss, type) {
  const history = { debaters: {}, judges: {} };
  const sheets = ss.getSheets();
  const regex = new RegExp(`^${type} \\d{4}-\\d{2}-\\d{2}$`);

  sheets.forEach(s => {
    if (regex.test(s.getName())) {
      const data = s.getDataRange().getValues();
      // Start at row 1 (skip header).
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (type === TYPES.TP) {
           // TP: Aff | Neg | Judge | Room
           processHistoryRow_(history, row[0], row[1], row[2], row[3]);
        } else {
           // LD: Round | Aff | Neg | Judge | Room
           processHistoryRow_(history, row[1], row[2], row[3], row[4]);
        }
      }
    }
  });
  return history;
}

function processHistoryRow_(h, affRaw, negRaw, judge, room) {
  if (!affRaw || !negRaw) return;

  // Helper to clean names (remove ironman tag, split teams if TP history was stored as combined string, 
  // though current design stores team names. Let's assume standard debater names).
  // Actually TP pairs are teams.
  
  const clean = (name) => name.replace(" (IRONMAN)", "").trim();
  const aff = clean(affRaw);
  const neg = clean(negRaw);

  // Ensure objects exist
  if (!h.debaters[aff]) h.debaters[aff] = {opponents: [], judges: [], byes: 0};
  if (!h.debaters[neg]) h.debaters[neg] = {opponents: [], judges: [], byes: 0};
  if (judge && !h.judges[judge]) h.judges[judge] = {rooms: [], debaters: []};

  // Log match
  if (neg === "BYE") {
    h.debaters[aff].byes++;
  } else {
    h.debaters[aff].opponents.push(neg);
    h.debaters[neg].opponents.push(aff);
  }

  // Log judge
  if (judge && judge !== "") {
    if (neg !== "BYE") {
       h.debaters[aff].judges.push(judge);
       h.debaters[neg].judges.push(judge);
       h.judges[judge].debaters.push(aff, neg);
    } else {
       // If it was a BYE, they might still have been assigned a judge in error, but let's ignore.
    }
    if (room) h.judges[judge].rooms.push(room);
  }
}

/* =========================================
   5. PAIRING LOGIC ENGINES
   ========================================= */

/**
 * Orchestrates the generation process.
 */
function generateMatches_(type) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Load Environment
  const env = loadEnvironment_(type);
  
  if (env.debaters.length === 0) {
    ui.alert("No Participants", `No ${type} debaters are marked 'Yes' in Availability.`, ui.Button.OK);
    return;
  }

  let finalPairings = [];

  // 2. Run specific pairing logic
  if (type === TYPES.TP) {
    finalPairings = runTPPairing_(env, ui);
  } else {
    finalPairings = runLDPairing_(env, ui);
  }

  if (!finalPairings) return; // Failed due to critical resource shortage

  // 3. Output results
  const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const sheetName = `${type} ${dateStr}`;
  const newSheet = ss.insertSheet(sheetName);

  if (type === TYPES.TP) {
    newSheet.appendRow(['Affirmative', 'Negative', 'Judge', 'Room']);
    if (finalPairings.length > 0) {
      newSheet.getRange(2, 1, finalPairings.length, 4).setValues(finalPairings);
    }
  } else {
    // LD
    newSheet.appendRow(['Round', 'Affirmative', 'Negative', 'Judge', 'Room']);
    if (finalPairings.length > 0) {
       // Sort by Round then Room for LD
       finalPairings.sort((a,b) => a[0] - b[0] || (a[4] || "").localeCompare(b[4] || ""));
       newSheet.getRange(2, 1, finalPairings.length, 5).setValues(finalPairings);
    }
  }

  // 4. Final Formatting
  newSheet.setFrozenRows(1);
  newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).setFontWeight("bold");
  sortTabs_(ss);
  
  // Apply integrity highlighting to the new sheet
  applyMatchSheetValidation_(newSheet, type);
}

/**
 * TP Pairing Algorithm
 */
function runTPPairing_(env, ui) {
  // 1. Form Teams (handle Ironmen)
  const teams = formTPTeams_(env.debaters);
  
  // 2. Handle BYE if odd
  let byeTeam = null;
  if (teams.length % 2 !== 0) {
    // Sort by num BYEs ascending, then random
    teams.sort((a,b) => a.avgByes - b.avgByes || 0.5 - Math.random());
    byeTeam = teams.shift();
  }

  // 3. Resource Check
  const matchesNeeded = Math.ceil(teams.length / 2);
  if (env.judges.length < matchesNeeded) {
    ui.alert("Insufficient Judges", `Need ${matchesNeeded} judges for TP, but only have ${env.judges.length} available.`, ui.Button.OK);
    return null;
  }
  if (env.rooms.length < matchesNeeded) {
    ui.alert("Insufficient Rooms", `Need ${matchesNeeded} rooms for TP, but only have ${env.rooms.length} available.`, ui.Button.OK);
    return null;
  }

  // 4. Pair Teams (Simple greedy with scoring)
  // Shuffle first for randomness base
  teams.sort(() => 0.5 - Math.random());
  
  const pairings = [];
  const pairedIndices = new Set();

  for (let i = 0; i < teams.length; i++) {
    if (pairedIndices.has(i)) continue;
    
    let bestScore = -Infinity;
    let bestOpponentIndex = -1;

    for (let j = i + 1; j < teams.length; j++) {
      if (pairedIndices.has(j)) continue;
      const score = scoreTPMatchup_(teams[i], teams[j]);
      if (score > bestScore) {
        bestScore = score;
        bestOpponentIndex = j;
      }
    }

    if (bestOpponentIndex !== -1) {
      pairings.push({ aff: teams[i], neg: teams[bestOpponentIndex] });
      pairedIndices.add(i);
      pairedIndices.add(bestOpponentIndex);
    }
  }

  // 5. Assign Judges & Rooms
  const usedJudges = new Set();
  const usedRooms = new Set();
  const output = [];

  pairings.forEach(p => {
    // Find best judge
    let bestJudge = null;
    let bestJScore = -Infinity;

    for (const judge of env.judges) {
      if (usedJudges.has(judge.name)) continue;
      
      // HARD CONSTRAINT: Parent/Child
      if (isParentOfTeam_(judge, p.aff) || isParentOfTeam_(judge, p.neg)) continue;

      const score = scoreJudgeForTeam_(judge, p.aff) + scoreJudgeForTeam_(judge, p.neg);
      if (score > bestJScore) {
        bestJScore = score;
        bestJudge = judge;
      }
    }

    // Fallback if no valid judge found (extremely rare if resource check passed, but possible with many conflicts)
    // In a real scenario, we might need to swap judges. For simplicity here, we take the first non-conflicted, or just force it if desperate.
    // If truly stuck, leave judge blank to alert admin.
    const judgeName = bestJudge ? bestJudge.name : "CONFLICT ERROR";
    if (bestJudge) usedJudges.add(bestJudge.name);

    // Assign Room (just pick next available for simplicity, or score by affinity if requested)
    const room = env.rooms.find(r => !usedRooms.has(r)) || "NO ROOM";
    usedRooms.add(room);

    output.push([p.aff.displayName, p.neg.displayName, judgeName, room]);
  });

  if (byeTeam) {
    output.push([byeTeam.displayName, "BYE", "", ""]);
  }

  return output;
}

/**
 * LD Pairing Algorithm (Two Rounds)
 */
function runLDPairing_(env, ui) {
  let debaters = [...env.debaters];
  
  // Resource check for one round (need enough for max concurrent matches)
  const maxMatchesPerRound = Math.ceil(debaters.length / 2);
  // LD Judges can judge both rounds, so we just need enough unique judges for one round's worth of matches.
  // Actually, they judge a "flight".
  if (env.judges.length < maxMatchesPerRound) {
     ui.alert("Insufficient Judges", `Need ${maxMatchesPerRound} judges for LD concurrent rounds, have ${env.judges.length}.`, ui.Button.OK);
     return null;
  }
  if (env.rooms.length < maxMatchesPerRound) {
     ui.alert("Insufficient Rooms", `Need ${maxMatchesPerRound} rooms for LD, have ${env.rooms.length}.`, ui.Button.OK);
     return null;
  }

  const round1 = generateRound_(debaters, env.judges, env.rooms, 1, []);
  
  // Update history in-memory for Round 2
  // Deep copy isn't strictly needed if we just append to the history arrays temporarily.
  round1.matches.forEach(m => {
    if (m.neg === "BYE") return; // BYE handled separately
    // Add to temp history to avoid rematch in R2
    findDebater_(debaters, m.aff).history.opponents.push(m.neg);
    findDebater_(debaters, m.neg).history.opponents.push(m.aff);
  });
  // Also update BYE history temporarily so same person doesn't get it twice
  if (round1.bye) {
    findDebater_(debaters, round1.bye).history.byes++;
  }

  const round2 = generateRound_(debaters, env.judges, env.rooms, 2, [round1.bye]);

  // Combine outputs: [Round, Aff, Neg, Judge, Room]
  const formatOutput = (rNum, matches, bye) => {
    const rows = matches.map(m => [rNum, m.aff, m.neg, m.judge, m.room]);
    if (bye) rows.push([rNum, bye, "BYE", "", ""]);
    return rows;
  };

  return [...formatOutput(1, round1.matches, round1.bye), ...formatOutput(2, round2.matches, round2.bye)];
}

function generateRound_(debatersPool, judges, rooms, roundNum, excludedFromBye) {
  // 1. Handle BYE
  let byeDebater = null;
  let activePool = [...debatersPool];
  if (activePool.length % 2 !== 0) {
     activePool.sort((a,b) => a.history.byes - b.history.byes || 0.5 - Math.random());
     // Ensure we don't give BYE to someone excluded (e.g. had it R1)
     let byeIdx = 0;
     while (excludedFromBye.includes(activePool[byeIdx].name) && byeIdx < activePool.length - 1) {
       byeIdx++;
     }
     byeDebater = activePool[byeIdx];
     activePool.splice(byeIdx, 1);
  }

  // 2. Pair (Greedy scored)
  activePool.sort(() => 0.5 - Math.random());
  const matches = [];
  const paired = new Set();

  for (let i = 0; i < activePool.length; i++) {
    if (paired.has(i)) continue;
    let bestScore = -Infinity;
    let bestOp = -1;
    for (let j = i + 1; j < activePool.length; j++) {
      if (paired.has(j)) continue;
      // Score LD match: prefer same Hard Mode, strong avoid previous opponents
      let score = 1000;
      if (activePool[i].hm === activePool[j].hm) score += 100;
      if (activePool[i].history.opponents.includes(activePool[j].name)) score -= 5000; // Huge penalty for immediate rematch
      
      if (score > bestScore) { bestScore = score; bestOp = j; }
    }
    if (bestOp !== -1) {
      matches.push({ aff: activePool[i], neg: activePool[bestOp] });
      paired.add(i);
      paired.add(bestOp);
    }
  }

  // 3. Assign Judges & Rooms (Greedy)
  // Note: For LD R2, we *could* try to keep judges in same rooms, but simpler to just re-assign fresh to ensure no conflicts with new matchups.
  const usedJudges = new Set();
  const usedRooms = new Set();
  const finalMatches = matches.map(m => {
    // Find judge
    let bestJudge = null;
    let bestScore = -Infinity;
    for (const j of judges) {
      if (usedJudges.has(j.name)) continue;
      if (j.children.includes(m.aff.name) || j.children.includes(m.neg.name)) continue; // Hard conflict

      let score = 1000;
      if (m.aff.history.judges.includes(j.name)) score -= 200;
      if (m.neg.history.judges.includes(j.name)) score -= 200;
      
      if (score > bestScore) { bestScore = score; bestJudge = j; }
    }
    if (bestJudge) usedJudges.add(bestJudge.name);

    const room = rooms.find(r => !usedRooms.has(r)) || "NO ROOM";
    usedRooms.add(room);

    return { aff: m.aff.name, neg: m.neg.name, judge: bestJudge ? bestJudge.name : "CONFLICT", room: room };
  });

  return { matches: finalMatches, bye: byeDebater ? byeDebater.name : null };
}


/* =========================================
   6. HELPERS & FORMATTING
   ========================================= */

/**
 * Helper to find a debater object by name in an array.
 */
function findDebater_(list, name) {
  return list.find(d => d.name === name);
}

/**
 * Forms TP teams from individual debaters. Handles Ironmen.
 */
function formTPTeams_(debaters) {
  const teamMap = new Map();
  const processed = new Set();

  debaters.forEach(d => {
    if (processed.has(d.name)) return;

    const partner = debaters.find(p => p.name === d.partner);
    if (partner) {
      // Full team present
      const teamName = [d.name, partner.name].sort().join(" & ");
      teamMap.set(teamName, {
        displayName: teamName,
        members: [d.name, partner.name],
        hm: d.hm, // Assume validated consistency
        avgByes: (d.history.byes + partner.history.byes) / 2,
        history: d.history // simplified: use one partner's history for ad-hoc scoring or merge them if complex
      });
      processed.add(d.name);
      processed.add(partner.name);
    } else {
      // Ironman
      const displayName = `${d.name} (IRONMAN)`;
      teamMap.set(displayName, {
        displayName: displayName,
        members: [d.name],
        hm: d.hm,
        avgByes: d.history.byes,
        history: d.history
      });
      processed.add(d.name);
    }
  });
  return Array.from(teamMap.values());
}

/**
 * Checks if judge is parent of any member of a TP team.
 */
function isParentOfTeam_(judge, team) {
  return team.members.some(member => judge.children.includes(member));
}

/**
 * Scores a potential TP matchup. Higher is better.
 */
function scoreTPMatchup_(teamA, teamB) {
  let score = 1000;
  // Soft: Hard Mode matching
  if (teamA.hm === teamB.hm) score += 100;
  
  // Soft: Minimize rematches. Check if ANY member of A matched ANY member of B recently.
  // This is a simplified check against the team history representative.
  // A full check would loop all members against all members' histories.
  const aMetB = teamA.members.some(m => teamA.history.opponents.some(op => teamB.displayName.includes(op))); // rudimentary check
  if (aMetB) score -= 500;

  return score;
}

/**
 * Scores a judge for a team.
 */
function scoreJudgeForTeam_(judge, team) {
  let score = 0;
  // Avoid re-judging
  const rejudge = team.members.some(m => team.history.judges.includes(judge.name));
  if (rejudge) score -= 200;
  return score;
}

/**
 * Sorts tabs: Permanent 4 first, then others by date descending (LD before TP on same date).
 */
function sortTabs_(ss) {
  const permanent = [TABS.AVAILABILITY, TABS.DEBATERS, TABS.JUDGES, TABS.ROOMS];
  const sheets = ss.getSheets();
  const sheetMap = sheets.map(s => ({ sheet: s, name: s.getName() }));

  sheetMap.sort((a, b) => {
    const aPermIdx = permanent.indexOf(a.name);
    const bPermIdx = permanent.indexOf(b.name);

    if (aPermIdx !== -1 && bPermIdx !== -1) return aPermIdx - bPermIdx;
    if (aPermIdx !== -1) return -1;
    if (bPermIdx !== -1) return 1;

    // Sort generated tabs by date desc, then type (LD before TP -> LD is 'smaller' string? No, LD vs TP. L comes before T.
    // If we want LD before TP for SAME date, and we are sorting descending, we want later dates first.
    // Name format: "TYPE YYYY-MM-DD"
    // Let's parse them.
    const parseSheetName = (name) => {
      const parts = name.split(' ');
      return { type: parts[0], date: parts[1] || '' };
    };
    const aP = parseSheetName(a.name);
    const bP = parseSheetName(b.name);

    if (aP.date !== bP.date) {
      return bP.date.localeCompare(aP.date); // Date descending
    }
    return aP.type.localeCompare(bP.type); // LD before TP (alphabetical works: LD < TP)
  });

  sheetMap.forEach((so, index) => {
    try { so.sheet.activate(); ss.moveActiveSheet(index + 1); } catch(e) {} // activate/move can sometimes fail slightly if too fast, but usually ok.
  });
  
  // Switch back to first tab for neatness
  ss.getSheetByName(TABS.AVAILABILITY).activate();
}

/**
 * Gets all data from a sheet, skipping header.
 */
function getData_(ss, tabName) {
  const s = ss.getSheetByName(tabName);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues();
}

/**
 * Deletes the default "Sheet1" if it exists and isn't needed.
 */
function deleteSheet1_(ss) {
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

/**
 * Applies integrity checking conditional formatting to a newly generated match sheet.
 */
function applyMatchSheetValidation_(sheet, type) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Ranges
  // TP: A=Aff, B=Neg, C=Judge, D=Room
  // LD: A=Round, B=Aff, C=Neg, D=Judge, E=Room
  const affCol = type === TYPES.TP ? "A" : "B";
  const negCol = type === TYPES.TP ? "B" : "C";
  const judgeCol = type === TYPES.TP ? "C" : "D";
  const roomCol = type === TYPES.TP ? "D" : "E";

  const rules = [];
  const r = (col) => sheet.getRange(`${col}2:${col}${lastRow}`);

  // 1. Unknown Participant/Resource (Simple INDIRECT checks are too heavy, using simple COUNTIF within sheet if possible, but rosters are external.
  // Due to Apps Script limitations on complex CF formulas with external references efficiently,
  // we'll stick to internal consistency checks which are most critical for immediate "did the script break" feedback.
  // (Real external validation is best done by the pre-flight check script, not fragile CF).

  // 2. Duplicates (Internal Consistency)
  // Highlight if same Debater is listed twice in Aff or Neg cols (ignore BYE)
  // For LD, it's allowed TWICE (once per round). This is hard to CF simply.
  // Let's just do basic "Duplicates in same column" for TP.
  if (type === TYPES.TP) {
     const duplicateFormula = `=COUNTIF($A$2:$B, INDIRECT(ADDRESS(ROW(), COLUMN()))) > 1`;
     // Applying to A & B combined is tricky.
  }
  
  // Just simple "Is blank but shouldn't be" for now as a basic check
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground("#F4CCCC")
    .setRanges([r(affCol), r(negCol)]) // Aff/Neg shouldn't be empty. Judge/Room might be if BYE.
    .build());

  sheet.setConditionalFormatRules(rules);
}
