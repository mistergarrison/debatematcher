/**
 * @OnlyCurrentDoc
 *
 * This script provides an automated debate pairing tool for a Google Sheet.
 * It manages rosters, availability, and generates balanced pairings for
 * Team Policy (TP) and Lincoln-Douglas (LD) debate formats, while respecting
 * constraints like judge conflicts and resource availability.
 */

// --- CONSTANTS ---

// Menu
const ADMIN_MENU_NAME = 'Club Admin';

// Tab Names
const AVAILABILITY_TAB = 'Availability';
const DEBATERS_TAB = 'Debaters';
const JUDGES_TAB = 'Judges';
const ROOMS_TAB = 'Rooms';
const PERMANENT_TABS = [AVAILABILITY_TAB, DEBATERS_TAB, JUDGES_TAB, ROOMS_TAB];

// Debate Types
const TP_TYPE = 'TP';
const LD_TYPE = 'LD';

// Column Headers
const DEBATER_HEADERS = ['Name', 'Debate Type', 'Partner', 'Hard Mode'];
const JUDGE_HEADERS = ['Name', 'Children\'s Names', 'Debate Type'];
const ROOM_HEADERS = ['Room Name', 'Debate Type'];
const AVAILABILITY_HEADERS = ['Participant', 'Attending?'];
const TP_MATCH_HEADERS = ['Aff', 'Neg', 'Judge', 'Room'];
const LD_MATCH_HEADERS = ['Round', 'Aff', 'Neg', 'Judge', 'Room'];

// Values
const RSVP_YES = 'Yes';
const RSVP_NO = 'No';
const RSVP_DEFAULT = 'Not responded';
const HARD_MODE_YES = 'Yes';
const BYE_TEAM = 'BYE';
const IRONMAN_SUFFIX = ' (IRONMAN)';

// Formatting
const INVALID_ROW_COLOR = '#f4cccc';  // Light red 1

// --- SPREADSHEET UI & TRIGGERS ---

/**
 * Creates a custom menu in the spreadsheet UI when the document is opened.
 * Also applies conditional formatting rules to ensure data integrity is
 * visible.
 * @param {Object} e The event object.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu(ADMIN_MENU_NAME)
      .addItem('Initialize Sheet', 'initializeSheet')
      .addSeparator()
      .addItem('Generate TP Matches', 'generateTpMatches')
      .addItem('Generate LD Matches', 'generateLdMatches')
      .addSeparator()
      .addItem('Clear All RSVPs', 'clearRsvps')
      .addToUi();
  applyAllConditionalFormats();
}


// --- MENU FUNCTIONS ---

/**
 * Sets up the spreadsheet with the required tabs, headers, sample data,
 * and formatting. It is idempotent and will not overwrite existing tabs.
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const existingSheets = ss.getSheets().map(s => s.getName());
  let initializedCount = 0;

  // Function to create and format a sheet
  const createSheet = (name, headers, sampleData) => {
    if (existingSheets.includes(name)) {
      return;  // Skip if sheet already exists
    }
    const sheet = ss.insertSheet(name, ss.getSheets().length);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    if (sampleData && sampleData.length > 0) {
      sheet.getRange(2, 1, sampleData.length, sampleData[0].length)
          .setValues(sampleData);
    }
    initializedCount++;
  };

  // Create permanent tabs with sample data
  createSheet(DEBATERS_TAB, DEBATER_HEADERS, getSampleDebaters());
  createSheet(JUDGES_TAB, JUDGE_HEADERS, getSampleJudges());
  createSheet(ROOMS_TAB, ROOM_HEADERS, getSampleRooms());

  // Special setup for Availability Tab
  if (!existingSheets.includes(AVAILABILITY_TAB)) {
    const sheet = ss.insertSheet(AVAILABILITY_TAB, 0);
    const headerRange = sheet.getRange(1, 1, 1, AVAILABILITY_HEADERS.length);
    headerRange.setValues([AVAILABILITY_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Set formula to auto-populate participants
    const formula = `=IFERROR(SORT(UNIQUE({${DEBATERS_TAB}!A2:A; ${
        JUDGES_TAB}!A2:A})), "")`;
    sheet.getRange('A2').setFormula(formula);

    // Set data validation for 'Attending?' column
    const rule =
        SpreadsheetApp.newDataValidation()
            .requireValueInList([RSVP_YES, RSVP_NO, RSVP_DEFAULT], true)
            .setAllowInvalid(false)
            .build();
    sheet.getRange('B2:B').setDataValidation(rule);

    clearRsvps();  // Populate with 'Not responded'
    initializedCount++;
  }

  if (initializedCount > 0) {
    applyAllConditionalFormats();
    sortAllTabs();
    ss.setActiveSheet(ss.getSheetByName(AVAILABILITY_TAB));
    ui.alert(
        'Sheet Initialized',
        'Required tabs and sample data have been created. Please review the rosters.',
        ui.ButtonSet.OK);
  } else {
    ui.alert(
        'Already Initialized',
        'All required permanent tabs already exist. No changes were made.',
        ui.ButtonSet.OK);
  }
}

/**
 * Triggers the pairing process for Team Policy (TP) debates.
 */
function generateTpMatches() {
  generatePairings(TP_TYPE);
}

/**
 * Triggers the pairing process for Lincoln-Douglas (LD) debates.
 */
function generateLdMatches() {
  generatePairings(LD_TYPE);
}

/**
 * Resets the 'Attending?' column in the Availability tab for all listed
 * participants.
 */
function clearRsvps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availabilitySheet = ss.getSheetByName(AVAILABILITY_TAB);
  if (!availabilitySheet) {
    showAlert(`The '${
        AVAILABILITY_TAB}' tab is missing. Please run "Initialize Sheet" first.`);
    return;
  }
  const lastRow = availabilitySheet.getLastRow();
  if (lastRow < 2) return;  // Nothing to clear

  const participantRange = availabilitySheet.getRange(`A2:A${lastRow}`);
  const rsvpRange = availabilitySheet.getRange(`B2:B${lastRow}`);

  const participants = participantRange.getValues();
  const rsvps = rsvpRange.getValues();

  for (let i = 0; i < participants.length; i++) {
    if (participants[i][0] !== '') {
      rsvps[i][0] = RSVP_DEFAULT;
    } else {
      rsvps[i][0] = '';
    }
  }
  rsvpRange.setValues(rsvps);
}


// --- CORE PAIRING LOGIC ---

/**
 * Main engine for generating debate pairings for a given format.
 * @param {string} debateType The debate type ('TP' or 'LD').
 */
function generatePairings(debateType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const todayStr = getTodaysDateString();
  const newSheetName = `${debateType} ${todayStr}`;

  try {
    // 1. Pre-flight checks
    if (ss.getSheetByName(newSheetName)) {
      throw new Error(`A sheet named '${
          newSheetName}' already exists for today. Please delete it to regenerate.`);
    }
    runDataIntegrityChecks();

    // 2. Get data
    const available = getAvailableParticipants(debateType);
    if (available.debaters.length < 2) {
      throw new Error(`Not enough available debaters for type '${
          debateType}' to create any matches.`);
    }

    const history = getMatchHistory();
    let allPairings = [];

    if (debateType === LD_TYPE) {
      // --- FIXED: Explicit LD Two-Round Logic with correct BYE handling ---
      let round1Bye = null;
      let round2Bye = null;
      const historyForR2 =
          JSON.parse(JSON.stringify(history));  // Deep copy for R2

      // --- Round 1 Generation ---
      let r1Debaters = [...available.debaters];
      r1Debaters.sort(
          (a, b) => (history.byes[a.Name] || 0) - (history.byes[b.Name] || 0));

      if (r1Debaters.length % 2 !== 0) {
        round1Bye = r1Debaters.shift();  // Debater with fewest historical BYEs
                                         // gets the R1 BYE
      }

      const round1Pairings = generateRoundPairings(
          r1Debaters, available.judges, available.rooms, history, debateType);
      updateHistoryInMemory(
          historyForR2, round1Pairings, round1Bye ? round1Bye.Name : null);

      if (round1Bye) {
        round1Pairings.push(
            {Aff: round1Bye.Name, Neg: BYE_TEAM, Judge: '', Room: ''});
      }
      allPairings.push(...round1Pairings.map(p => ({Round: 1, ...p})));

      // --- Round 2 Generation ---
      let r2Debaters = [...available.debaters];
      // Sort by the *updated* history to find the next person for a BYE
      r2Debaters.sort(
          (a, b) => (historyForR2.byes[a.Name] || 0) -
              (historyForR2.byes[b.Name] || 0));

      if (r2Debaters.length % 2 !== 0) {
        round2Bye = r2Debaters.shift();  // Next person in line gets the R2 BYE
      }

      const round2Pairings = generateRoundPairings(
          r2Debaters, available.judges, available.rooms, historyForR2,
          debateType);

      if (round2Bye) {
        round2Pairings.push(
            {Aff: round2Bye.Name, Neg: BYE_TEAM, Judge: '', Room: ''});
      }
      allPairings.push(...round2Pairings.map(p => ({Round: 2, ...p})));

    } else {
      // --- TP (or other single-round) Logic ---
      const {pairings, bye} = generateRoundPairings(
          available.debaters, available.judges, available.rooms, history,
          debateType);
      allPairings.push(...pairings);
      if (bye) {
        allPairings.push({Aff: bye, Neg: BYE_TEAM, Judge: '', Room: ''});
      }
    }

    // 3. Write to sheet
    writePairingsToSheet(allPairings, debateType, newSheetName);
    sortAllTabs();
    applyAllConditionalFormats();  // Apply validation to the new sheet
    ss.getSheetByName(newSheetName).activate();

  } catch (e) {
    showAlert(`Pairing Failed: ${e.message}`);
  }
}

/**
 * Creates a single, complete set of pairings for a given pool of participants.
 * Assumes an even number of debaters for LD. Handles odd numbers/BYEs for TP.
 * @param {Array<object>} debaters - The pool of debaters for this round.
 * @param {Array<object>} judges - The pool of available judges.
 * @param {Array<object>} rooms - The pool of available rooms.
 * @param {object} history - The historical match data.
 * @param {string} debateType - The debate format ('TP' or 'LD').
 * @returns {object} For TP: { pairings: Array, bye: string|null }. For LD: An
 *     array of pairings.
 */
function generateRoundPairings(debaters, judges, rooms, history, debateType) {
  let availableDebaters = [...debaters];
  let availableJudges = [...judges];
  let availableRooms = [...rooms];

  let pairings = [];
  let bye = null;

  if (debateType === TP_TYPE) {
    // Consolidate TP debaters into teams or Ironman
    let teams = {};
    availableDebaters.forEach(d => {
      if (d.Partner) {
        if (!teams[d.Partner]) {  // If partner isn't in teams yet
          teams[d.Name] = {members: [d], hardMode: d['Hard Mode']};
        } else {
          teams[d.Partner].members.push(d);
        }
      } else {  // Ironman
        teams[d.Name] = {members: [d], hardMode: d['Hard Mode']};
      }
    });

    availableDebaters = Object.entries(teams).map(([key, value]) => {
      const name = value.members.length > 1 ?
          `${value.members[0].Name} / ${value.members[1].Name}` :
          `${value.members[0].Name}${IRONMAN_SUFFIX}`;
      return {
        Name: name,
        'Hard Mode': value.hardMode,
        isTeam: true,
        rawNames: value.members.map(m => m.Name)
      };
    });

    // Assign BYE if necessary for TP
    if (availableDebaters.length % 2 !== 0) {
      availableDebaters.sort(
          (a, b) => (history.byes[a.Name] || 0) - (history.byes[b.Name] || 0));
      bye = availableDebaters.shift().Name;
    }
  }

  // Check for sufficient resources for the matches that will be made
  const requiredMatches = Math.floor(availableDebaters.length / 2);
  if (availableJudges.length < requiredMatches)
    throw new Error(`Insufficient judges. Need ${requiredMatches}, have ${
        availableJudges.length}.`);
  if (availableRooms.length < requiredMatches)
    throw new Error(`Insufficient rooms. Need ${requiredMatches}, have ${
        availableRooms.length}.`);

  // Main pairing loop
  let unassignedDebaters = shuffleArray(availableDebaters);
  while (unassignedDebaters.length > 1) {  // Stop when 1 or 0 are left
    const aff = unassignedDebaters.shift();
    const bestMatch = findBestMatch(
        aff, unassignedDebaters, availableJudges, availableRooms, history);

    if (!bestMatch.neg) {
      throw new Error(`Could not find a valid opponent for ${
          aff.Name}. Please check for restrictive judge conflicts or roster issues.`);
    }

    pairings.push({
      Aff: aff.Name,
      Neg: bestMatch.neg.Name,
      Judge: bestMatch.judge.Name,
      Room: bestMatch.room['Room Name']
    });

    // Remove used participants from pools
    unassignedDebaters =
        unassignedDebaters.filter(d => d.Name !== bestMatch.neg.Name);
    availableJudges =
        availableJudges.filter(j => j.Name !== bestMatch.judge.Name);
    availableRooms = availableRooms.filter(
        r => r['Room Name'] !== bestMatch.room['Room Name']);
  }

  return debateType === TP_TYPE ? {pairings, bye} : pairings;
}


/**
 * Finds the best opponent, judge, and room for a given debater (Aff).
 * @param {object} aff - The debater/team to be paired.
 * @param {Array<object>} opponents - The pool of potential opponents.
 * @param {Array<object>} judges - The pool of available judges.
 * @param {Array<object>} rooms - The pool of available rooms.
 * @param {object} history - The historical match data.
 * @returns {object} An object containing the best `neg`, `judge`, and `room`.
 */
function findBestMatch(aff, opponents, judges, rooms, history) {
  let bestMatch = {neg: null, judge: null, room: null, score: -Infinity};

  const affRawNames = aff.isTeam ? aff.rawNames : [aff.Name];

  for (const neg of opponents) {
    const negRawNames = neg.isTeam ? neg.rawNames : [neg.Name];

    for (const judge of judges) {
      // Hard Constraint: Judge-child conflict
      const judgeChildren = judge['Children\'s Names'] ?
          judge['Children\'s Names'].split(',').map(c => c.trim()) :
          [];
      const allDebatersInMatch = [...affRawNames, ...negRawNames];
      if (judgeChildren.some(child => allDebatersInMatch.includes(child))) {
        continue;  // Skip this judge
      }

      for (const room of rooms) {
        const score = calculateMatchScore(aff, neg, judge, room, history);
        if (score > bestMatch.score) {
          bestMatch = {neg, judge, room, score};
        }
      }
    }
  }
  return bestMatch;
}

/**
 * Calculates a score for a potential matchup based on soft constraints.
 * Higher scores are better.
 * @param {object} aff - The Aff debater/team.
 * @param {object} neg - The Neg debater/team.
 * @param {object} judge - The potential judge.
 * @param {object} room - The potential room.
 * @param {object} history - The historical match data.
 * @returns {number} The calculated score.
 */
function calculateMatchScore(aff, neg, judge, room, history) {
  let score = 0;
  const affHistory = history.opponents[aff.Name] || [];
  const negHistory = history.opponents[neg.Name] || [];
  const affJudgeHistory = history.judges[aff.Name] || [];
  const judgeRoomHistory = history.rooms[judge.Name] || [];

  // Priority 1: Pair by Hard Mode status
  if (aff['Hard Mode'] === neg['Hard Mode']) {
    score += 100;
  }

  // Priority 2: Minimize rematches
  if (affHistory.includes(neg.Name) || negHistory.includes(aff.Name)) {
    score -= 50;
  }

  // Priority 3: Minimize re-judging
  if (affJudgeHistory.includes(judge.Name)) {
    score -= 10;
  }

  // Priority 4: Judge-room preference
  if (judgeRoomHistory.includes(room['Room Name'])) {
    score += 5;
  }

  // Add a small random factor to break ties
  score += Math.random();

  return score;
}

/**
 * Updates a history object in-memory with results from a round.
 * @param {object} history - The history object to modify.
 * @param {Array<object>} pairings - The list of pairings from the round.
 * @param {string|null} byeName - The name of the debater/team with a bye in
 *     this round.
 */
function updateHistoryInMemory(history, pairings, byeName) {
  if (byeName) {
    history.byes[byeName] = (history.byes[byeName] || 0) + 1;
  }

  pairings.forEach(p => {
    if (p.Neg === BYE_TEAM) return;  // Skip BYE pairings already handled
    // Update opponents
    if (!history.opponents[p.Aff]) history.opponents[p.Aff] = [];
    if (!history.opponents[p.Neg]) history.opponents[p.Neg] = [];
    history.opponents[p.Aff].push(p.Neg);
    history.opponents[p.Neg].push(p.Aff);

    // Update judges for each debater involved
    const debatersInMatch = [p.Aff, p.Neg].flatMap(
        side => side.includes('/') ? side.split('/').map(n => n.trim()) :
                                     [side.replace(IRONMAN_SUFFIX, '').trim()]);
    debatersInMatch.forEach(debaterName => {
      if (!history.judges[debaterName]) history.judges[debaterName] = [];
      history.judges[debaterName].push(p.Judge);
    });
  });
}


// --- DATA RETRIEVAL ---

/**
 * Gets all data from a sheet as an array of objects.
 * @param {string} sheetName - The name of the sheet.
 * @returns {Array<Object>} An array of objects, where keys are header names.
 */
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];

  const headers = values.shift().map(h => h.trim());
  return values
      .map(row => {
        let obj = {};
        headers.forEach((header, i) => {
          obj[header] = row[i];
        });
        return obj;
      })
      .filter(obj => obj[headers[0]] !== '');  // Filter out empty rows
}

/**
 * Retrieves lists of available debaters, judges, and rooms for a specific
 * debate type.
 * @param {string} debateType - The debate format ('TP' or 'LD').
 * @returns {object} An object containing arrays of available participants.
 */
function getAvailableParticipants(debateType) {
  const availability =
      getSheetData(AVAILABILITY_TAB).filter(p => p['Attending?'] === RSVP_YES);
  const availableNames = new Set(availability.map(p => p.Participant));

  const debaters = getSheetData(DEBATERS_TAB)
                       .filter(
                           d => d['Debate Type'] === debateType &&
                               availableNames.has(d.Name));

  const judges =
      getSheetData(JUDGES_TAB)
          .filter(
              j => (j['Debate Type'] === debateType || !j['Debate Type']) &&
                  availableNames.has(
                      j.Name));  // Allow judges with no type specified

  const rooms =
      getSheetData(ROOMS_TAB).filter(r => r['Debate Type'] === debateType);

  return {debaters, judges, rooms};
}

/**
 * Parses all historical match sheets to build a record of opponents, judges,
 * and byes.
 * @returns {object} An object containing `opponents`, `judges`, `byes`, and
 *     `rooms` history.
 */
function getMatchHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const history = {opponents: {}, judges: {}, byes: {}, rooms: {}};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const isTpMatch = sheetName.startsWith(TP_TYPE + ' ');
    const isLdMatch = sheetName.startsWith(LD_TYPE + ' ');

    if (isTpMatch || isLdMatch) {
      const data = getSheetData(sheetName);
      data.forEach(match => {
        const {Aff, Neg, Judge, Room} = match;
        if (!Aff || !Neg) return;

        if (Neg === BYE_TEAM) {
          const byeName = Aff.replace(IRONMAN_SUFFIX, '').trim();
          history.byes[byeName] = (history.byes[byeName] || 0) + 1;
        } else {
          // Opponent History for teams/individuals
          const affSide = Aff;
          const negSide = Neg;
          if (!history.opponents[affSide]) history.opponents[affSide] = [];
          history.opponents[affSide].push(negSide);
          if (!history.opponents[negSide]) history.opponents[negSide] = [];
          history.opponents[negSide].push(affSide);

          // Judge History for each individual debater
          const allDebaters = [];
          [Aff, Neg].forEach(side => {
            if (side.includes('/')) {
              allDebaters.push(...side.split('/').map(n => n.trim()));
            } else {
              allDebaters.push(side.replace(IRONMAN_SUFFIX, '').trim());
            }
          });
          allDebaters.forEach(debater => {
            if (!history.judges[debater]) history.judges[debater] = [];
            history.judges[debater].push(Judge);
          });

          // Judge Room History
          if (Judge && Room) {
            if (!history.rooms[Judge]) history.rooms[Judge] = [];
            history.rooms[Judge].push(Room);
          }
        }
      });
    }
  });
  return history;
}


// --- DATA VALIDATION & FORMATTING ---

/**
 * Runs all data integrity checks before generating pairings. Throws an error on
 * failure.
 */
function runDataIntegrityChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allDebaters = new Set(getSheetData(DEBATERS_TAB).map(d => d.Name));
  const allJudges = new Set(getSheetData(JUDGES_TAB).map(j => j.Name));
  const allRooms = getSheetData(ROOMS_TAB);

  // Check: Room names are unique
  const roomNames = allRooms.map(r => r['Room Name']);
  if (new Set(roomNames).size !== roomNames.length) {
    throw new Error(`Duplicate room name found on the '${
        ROOMS_TAB}' tab. Please ensure all room names are unique.`);
  }

  // Check: No person is both a debater and a judge
  const intersection = new Set([...allDebaters].filter(x => allJudges.has(x)));
  if (intersection.size > 0) {
    throw new Error(
        `The following people are listed as both a Debater and a Judge: ${
                [...intersection].join(
                    ', ')}. Please remove them from one list.`);
  }

  // Check: Judge's children exist in Debaters roster
  getSheetData(JUDGES_TAB).forEach(judge => {
    if (judge['Children\'s Names']) {
      const children = judge['Children\'s Names'].split(',').map(c => c.trim());
      children.forEach(child => {
        if (child && !allDebaters.has(child)) {
          throw new Error(`On the '${JUDGES_TAB}' tab, judge '${
              judge.Name}' has a child listed ('${child}') who is not in the '${
              DEBATERS_TAB}' roster.`);
        }
      });
    }
  });

  // Check: TP Partnership Consistency
  const debatersData = getSheetData(DEBATERS_TAB);
  const debaterMap = new Map(debatersData.map(d => [d.Name, d]));
  debatersData.forEach(debater => {
    if (debater['Debate Type'] === TP_TYPE && debater.Partner) {
      const partner = debaterMap.get(debater.Partner);
      if (!partner) {
        throw new Error(
            `On '${DEBATERS_TAB}', debater '${debater.Name}' lists partner '${
                debater.Partner}' who does not exist in the roster.`);
      }
      if (partner.Partner !== debater.Name) {
        throw new Error(`On '${DEBATERS_TAB}', partnership for '${
            debater.Name}' and '${partner.Name}' is not reciprocal.`);
      }
      if (partner['Hard Mode'] !== debater['Hard Mode']) {
        throw new Error(
            `On '${DEBATERS_TAB}', partners '${debater.Name}' and '${
                partner.Name}' have different 'Hard Mode' settings.`);
      }
    }
  });
}

/**
 * Applies all conditional formatting rules to the relevant sheets.
 * This is designed to be run on open to ensure rules are always in place.
 */
function applyAllConditionalFormats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Debaters Tab
  const debatersSheet = ss.getSheetByName(DEBATERS_TAB);
  if (debatersSheet) {
    const rules = [
      // Partner exists
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(C2<>"", ISNA(MATCH(C2, INDIRECT("${
              DEBATERS_TAB}!$A$2:$A"), 0)))`)
          .setRanges([debatersSheet.getRange('C2:C')])
          .setBackground(INVALID_ROW_COLOR)
          .build(),
      // Partnership is reciprocal
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(C2<>"", VLOOKUP(C2, INDIRECT("${
              DEBATERS_TAB}!$A$2:$C"), 3, FALSE)<>A2)`)
          .setRanges([debatersSheet.getRange('A2:A')])
          .setBackground(INVALID_ROW_COLOR)
          .build(),
      // Partners have same hard mode
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(C2<>"", VLOOKUP(C2, INDIRECT("${
              DEBATERS_TAB}!$A$2:$D"), 4, FALSE)<>D2)`)
          .setRanges([debatersSheet.getRange('D2:D')])
          .setBackground(INVALID_ROW_COLOR)
          .build()
    ];
    debatersSheet.setConditionalFormatRules(rules);
  }

  // Judges Tab
  const judgesSheet = ss.getSheetByName(JUDGES_TAB);
  if (judgesSheet) {
    const rules = [
      // Judge is not also a debater
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(
              `=COUNTIF(INDIRECT("${DEBATERS_TAB}!$A$2:$A"), A2)>0`)
          .setRanges([judgesSheet.getRange('A2:A')])
          .setBackground(INVALID_ROW_COLOR)
          .build(),
      // Children exist in debater list (using a custom function for
      // reliability)
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(B2<>"", HAS_INVALID_CHILD(B2, INDIRECT("${
              DEBATERS_TAB}!$A$2:$A")))`)
          .setRanges([judgesSheet.getRange('B2:B')])
          .setBackground(INVALID_ROW_COLOR)
          .build()
    ];
    judgesSheet.setConditionalFormatRules(rules);
  }

  // Rooms Tab
  const roomsSheet = ss.getSheetByName(ROOMS_TAB);
  if (roomsSheet) {
    const rules = [
      // Room name is unique
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=COUNTIF($A$2:$A, A2)>1`)
          .setRanges([roomsSheet.getRange('A2:A')])
          .setBackground(INVALID_ROW_COLOR)
          .build()
    ];
    roomsSheet.setConditionalFormatRules(rules);
  }

  // Availability Tab
  const availabilitySheet = ss.getSheetByName(AVAILABILITY_TAB);
  if (availabilitySheet) {
    const rules = [
      // Highlight blank 'Attending?' cells
      SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(A2<>"", B2="")`)
          .setRanges([availabilitySheet.getRange('B2:B')])
          .setBackground(INVALID_ROW_COLOR)
          .build()
    ];
    availabilitySheet.setConditionalFormatRules(rules);
  }
}

// --- CUSTOM SHEET FUNCTIONS ---

/**
 * Checks if a comma-separated list of names contains any names not in the
 * Debaters roster. This is used for conditional formatting in the 'Judges'
 * sheet.
 * @param {string} childrenString The comma-separated string from the
 *     "Children's Names" column.
 * @param {any[][]} debaterNamesRange The range of names from the 'Debaters' tab
 *     (e.g., A2:A).
 * @returns {boolean} TRUE if an invalid child name is found, otherwise FALSE.
 * @customfunction
 */
function HAS_INVALID_CHILD(childrenString, debaterNamesRange) {
  // This function is designed to be called from a spreadsheet formula.
  if (!childrenString || typeof childrenString !== 'string' ||
      childrenString.trim() === '') {
    return false;
  }

  // Convert the 2D array from the sheet into a Set for fast lookups.
  const validNames = new Set(debaterNamesRange.flat().filter(String));

  if (validNames.size === 0) return true;

  const children = childrenString.split(',').map(name => name.trim());

  // Check if any listed child is not in the set of valid debater names.
  for (const child of children) {
    if (child && !validNames.has(child)) {
      return true;  // Found a child who is not in the valid names set.
    }
  }

  return false;  // All children found were valid.
}


// --- OUTPUT AND UTILITIES ---

/**
 * Writes the generated pairings to a new, formatted sheet.
 * @param {Array<Object>} pairings - The list of pairings to write.
 * @param {string} debateType - The debate format ('TP' or 'LD').
 * @param {string} newSheetName - The name for the new sheet.
 */
function writePairingsToSheet(pairings, debateType, newSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = debateType === LD_TYPE ? LD_MATCH_HEADERS : TP_MATCH_HEADERS;
  const sheet = ss.insertSheet(newSheetName, PERMANENT_TABS.length);

  // Set headers and formatting
  sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (pairings.length === 0) return;

  // Prepare data for writing
  const outputData = pairings.map(p => headers.map(h => p[h] || ''));

  // Sort LD data
  if (debateType === LD_TYPE) {
    outputData.sort((a, b) => {
      const roundA = a[0];
      const roundB = b[0];  // Round
      const roomA = a[4];
      const roomB = b[4];  // Room
      if (roundA !== roundB) return roundA - roundB;
      // Sort BYE to the bottom
      if (a[1] === BYE_TEAM || b[1] === BYE_TEAM)
        return a[1] === BYE_TEAM ? 1 : -1;
      if (a[2] === BYE_TEAM || b[2] === BYE_TEAM)
        return a[2] === BYE_TEAM ? 1 : -1;
      return roomA.localeCompare(roomB);
    });
  }

  // Write data
  sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Sorts all spreadsheet tabs, keeping permanent tabs first, followed by
 * generated match tabs sorted by date (newest first), then type (LD before TP).
 */
function sortAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetData =
      allSheets.map((s, i) => ({sheet: s, name: s.getName(), index: i}));

  const permanentSheetNames = new Set(PERMANENT_TABS);

  const generatedSheets =
      sheetData.filter(s => !permanentSheetNames.has(s.name));

  generatedSheets.sort((a, b) => {
    // Regex to extract type and date
    const regex = /^(TP|LD) (\d{4}-\d{2}-\d{2})$/;
    const matchA = a.name.match(regex);
    const matchB = b.name.match(regex);

    if (!matchA || !matchB) return 0;  // Don't sort non-standard sheets

    const typeA = matchA[1];
    const dateA = matchA[2];
    const typeB = matchB[1];
    const dateB = matchB[2];

    // Sort by date descending
    if (dateA !== dateB) return dateB.localeCompare(dateA);
    // Then by type (LD comes before TP)
    if (typeA !== typeB) return typeA === LD_TYPE ? -1 : 1;

    return 0;
  });

  // Reorder sheets
  generatedSheets.forEach((s, i) => {
    ss.setActiveSheet(s.sheet);
    ss.moveActiveSheet(PERMANENT_TABS.length + i + 1);
  });

  // Ensure permanent tabs are in the correct order at the beginning
  PERMANENT_TABS.reverse().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(1);
    }
  });

  ss.getSheetByName(AVAILABILITY_TAB).activate();
}

/**
 * Displays a UI alert with a specified message.
 * @param {string} message The message to display.
 */
function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Returns the current date as a 'YYYY-MM-DD' formatted string.
 * @returns {string} The formatted date string.
 */
function getTodaysDateString() {
  return Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Shuffles an array in place using the Fisher-Yates algorithm.
 * @param {Array} array The array to shuffle.
 * @returns {Array} The shuffled array.
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}


// --- SAMPLE DATA PROVIDERS ---

/**
 * Provides sample data for the Debaters tab.
 * @returns {Array<Array<string>>} 2D array of debater data.
 */
function getSampleDebaters() {
  return [
    ['Abraham Lincoln', 'LD', '', 'No'],
    ['Stephen A. Douglas', 'LD', '', 'No'],
    ['Clarence Darrow', 'LD', '', 'No'],
    ['William Jennings Bryan', 'LD', '', 'No'],
    ['William F. Buckley Jr.', 'LD', '', 'Yes'],
    ['Gore Vidal', 'LD', '', 'Yes'],
    ['Christopher Hitchens', 'LD', '', 'Yes'],
    ['Tony Blair', 'LD', '', 'Yes'],
    ['Jordan Peterson', 'LD', '', 'No'],
    ['Slavoj Žižek', 'LD', '', 'No'],
    ['Lloyd Bentsen', 'LD', '', 'No'],
    ['Dan Quayle', 'LD', '', 'No'],
    ['Richard Dawkins', 'LD', '', 'Yes'],
    ['Rowan Williams', 'LD', '', 'No'],
    ['Diogenes', 'LD', '', 'No'],
    ['Noam Chomsky', 'TP', 'Michel Foucault', 'Yes'],
    ['Michel Foucault', 'TP', 'Noam Chomsky', 'Yes'],
    ['Harlow Shapley', 'TP', 'Heber Curtis', 'No'],
    ['Heber Curtis', 'TP', 'Harlow Shapley', 'No'],
    ['Muhammad Ali', 'TP', 'George Foreman', 'No'],
    ['George Foreman', 'TP', 'Muhammad Ali', 'No'],
    ['Richard Nixon', 'TP', 'Nikita Khrushchev', 'No'],
    ['Nikita Khrushchev', 'TP', 'Richard Nixon', 'No'],
    ['Thomas Henry Huxley', 'TP', 'Samuel Wilberforce', 'Yes'],
    ['Samuel Wilberforce', 'TP', 'Thomas Henry Huxley', 'Yes'],
    ['John F. Kennedy', 'TP', 'David Frost', 'No'],
    ['David Frost', 'TP', 'John F. Kennedy', 'No'],
    ['Bob Dole', 'TP', 'Bill Clinton', 'No'],
    ['Bill Clinton', 'TP', 'Bob Dole', 'No']
  ];
}

/**
 * Provides sample data for the Judges tab.
 * @returns {Array<Array<string>>} 2D array of judge data.
 */
function getSampleJudges() {
  return [
    ['Howard K. Smith', 'John F. Kennedy', 'TP'],
    ['Fons Elders', 'Noam Chomsky, Michel Foucault', 'TP'],
    ['John Stevens Henslow', 'Samuel Wilberforce', 'TP'],
    ['Jim Lehrer', 'Bill Clinton', 'TP'], ['Judy Woodruff', '', 'TP'],
    ['Tom Brokaw', '', 'TP'], ['Frank McGee', '', 'TP'],
    ['Quincy Howe', '', 'TP'], ['John T. Raulston', 'Clarence Darrow', 'LD'],
    ['Rudyard Griffiths', 'Jordan Peterson', 'LD'],
    ['Stephen J. Blackwood', '', 'LD'], ['Brit Hume', '', 'LD'],
    ['Jon Margolis', '', 'LD'], ['Bill Shadel', '', 'LD'],
    ['Judge Judy', '', 'LD']
  ];
}

/**
 * Provides sample data for the Rooms tab.
 * @returns {Array<Array<string>>} 2D array of room data.
 */
function getSampleRooms() {
  return [
    ['Room 101', 'LD'], ['Room 102', 'LD'], ['Room 103', 'LD'],
    ['Room 201', 'LD'], ['Sanctuary right', 'LD'], ['Sanctuary left', 'LD'],
    ['Pantry', 'LD'], ['Chapel', 'TP'], ['Library', 'TP'],
    ['Music lounge', 'TP'], ['Cry room', 'TP'], ['Office', 'TP'],
    ['Office hallway', 'TP']
  ];
}
