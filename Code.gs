/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Debate Automation')
      .addItem('Initialize Sheet', 'initializeSheet')
      .addItem('New Debate', 'createNewDebateSheet')
      .addItem('Generate Matches', 'generateMatches')
      .addToUi();
}

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());

  const requiredSheets = {
    "Debaters": [["Name", "Debate Type", "Partner", "Hard Mode"]],
    "Judges": [["Name", "Child"]],
    "Rooms": [["Room Name"]],
    "Availability": [["Name", "Available"]]
  };

  for (const sheetName in requiredSheets) {
    if (sheets.indexOf(sheetName) === -1) {
      const newSheet = ss.insertSheet(sheetName);
      const headers = requiredSheets[sheetName];
      newSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
      // Add sample data
      if (sheetName === "Debaters") {
        newSheet.getRange(2, 1, 2, 4).setValues([
          ["John Doe", "TP", "Jane Smith", "No"],
          ["Alice", "LD", "", "Yes"]
        ]);
      } else if (sheetName === "Judges") {
        newSheet.getRange(2, 1, 2, 2).setValues([
          ["Judge Judy", "John Doe"],
          ["Judge Mathis", ""]
        ]);
      } else if (sheetName === "Rooms") {
        newSheet.getRange(2, 1, 2, 1).setValues([
          ["Room 101"],
          ["Room 102"]
        ]);
      }
    }
  }
}

function createMatches(debaters, judges, rooms, allJudges, ss, isTp = false) {
  const matches = [];
  const matchHistory = getMatchHistory(ss);

  // Helper to form teams, especially for TP
  function formTeams(debaters) {
    const teams = [];
    const debaterPool = [...debaters];
    const pairedDebaters = new Set();

    if (isTp) {
      // Prioritize listed partners
      for (const debater of debaterPool) {
        if (pairedDebaters.has(debater.name)) continue;

        const partnerName = debater.partner;
        if (partnerName) {
          const partnerIndex = debaterPool.findIndex(p => p.name === partnerName && !pairedDebaters.has(p.name));
          if (partnerIndex !== -1) {
            const partner = debaterPool.splice(partnerIndex, 1)[0];
            teams.push([debater, partner]);
            pairedDebaters.add(debater.name);
            pairedDebaters.add(partner.name);
          }
        }
      }
    }

    // Pair remaining debaters
    let unpaired = debaterPool.filter(d => !pairedDebaters.has(d.name));
    if (isTp) {
      while (unpaired.length >= 2) {
        teams.push([unpaired.pop(), unpaired.pop()]);
      }
    } else {
      teams.push(...unpaired.map(d => [d]));
    }

    return teams;
  }

  let teams = formTeams(debaters);

  // Separate teams by hard mode
  const hardModeTeams = teams.filter(t => t.every(d => d.hardMode === 'Yes'));
  const regularTeams = teams.filter(t => t.every(d => d.hardMode !== 'Yes'));

  function pairTeams(teamList) {
    const pairedMatches = [];
    teamList = teamList.sort(() => Math.random() - 0.5); // Shuffle for randomness

    while (teamList.length >= 2) {
      const team1 = teamList.pop();
      let bestMatchIndex = -1;
      let minRematches = Infinity;

      // Find the best team to pair with, avoiding rematches
      for (let i = 0; i < teamList.length; i++) {
        const team2 = teamList[i];
        const team1Name = team1.map(d => d.name).join(" & ");
        const team2Name = team2.map(d => d.name).join(" & ");
        const rematches = (matchHistory[team1Name] || []).filter(t => t === team2Name).length;

        if (rematches < minRematches) {
          minRematches = rematches;
          bestMatchIndex = i;
        }
      }

      const team2 = teamList.splice(bestMatchIndex, 1)[0];
      pairedMatches.push({ team1, team2 });
    }
    return pairedMatches;
  }

  let prelimMatches = pairTeams(hardModeTeams).concat(pairTeams(regularTeams));

  // Combine any remaining teams from both pools
  let remainingTeams = hardModeTeams.concat(regularTeams);
  prelimMatches = prelimMatches.concat(pairTeams(remainingTeams));

  for (const prematch of prelimMatches) {
    const { team1, team2 } = prematch;
    const match = {
      team1: team1.map(d => d.name).join(" & "),
      team2: team2.map(d => d.name).join(" & "),
      judge: null,
      room: null
    };

    // Assign Judge
    let judgeAssigned = false;
    for (let i = 0; i < judges.length; i++) {
      const judge = judges[i];
      const isChildInMatch = team1.some(d => d.name === judge.child) || team2.some(d => d.name === judge.child);
      if (!isChildInMatch) {
        match.judge = judge.name;
        judges.splice(i, 1); // Remove judge from pool
        judgeAssigned = true;
        break;
      }
    }
    if (!judgeAssigned) {
      match.judge = "UNASSIGNED";
    }

    // Assign Room
    if (rooms.length > 0) {
      match.room = rooms.pop();
    } else {
      match.room = "UNASSIGNED";
    }
    matches.push(match);
  }

  return matches;
}

function getMatchHistory(ss) {
  const history = {};
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const team1 = row[3];
        const team2 = row[4];
        if (team1 && team2) {
          if (!history[team1]) history[team1] = [];
          if (!history[team2]) history[team2] = [];
          history[team1].push(team2);
          history[team2].push(team1);
        }
      }
    }
  }
  return history;
}

function createNewDebateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availabilitySheet = ss.getSheetByName("Availability");
  if (!availabilitySheet) {
    SpreadsheetApp.getUi().alert("Availability sheet not found. Please run 'Initialize Sheet' first.");
    return;
  }

  const data = availabilitySheet.getDataRange().getValues();
  const availablePeople = data.filter(row => row[1] === "Yes");

  if (availablePeople.length === 0) {
    SpreadsheetApp.getUi().alert("No one is available for the debate.");
    return;
  }

  const today = new Date();
  const sheetName = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  let newSheet = ss.getSheetByName(sheetName);
  if (newSheet) {
    newSheet.clear();
  } else {
    newSheet = ss.insertSheet(sheetName);
  }

  newSheet.getRange(1, 1, availablePeople.length, availablePeople[0].length).setValues(availablePeople);
  availabilitySheet.getRange(2, 2, availabilitySheet.getLastRow() - 1, 1).clearContent();
}

function generateMatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // Basic validation to ensure we're on a weekly debate sheet
  if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) {
    SpreadsheetApp.getUi().alert("Please run this script from a weekly debate sheet (e.g., 2025-09-15).");
    return;
  }

  // Read all necessary data
  const debatersSheet = ss.getSheetByName("Debaters");
  const judgesSheet = ss.getSheetByName("Judges");
  const roomsSheet = ss.getSheetByName("Rooms");

  if (!debatersSheet || !judgesSheet || !roomsSheet) {
    SpreadsheetApp.getUi().alert("Debaters, Judges, or Rooms sheet not found. Please run 'Initialize Sheet' first.");
    return;
  }

  const availableData = sheet.getDataRange().getValues();
  const allDebaters = debatersSheet.getDataRange().getValues().slice(1);
  const allJudges = judgesSheet.getDataRange().getValues().slice(1);
  const allRooms = roomsSheet.getDataRange().getValues().slice(1).map(row => row[0]);

  const availableDebaters = availableData.map(row => row[0]);
  const availableJudges = allJudges.filter(judge => availableDebaters.includes(judge[0])).map(judge => ({ name: judge[0], child: judge[1] }));
  const availableRooms = allRooms.filter(room => availableDebaters.includes(room));

  const debaterDetails = allDebaters.map(d => ({ name: d[0], type: d[1], partner: d[2], hardMode: d[3] }));

  const ldDebaters = debaterDetails.filter(d => d.type === 'LD' && availableDebaters.includes(d.name));
  const tpDebaters = debaterDetails.filter(d => d.type === 'TP' && availableDebaters.includes(d.name));

  // Placeholder for pairings
  const pairings = {
    "TP": [],
    "LD": []
  };

  // Implement pairing logic
  const ldMatches = createMatches(ldDebaters, availableJudges, availableRooms, allJudges, ss);
  const tpMatches = createMatches(tpDebaters, availableJudges, availableRooms, allJudges, ss, true);

  pairings.LD = ldMatches;
  pairings.TP = tpMatches;

  // Write pairings back to the sheet
  sheet.getRange(1, 4, 1, 4).setValues([["Team 1", "Team 2", "Judge", "Room"]]);
  let row = 2; // Start writing from the second row
  sheet.getRange("D2:G" + sheet.getMaxRows()).clearContent(); // Clear old pairings

  for (const format in pairings) {
    for (const match of pairings[format]) {
      const rowData = [match.team1, match.team2, match.judge, match.room];
      sheet.getRange(row, 4, 1, 4).setValues([rowData]);
      row++;
    }
  }
}
