/**
 * Reconciliation engine for AYSO Region 154 referee operations.
 * Audits MatchTrak referee assignments against Game Day Form check-ins.
 */

/**
 * Main audit routine: Parses check-in responses, matches them against
 * MatchTrak referee rosters by Game #, and generates the audit table.
 */
function runRefereeAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  const formResponseSheet = getFormResponseSheet(ss);

  if (!refSheet) {
    Logger.log("MatchTrak ref data sheet missing. Aborting audit.");
    return;
  }

  if (!formResponseSheet) {
    Logger.log("Form response sheet not found. Aborting audit.");
    return;
  }

  // 1. Index MatchTrak ref assignments by Game ID
  const refData = refSheet.getDataRange().getValues();
  if (refData.length <= 1) return;

  const matchTrakAssignments = buildMatchTrakRefIndex(refData);

  // 2. Index Form check-ins by Game ID
  const responseData = formResponseSheet.getDataRange().getValues();
  const checkIns = buildCheckInIndex(responseData);

  // 3. Reconcile records
  const auditRows = [];
  const headers = [
    "Game #",
    "Match Details",
    "Field",
    "Position",
    "Assigned in MatchTrak",
    "Checked-In Volunteer",
    "Audit Status",
    "Points Eligible"
  ];
  auditRows.push(headers);

  // Iterate over all active MatchTrak games
  Object.keys(matchTrakAssignments).forEach(gameId => {
    const game = matchTrakAssignments[gameId];
    const reported = checkIns[gameId] || [];

    // Evaluate Center Referee
    auditRows.push(evaluatePositionAudit(gameId, game, reported, "Center Referee", game.centerRef));

    // Evaluate Assistant Referee 1
    if (game.ar1 || reported.some(r => r.position.includes("AR1") || r.position.includes("Assistant 1"))) {
      auditRows.push(evaluatePositionAudit(gameId, game, reported, "Assistant Referee 1", game.ar1));
    }

    // Evaluate Assistant Referee 2
    if (game.ar2 || reported.some(r => r.position.includes("AR2") || r.position.includes("Assistant 2"))) {
      auditRows.push(evaluatePositionAudit(gameId, game, reported, "Assistant Referee 2", game.ar2));
    }
  });

  // 4. Output results to Ref_Audit tab
  const auditSheet = ss.getSheetByName(CONFIG.SHEET_TABS.AUDIT) || ss.insertSheet(CONFIG.SHEET_TABS.AUDIT);
  auditSheet.clearContents();
  auditSheet.getRange(1, 1, auditRows.length, auditRows[0].length).setValues(auditRows);

  // Header styling
  auditSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff");

  auditSheet.setFrozenRows(1);
  auditSheet.autoResizeColumns(1, headers.length);

  Logger.log(`Audit complete. Generated ${auditRows.length - 1} audit verification rows.`);
}

/**
 * Builds an index of referee assignments from MatchTrak export rows.
 * Dynamically detects referee column headers if available.
 */
function buildMatchTrakRefIndex(rows) {
  const headers = rows[0].map(h => String(h).trim().toLowerCase());
  
  // Locate column indices
  let gameIdCol = 3; // default col D
  let fieldCol = 8;  // default col I
  let crCol = headers.findIndex(h => h.includes("center") || h === "cr" || h === "referee");
  let ar1Col = headers.findIndex(h => h.includes("ar1") || h.includes("assistant 1"));
  let ar2Col = headers.findIndex(h => h.includes("ar2") || h.includes("assistant 2"));

  const index = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const gameId = String(row[gameIdCol] || "").trim();
    if (!gameId) continue;

    index[gameId] = {
      gameId: gameId,
      division: String(row[4] || "").trim(),
      field: cleanFieldName(row[fieldCol]),
      timeStr: cleanTime(row[2]),
      homeCoach: cleanTeamName(row[9]),
      awayCoach: cleanTeamName(row[11]),
      centerRef: crCol !== -1 ? String(row[crCol] || "").trim() : "",
      ar1: ar1Col !== -1 ? String(row[ar1Col] || "").trim() : "",
      ar2: ar2Col !== -1 ? String(row[ar2Col] || "").trim() : ""
    };
  }

  return index;
}

/**
 * Extracts checked-in volunteer data from Form response submissions.
 */
function buildCheckInIndex(rows) {
  if (rows.length <= 1) return {};

  const index = {};

  // Form Responses standard layout: Col 0: Timestamp, Col 1: Name/Email, Col 2: Game selection, Col 3: Role
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const gameDropdownValue = String(row[2] || "");
    const volunteerName = String(row[1] || "").trim();
    const role = String(row[3] || "Center Referee").trim();

    // Extract Game ID using regex from bracketed [#GameID]
    const match = gameDropdownValue.match(/#(\d+)/);
    if (!match) continue;

    const gameId = match[1];
    if (!index[gameId]) index[gameId] = [];

    index[gameId].push({
      volunteerName: volunteerName,
      position: role
    });
  }

  return index;
}

/**
 * Evaluates match status between assigned official and checked-in volunteer.
 */
function evaluatePositionAudit(gameId, game, reportedList, positionName, assignedName) {
  // Find report for this specific role
  const checkIn = reportedList.find(r => 
    r.position.toLowerCase().includes(positionName.toLowerCase()) ||
    (positionName.includes("Center") && r.position.toLowerCase().includes("center"))
  );

  const checkedInName = checkIn ? checkIn.volunteerName : "";
  let status = "Unreported / No-Show";
  let points = "0";

  const hasAssigned = assignedName && assignedName !== "" && !assignedName.toLowerCase().includes("unassigned");
  const hasCheckedIn = checkedInName && checkedInName !== "";

  if (hasAssigned && hasCheckedIn) {
    if (normalizeName(assignedName) === normalizeName(checkedInName)) {
      status = "Verified Match";
      points = "Yes";
    } else {
      status = "Substitute / Swap";
      points = "Yes (Credit Substitute)";
    }
  } else if (!hasAssigned && hasCheckedIn) {
    status = "Walk-On / Fill-In";
    points = "Yes (Credit Volunteer)";
  } else if (hasAssigned && !hasCheckedIn) {
    status = "Unreported / No-Show";
    points = "No (0 pts)";
  } else {
    status = "Unassigned / Open";
    points = "—";
  }

  return [
    gameId,
    `${game.division}: ${game.homeCoach} vs ${game.awayCoach} (${game.timeStr})`,
    game.field,
    positionName,
    assignedName || "[Open]",
    checkedInName || "[None]",
    status,
    points
  ];
}

/**
 * Normalizes volunteer names for matching (handles "Last, First" vs "First Last").
 */
function normalizeName(name) {
  return String(name)
    .toLowerCase()
    .replace(/[^a-z]/g, "");
}

/**
 * Helper to locate the Form response sheet by prefix or standard name.
 */
function getFormResponseSheet(ss) {
  const sheets = ss.getSheets();
  return sheets.find(s => s.getName().toLowerCase().includes("form responses")) || 
         ss.getSheetByName("Form Responses 1");
}
