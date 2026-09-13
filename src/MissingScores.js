/**
 * Scans schedules for completed matches missing scores.
 * Uses defensive director resolution and universal timestamps.
 */

function getDivisionDirector(divStr) {
  const clean = String(divStr || "").toUpperCase();
  const directors = (typeof CONFIG !== "undefined" && CONFIG.DIVISION_DIRECTORS) ? CONFIG.DIVISION_DIRECTORS : {};
  const fallback = { name: "Division Coordinator", email: "board@ayso154.org" };

  if (clean.includes("05") || clean.includes("5U")) return directors["5U"] || fallback;
  if (clean.includes("06") || clean.includes("6U")) return directors["6U"] || fallback;
  if (clean.includes("08") || clean.includes("8U")) return directors["8U"] || fallback;
  if (clean.includes("10") || clean.includes("10U")) return directors["10U"] || fallback;
  if (clean.includes("12") || clean.includes("12U")) return directors["12U"] || fallback;
  if (clean.includes("14") || clean.includes("14U")) return directors["14U"] || fallback;
  if (clean.includes("16") || clean.includes("16U")) return directors["16U"] || fallback;
  if (clean.includes("19") || clean.includes("19U")) return directors["19U"] || fallback;
  if (clean.includes("EPIC")) return directors["EPIC"] || fallback;
  if (clean.includes("04") || clean.includes("4U")) return directors["4U"] || fallback;
  return fallback;
}

function getDivisionSortWeight(divStr) {
  const div = String(divStr || "").trim().toUpperCase();
  const match = div.match(/([BG])(?:U)?0?(\d+)/i);
  if (!match) return 999;
  const gender = match[1] === "B" ? 1 : 2;
  const age = parseInt(match[2], 10);
  return age * 10 + gender;
}

function extractMatchDateStr(rawDate) {
  if (!rawDate) return "";
  if (rawDate instanceof Date) return Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M/d");
  const str = String(rawDate).trim();
  const match = str.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
  return match ? `${parseInt(match[1], 10)}/${parseInt(match[2], 10)}` : str;
}

function getDismissedGameIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEET_TABS.RESOLVED_LOG);
  const dismissed = new Set();
  if (logSheet && logSheet.getLastRow() > 1) {
    const ids = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
    ids.forEach(r => {
      const val = String(r[0]).trim();
      if (val) dismissed.add(val);
    });
  }
  return dismissed;
}

function generateMissingScoresReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return;

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.MISSING_SCORES) || ss.insertSheet(CONFIG.SHEET_TABS.MISSING_SCORES);
  const dismissedIds = getDismissedGameIds();

  // 1. Preserve state
  const preservedState = {};
  if (targetSheet.getLastRow() > 2) {
    const oldRows = targetSheet.getDataRange().getValues();
    for (let r = 2; r < oldRows.length; r++) {
      const gId = String(oldRows[r][0] || "").trim();
      if (gId && !isNaN(gId)) {
        preservedState[gId] = {
          contacted: oldRows[r][10] === true,
          waived: oldRows[r][11] === true
        };
      }
    }
  }

  // 2. Perform total clear to remove ghost checkboxes
  targetSheet.clear();

  const rawRows = sourceSheet.getDataRange().getValues();
  const headerMap = buildHeaderMap(rawRows[0]);

  let homeScoreCol = -1, awayScoreCol = -1;
  rawRows[0].forEach((h, idx) => {
    const col = String(h).trim().toLowerCase();
    if (col.includes("home score")) homeScoreCol = idx;
    if (col.includes("away score")) awayScoreCol = idx;
  });

  const now = new Date();
  const todayTimestamp = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const divisionBuckets = {};
  let totalMissing = 0;

  for (let i = 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!isE154HomeGame(row, headerMap)) continue;

    const gameId = String(row[headerMap.gameId] || "").trim();
    if (dismissedIds.has(gameId)) continue;

    const division = String(row[headerMap.division] || "").trim();
    if (CONFIG.EXEMPT_DEVELOPMENTAL_DIVISIONS && /^[BG]U0?[4568]/i.test(division)) continue;

    const rawDate = row[headerMap.date];
    const matchTimestamp = parseDateToTimestamp(rawDate);
    
    if (matchTimestamp > 0 && matchTimestamp < todayTimestamp) {
      const homeScore = homeScoreCol > -1 ? String(row[homeScoreCol] || "").trim() : "";
      const awayScore = awayScoreCol > -1 ? String(row[awayScoreCol] || "").trim() : "";

      if (homeScore === "" || awayScore === "") {
        totalMissing++;
        const dateStr = extractMatchDateStr(rawDate);
        const timeStr = cleanTime(row[headerMap.time]);
        const field = cleanFieldName(row[headerMap.field]);
        const homeCoach = cleanTeamName(row[headerMap.home]);
        const awayCoach = cleanTeamName(row[headerMap.away]);
        const daysOverdue = Math.max(1, Math.floor((todayTimestamp - matchTimestamp) / (1000 * 60 * 60 * 24)));

        let missingSide = "Both Missing";
        if (homeScore === "" && awayScore !== "") missingSide = "Home Missing";
        if (homeScore !== "" && awayScore === "") missingSide = "Away Missing";

        if (!divisionBuckets[division]) divisionBuckets[division] = [];

        const prior = preservedState[gameId] || { contacted: false, waived: false };

        divisionBuckets[division].push({
          gameId: gameId,
          dateStr: dateStr,
          timestamp: matchTimestamp,
          timeStr: timeStr,
          division: division,
          field: field,
          matchup: `${homeCoach} vs ${awayCoach}`,
          homeCoach: homeCoach,
          awayCoach: awayCoach,
          missingSide: missingSide,
          urgency: `${daysOverdue} day(s) overdue`,
          contacted: prior.contacted,
          waived: prior.waived
        });
      }
    }
  }

  const times = getPipelineTimestamps();
  const sortedDivisions = Object.keys(divisionBuckets).sort((a, b) => getDivisionSortWeight(a) - getDivisionSortWeight(b));

  targetSheet.getRange(1, 1).setValue(
    `⚠️ UNREPORTED COACH SCORES | MatchTrak Schedule Export: ${times.exportTime} | Last Updated in Google Sheets: ${times.sheetTime} | Total: ${totalMissing} Pending Across ${sortedDivisions.length} Active Divisions`
  );
  targetSheet.getRange(1, 1, 1, 12).merge()
    .setFontWeight("bold")
    .setFontColor("#991b1b")
    .setBackground("#fee2e2")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(1, 28);

  const headers = [
    "Game #", "Date", "Time", "Division", "Field", 
    "Matchup", "Home Coach", "Away Coach", "Missing Side", 
    "Urgency", "Director Followed Up?", "Waive / Exempt?"
  ];
  targetSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  targetSheet.getRange(2, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#991b1b")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(2, 30);

  let currentRow = 3;

  sortedDivisions.forEach(divName => {
    const divGames = divisionBuckets[divName];
    divGames.sort((a, b) => b.timestamp - a.timestamp);
    const director = getDivisionDirector(divName);

    targetSheet.getRange(currentRow, 1).setValue(
      `⚽ DIVISION ${divName}  |  Director: ${director.name}  ✉️ ${director.email}  (${divGames.length} Matches Unreported)`
    );
    targetSheet.getRange(currentRow, 1, 1, headers.length).merge()
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setBackground("#1e293b")
      .setFontSize(10)
      .setVerticalAlignment("middle");
    targetSheet.setRowHeight(currentRow, 26);
    currentRow++;

    const rows = divGames.map(g => [
      g.gameId, g.dateStr, g.timeStr, g.division, g.field,
      g.matchup, g.homeCoach, g.awayCoach, g.missingSide, g.urgency, 
      g.contacted, g.waived
    ]);

    const dataRange = targetSheet.getRange(currentRow, 1, rows.length, headers.length);
    dataRange.setValues(rows);

    targetSheet.getRange(currentRow, 11, rows.length, 2).insertCheckboxes();
    targetSheet.getRange(currentRow, 1, rows.length, 4).setHorizontalAlignment("center");
    targetSheet.getRange(currentRow, 9, rows.length, 4).setHorizontalAlignment("center");
    dataRange.setVerticalAlignment("middle");

    for (let i = 0; i < rows.length; i++) {
      const rIndex = currentRow + i;
      targetSheet.setRowHeight(rIndex, 25);
      if (i % 2 === 1) targetSheet.getRange(rIndex, 1, 1, headers.length).setBackground("#f8fafc");
    }

    currentRow += rows.length;
  });

  const widths = [85, 75, 95, 80, 180, 230, 130, 130, 120, 125, 155, 130];
  widths.forEach((w, idx) => targetSheet.setColumnWidth(idx + 1, w));

  targetSheet.setFrozenRows(2);
}

function archiveResolvedMissingScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(CONFIG.SHEET_TABS.MISSING_SCORES);
  if (!source || source.getLastRow() <= 2) return;

  const logSheet = ss.getSheetByName(CONFIG.SHEET_TABS.RESOLVED_LOG) || ss.insertSheet(CONFIG.SHEET_TABS.RESOLVED_LOG);
  if (logSheet.getLastRow() === 0) {
    const logHeaders = ["Game #", "Date", "Division", "Matchup", "Resolved Reason", "Archived Timestamp"];
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders])
      .setFontWeight("bold")
      .setBackground("#1b365d")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    logSheet.setFrozenRows(1);
  }

  const rows = source.getDataRange().getValues();
  const toArchive = [];
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a");

  for (let r = 2; r < rows.length; r++) {
    const gameId = String(rows[r][0] || "").trim();
    if (!gameId || isNaN(gameId)) continue;

    const followedUp = rows[r][10] === true;
    const waived = rows[r][11] === true;

    if (followedUp || waived) {
      const reason = waived ? "Exempt / Waived" : "Director Followed Up";
      toArchive.push([gameId, rows[r][1], rows[r][3], rows[r][5], reason, timestamp]);
    }
  }

  if (toArchive.length === 0) {
    SpreadsheetApp.getUi().alert("No checked rows found to dismiss. Check 'Director Followed Up?' or 'Waive / Exempt?' first.");
    return;
  }

  logSheet.getRange(logSheet.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
  logSheet.autoResizeColumns(1, 6);

  generateMissingScoresReport();
  generateExecutiveDashboard();
  ss.toast(`Archived ${toArchive.length} resolved matches to ${CONFIG.SHEET_TABS.RESOLVED_LOG}!`, "Dismissed", 4);
}

function resetAllMissingScoreCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TABS.MISSING_SCORES);
  if (!sheet || sheet.getLastRow() <= 2) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Follow-Up Checkboxes",
    "Are you sure you want to uncheck all 'Director Followed Up' boxes across all divisions?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(3, 11, lastRow - 2, 1);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (typeof values[i][0] === "boolean") values[i][0] = false;
  }

  range.setValues(values);
  ss.toast("All follow-up checkboxes have been reset.", "Checkboxes Reset", 3);
}
