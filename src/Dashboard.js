/**
 * Executive Board High-Level KPI Dashboard for AYSO Region 154.
 */

function generateExecutiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  if (!refSheet || refSheet.getLastRow() <= 1) return;

  const rawRows = refSheet.getDataRange().getValues();
  const games = transformRawSchedule(rawRows);
  const times = getPipelineTimestamps();

  const totalHomeGames = games.length;
  const uniqueFields = [...new Set(games.map(g => g.cleanField))];
  const assignedRefGames = games.filter(g => g.centerRef && g.centerRef !== "[Open]").length;
  const openRefGames = totalHomeGames - assignedRefGames;
  const refCoveragePct = totalHomeGames > 0 ? Math.round((assignedRefGames / totalHomeGames) * 100) : 0;

  const now = new Date();
  const todayTimestamp = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const twoWeeksFuture = todayTimestamp + (14 * 24 * 60 * 60 * 1000);
  const upcomingGamesCount = games.filter(g => g.timestamp >= todayTimestamp && g.timestamp <= twoWeeksFuture).length;

  const headerMap = buildHeaderMap(rawRows[0]);
  let homeScoreCol = -1, awayScoreCol = -1;
  rawRows[0].forEach((h, idx) => {
    const col = String(h).trim().toLowerCase();
    if (col.includes("home score")) homeScoreCol = idx;
    if (col.includes("away score")) awayScoreCol = idx;
  });

  const missingByDiv = {};
  let totalMissingScores = 0;

  for (let i = 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!isE154HomeGame(row, headerMap)) continue;

    const rawDate = row[headerMap.date];
    const matchTimestamp = parseDateToTimestamp(rawDate);
    if (matchTimestamp > 0 && matchTimestamp < todayTimestamp) {
      const hs = homeScoreCol > -1 ? String(row[homeScoreCol] || "").trim() : "";
      const as = awayScoreCol > -1 ? String(row[awayScoreCol] || "").trim() : "";
      if (hs === "" || as === "") {
        totalMissingScores++;
        const div = String(row[headerMap.division] || "").trim();
        missingByDiv[div] = (missingByDiv[div] || 0) + 1;
      }
    }
  }

  let dashSheet = ss.getSheetByName(CONFIG.SHEET_TABS.DASHBOARD);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(CONFIG.SHEET_TABS.DASHBOARD, 0);
  } else {
    ss.setActiveSheet(dashSheet);
    ss.moveActiveSheet(1);
  }
  dashSheet.clear();

  // Universal Top Banner
  dashSheet.getRange(1, 1).setValue(
    `⚡ AYSO 154 EXECUTIVE BOARD OVERVIEW | MatchTrak Schedule Export: ${times.exportTime} | Last Updated in Google Sheets: ${times.sheetTime}`
  );
  dashSheet.getRange(1, 1, 1, 8).merge()
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  dashSheet.setRowHeight(1, 28);

  // Executive KPI Summary Cards
  const kpiHeaders = ["Upcoming Home Games (Next 14 Days)", "Season Home Fixtures", "Active Facilities", "Referee Coverage", "Past Unreported Scores"];
  const kpiValues = [upcomingGamesCount, totalHomeGames, uniqueFields.length, `${refCoveragePct}%`, totalMissingScores];

  dashSheet.getRange(3, 1, 1, 5).setValues([kpiHeaders])
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1b365d")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(9);
  dashSheet.setRowHeight(3, 24);

  dashSheet.getRange(4, 1, 1, 5).setValues([kpiValues])
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#f8fafc")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(16);
  dashSheet.setRowHeight(4, 40);

  // Unreported Scores Breakdown Table
  dashSheet.getRange(6, 1).setValue("📊 UNREPORTED SCORES BREAKDOWN BY DIVISION")
    .setFontWeight("bold")
    .setFontColor("#991b1b")
    .setFontSize(11);

  const scoreSummaryHeaders = ["Division", "Division Director", "Contact Email", "Overdue Matches", "Compliance Status"];
  dashSheet.getRange(7, 1, 1, 5).setValues([scoreSummaryHeaders])
    .setFontWeight("bold")
    .setBackground("#991b1b")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  dashSheet.setRowHeight(7, 26);

  const sortedDivs = Object.keys(missingByDiv).sort((a, b) => getDivisionSortWeight(a) - getDivisionSortWeight(b));
  const scoreRows = sortedDivs.map(div => {
    const dir = getDivisionDirector(div);
    const count = missingByDiv[div];
    const status = count > 5 ? "⚠️ Action Required" : "Follow-Up Needed";
    return [div, dir.name, dir.email, count, status];
  });

  if (scoreRows.length > 0) {
    dashSheet.getRange(8, 1, scoreRows.length, 5).setValues(scoreRows);
    dashSheet.getRange(8, 1, scoreRows.length, 1).setHorizontalAlignment("center");
    dashSheet.getRange(8, 4, scoreRows.length, 2).setHorizontalAlignment("center");
    for (let i = 0; i < scoreRows.length; i++) {
      dashSheet.setRowHeight(8 + i, 24);
      if (i % 2 === 1) dashSheet.getRange(8 + i, 1, 1, 5).setBackground("#fef2f2");
    }
  }

  const widths = [180, 180, 240, 150, 170, 60, 60, 60];
  widths.forEach((w, idx) => dashSheet.setColumnWidth(idx + 1, w));
  dashSheet.setFrozenRows(1);
}
