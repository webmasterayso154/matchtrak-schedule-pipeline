/**
 * Referee Audit Engine with universal banner timestamps.
 */

function runRefereeAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  if (!refSheet || refSheet.getLastRow() <= 1) return;

  const rawData = refSheet.getDataRange().getValues();
  const games = transformRawSchedule(rawData);
  const responseSheet = ss.getSheets().find(s => s.getName().toLowerCase().includes("form responses"));
  
  const checkIns = {};
  if (responseSheet && responseSheet.getLastRow() > 1) {
    const responses = responseSheet.getDataRange().getValues();
    for (let i = 1; i < responses.length; i++) {
      const m = String(responses[i][2] || "").match(/#(\d+)/);
      if (m) {
        if (!checkIns[m[1]]) checkIns[m[1]] = [];
        checkIns[m[1]].push({
          name: String(responses[i][1] || "").trim(),
          pos: String(responses[i][3] || "").trim()
        });
      }
    }
  }

  const auditRows = [];
  games.forEach(g => {
    const reported = checkIns[g.gameId] || [];
    const crReport = reported.find(r => r.pos.toLowerCase().includes("center")) || reported[0];
    const checkedCR = crReport ? crReport.name : "";
    const assignedCR = g.centerRef;

    let status = "Unassigned / Open";
    let points = "—";

    if (!responseSheet) {
      status = (assignedCR && assignedCR !== "[Open]") ? "Pending Check-In" : "Unassigned / Open";
      points = "Pending";
    } else if (assignedCR && assignedCR !== "[Open]" && checkedCR) {
      status = (assignedCR.toLowerCase().replace(/[^a-z]/g, "") === checkedCR.toLowerCase().replace(/[^a-z]/g, "")) 
        ? "Verified Match" 
        : "Substitute / Swap";
      points = "Yes";
    } else if ((!assignedCR || assignedCR === "[Open]") && checkedCR) {
      status = "Walk-On / Fill-In";
      points = "Yes (Credit Volunteer)";
    } else if (assignedCR && assignedCR !== "[Open]" && !checkedCR) {
      status = "Unreported / No-Show";
      points = "0";
    }

    auditRows.push([
      g.gameId, 
      `${g.division}: ${g.matchup} (${g.dateStr} ${g.timeStr})`, 
      g.cleanField, 
      "Center Referee", 
      assignedCR, 
      checkedCR || (responseSheet ? "[None]" : "[Awaiting Form]"), 
      status, 
      points
    ]);
  });

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.AUDIT) || ss.insertSheet(CONFIG.SHEET_TABS.AUDIT);
  targetSheet.clear();

  const times = getPipelineTimestamps();
  targetSheet.getRange(1, 1).setValue(
    `⚡ AYSO 154 REFEREE AUDIT | MatchTrak Schedule Export: ${times.exportTime} | Last Updated in Google Sheets: ${times.sheetTime}`
  );
  targetSheet.getRange(1, 1, 1, 8).merge()
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(1, 28);

  const headers = ["Game #", "Match Details", "Field", "Position", "Assigned in MatchTrak", "Checked-In Volunteer", "Audit Status", "Points Eligible"];
  targetSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  if (auditRows.length > 0) {
    targetSheet.getRange(3, 1, auditRows.length, headers.length).setValues(auditRows);
  }

  targetSheet.getRange(2, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  targetSheet.getRange(3, 1, auditRows.length, 1).setHorizontalAlignment("center");
  targetSheet.getRange(3, 4, auditRows.length, 1).setHorizontalAlignment("center");
  targetSheet.getRange(3, 7, auditRows.length, 2).setHorizontalAlignment("center");

  targetSheet.setFrozenRows(2);
  targetSheet.autoResizeColumns(1, headers.length);
}
