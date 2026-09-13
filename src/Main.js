/**
 * Main application controller and UI menu for AYSO Region 154 MatchTrak Pipeline.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("? AYSO 154 Ops")
    .addItem("?? Run Full Weekend Sync", "menuRunFullSync")
    .addSeparator()
    .addItem("?? Update Setup & Takedown Sheet", "menuSetupTakedownOnly")
    .addItem("?? Populate Master Home Sheet", "menuPopulateMasterHomeOnly")
    .addItem("?? Run Referee Coverage Audit", "menuAuditOnly")
    .addSeparator()
    .addItem("??? Initialize Sheet Tabs & Headers", "initSheetTabs")
    .addItem("?? Seed Sample Match Data", "seedTestData")
    .addToUi();
}

function menuRunFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast("Reading schedule data...", "AYSO 154 Sync", 3);
    
    // Master home sheet population
    menuPopulateMasterHomeOnly();

    // Setup & takedown table computation
    menuSetupTakedownOnly();

    // Form sync (gracefully skipped if FORM_ID is disabled/blank)
    if (CONFIG.FORM_ID && !CONFIG.FORM_ID.includes("YOUR_") && CONFIG.FORM_ID !== "") {
      syncFormDropdown();
    }

    // Referee audit
    menuAuditOnly();

    ss.toast("All operational sheets synchronized!", "Success", 5);
  } catch (err) {
    Logger.log("Error during full sync: " + err.toString());
    SpreadsheetApp.getUi().alert("Sync encountered an error: " + err.message);
  }
}

function menuPopulateMasterHomeOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || 
                      ss.getSheetByName(CONFIG.SHEET_TABS.RAW);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return;

  const rawRows = sourceSheet.getDataRange().getValues();
  const games = transformRawSchedule(rawRows);
  if (games.length === 0) return;

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.MASTER_HOME) || 
                      ss.insertSheet(CONFIG.SHEET_TABS.MASTER_HOME);

  const rows = games.map(g => [
    g.gameId, g.dateStr, g.timeStr, g.division, g.cleanField, g.venue, g.matchup, g.homeCoach, g.awayCoach, g.dropdownLabel, g.circuit
  ]);

  const headers = ["Game #", "Date", "Time", "Division", "Field", "Venue", "Matchup", "Home Coach", "Away Coach", "Dropdown Label", "Circuit"];

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  targetSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  targetSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, headers.length);
}

function menuSetupTakedownOnly() {
  updateFieldSetupTakedownSheet();
}

function menuAuditOnly() {
  runRefereeAudit();
}
