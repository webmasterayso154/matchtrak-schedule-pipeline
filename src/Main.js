/**
 * Main application controller for AYSO Region 154 MatchTrak Pipeline.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚽ AYSO 154 Ops")
    .addItem("🔄 Sync Schedule & Logistics Now", "menuRunFullSync")
    .addItem("🧹 Clear Checked Missing Scores", "archiveResolvedMissingScores")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("⚙️ Admin Settings")
        .addItem("Set MatchTrak Export Timestamp", "adminSetExportTimestamp")
        .addItem("Reset Outreach Checkboxes", "resetAllMissingScoreCheckboxes")
        .addItem("Rebuild Tab Formatting", "initSheetTabs")
        .addSeparator()
        .addItem("Turn On Hourly Background Auto-Sync", "installAutoTrigger")
        .addItem("Turn Off Background Auto-Sync", "uninstallAutoTrigger")
    )
    .addToUi();
}

function menuRunFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    initSheetTabs();

    ss.toast("Checking Drive drop folder...", "Full Sync", 2);
    runScheduleIngest();

    const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
    if (!sourceSheet || sourceSheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert("No schedule data found in MatchTrak_Ref_Data. Drop files into the Drive folder first.");
      return;
    }

    ss.toast("Generating Executive Dashboard...", "Full Sync", 2);
    generateExecutiveDashboard();

    ss.toast("Populating Upcoming & Season Home matches...", "Full Sync", 2);
    menuPopulateMasterHomeOnly();

    ss.toast("Calculating Setup & Takedown duties...", "Full Sync", 2);
    menuSetupTakedownOnly();

    ss.toast("Auditing missing scores...", "Full Sync", 2);
    menuMissingScoresOnly();

    if (CONFIG.FORM_ID && !CONFIG.FORM_ID.includes("YOUR_") && CONFIG.FORM_ID !== "") {
      syncFormDropdown();
    }

    ss.toast("Reconciling referee coverage...", "Full Sync", 2);
    menuAuditOnly();

    const dash = ss.getSheetByName(CONFIG.SHEET_TABS.DASHBOARD);
    if (dash) dash.activate();

    ss.toast("All operational sheets synchronized!", "Success", 4);
  } catch (err) {
    Logger.log("Error during full sync: " + err.toString());
    SpreadsheetApp.getUi().alert("Sync encountered an error: " + err.message);
  }
}

function menuPopulateMasterHomeOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return;

  const rawRows = sourceSheet.getDataRange().getValues();
  const allGames = transformRawSchedule(rawRows);
  if (allGames.length === 0) return;

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.MASTER_HOME) || 
                      ss.insertSheet(CONFIG.SHEET_TABS.MASTER_HOME);

  targetSheet.clear();

  const times = getPipelineTimestamps();
  targetSheet.getRange(1, 1).setValue(
    `⚽ AYSO REGION 154 MASTER HOME FIXTURES | MatchTrak Schedule Export: ${times.exportTime} | Last Updated in Google Sheets: ${times.sheetTime} | Total: ${allGames.length} Home Matches`
  );
  targetSheet.getRange(1, 1, 1, 11).merge()
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(1, 28);

  const headers = ["Game #", "Date", "Time", "Division", "Field", "Venue", "Matchup", "Home Coach", "Away Coach", "Dropdown Label", "Circuit"];
  targetSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  targetSheet.getRange(2, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(2, 32);

  // Partition games into Upcoming (Next 14 Days), Later Season, and Past Completed
  const now = new Date();
  const todayTimestamp = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const fourteenDaysFuture = todayTimestamp + (14 * 24 * 60 * 60 * 1000);

  const upcomingGames = [];
  const futureGames = [];
  const pastGames = [];

  allGames.forEach(g => {
    if (g.timestamp < todayTimestamp) {
      pastGames.push(g);
    } else if (g.timestamp <= fourteenDaysFuture) {
      upcomingGames.push(g);
    } else {
      futureGames.push(g);
    }
  });

  let currentRow = 3;

  function renderGameBlock(sectionTitle, gamesList, barBgColor, barTxtColor) {
    if (gamesList.length === 0) return;

    targetSheet.getRange(currentRow, 1).setValue(sectionTitle);
    targetSheet.getRange(currentRow, 1, 1, headers.length).merge()
      .setFontWeight("bold")
      .setFontColor(barTxtColor)
      .setBackground(barBgColor)
      .setFontSize(10)
      .setVerticalAlignment("middle");
    targetSheet.setRowHeight(currentRow, 26);
    currentRow++;

    const rows = gamesList.map(g => [
      g.gameId, g.dateStr, g.timeStr, g.division, g.cleanField, g.venue, g.matchup, g.homeCoach, g.awayCoach, g.dropdownLabel, g.circuit
    ]);

    const dataRange = targetSheet.getRange(currentRow, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    targetSheet.getRange(currentRow, 1, rows.length, 4).setHorizontalAlignment("center");
    dataRange.setVerticalAlignment("middle");

    for (let i = 0; i < rows.length; i++) {
      const rIdx = currentRow + i;
      targetSheet.setRowHeight(rIdx, 25);
      if (i % 2 === 1) targetSheet.getRange(rIdx, 1, 1, headers.length).setBackground("#f8fafc");
    }

    currentRow += rows.length;
  }

  // 1. Upcoming Matches (9/18 - 9/26) Rendered First
  renderGameBlock("🌟 UPCOMING GAME DAYS (Next 14 Days: 9/18 - 9/26)", upcomingGames, "#0f172a", "#f59e0b");

  // 2. Future Matches (October and beyond)
  renderGameBlock("📅 FULL REMAINING SEASON FIXTURES (October & Beyond)", futureGames, "#334155", "#ffffff");

  // 3. Past Completed Matches (August 22 - September 12) Placed at Bottom
  renderGameBlock("📁 COMPLETED MATCHES (August 22 – September 12 Archive)", pastGames, "#94a3b8", "#1e293b");

  const widths = [85, 75, 95, 80, 180, 210, 240, 130, 130, 360, 130];
  widths.forEach((w, idx) => targetSheet.setColumnWidth(idx + 1, w));

  targetSheet.setFrozenRows(2);
}

function adminSetExportTimestamp() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty("MATCHTRAK_EXPORT_TIMESTAMP") || "9/12/2026 6:54 PM PT";
  const prompt = ui.prompt(
    "Set MatchTrak Export Timestamp",
    "Enter the exact Date and Time when MatchTrak schedules were exported (e.g., 9/12/2026 6:54 PM PT):",
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() === ui.Button.OK) {
    const entered = prompt.getResponseText().trim();
    if (entered) {
      PropertiesService.getScriptProperties().setProperty("MATCHTRAK_EXPORT_TIMESTAMP", entered);
      menuRunFullSync();
      ui.alert("Timestamp updated and all sheet banners refreshed!");
    }
  }
}

function menuSetupTakedownOnly() {
  updateFieldSetupTakedownSheet();
}

function menuMissingScoresOnly() {
  generateMissingScoresReport();
}

function menuAuditOnly() {
  runRefereeAudit();
}

function installAutoTrigger() {
  uninstallAutoTrigger();
  ScriptApp.newTrigger("menuRunFullSync")
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert("Hourly background sync enabled.");
}

function uninstallAutoTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "menuRunFullSync") {
      ScriptApp.deleteTrigger(t);
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Hourly background sync disabled.", "Settings Updated", 3);
}
