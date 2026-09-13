/**
 * Main application controller and UI menu for AYSO Region 154 MatchTrak Pipeline.
 * Binds ingestion, transformation, form sync, setup/takedown, and referee audits.
 */

/**
 * Automatically executed when the Google Sheet is opened.
 * Inserts a dedicated "AYSO 154 Ops" menu into the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚽ AYSO 154 Ops")
    .addItem("🔄 Run Full Weekend Sync", "menuRunFullSync")
    .addSeparator()
    .addItem("📥 Check Email & Drive Drops", "menuIngestOnly")
    .addItem("📋 Refresh Google Form Dropdown", "menuSyncFormOnly")
    .addItem("🥅 Update Field Setup / Takedown", "menuSetupTakedownOnly")
    .addItem("⚖️ Reconcile Referee Audit", "menuAuditOnly")
    .addToUi();
}

/**
 * Master orchestrator: Ingests new files, updates Form choices,
 * generates Setup/Takedown tables, and refreshes the Audit sheet.
 */
function menuRunFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Checking Gmail & Drive for latest MatchTrak files...", "AYSO 154 Sync", 5);

  try {
    // 1. Ingest files from Gmail and Drive drop folder
    runScheduleIngest();

    // 2. Synchronize Form dropdown
    ss.toast("Updating Game Day Google Form...", "AYSO 154 Sync", 5);
    syncFormDropdown();

    // 3. Compute Field Setup and Takedown duties
    ss.toast("Calculating Setup & Takedown matches...", "AYSO 154 Sync", 5);
    updateFieldSetupTakedownSheet();

    // 4. Run referee check-in audit
    ss.toast("Auditing referee coverage...", "AYSO 154 Sync", 5);
    runRefereeAudit();

    ss.toast("All operations successfully synchronized!", "Complete", 7);
  } catch (err) {
    Logger.log("Error during full sync: " + err.toString());
    SpreadsheetApp.getUi().alert("Sync encountered an error: " + err.message);
  }
}

/**
 * Menu action: Ingest files only.
 */
function menuIngestOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Checking Gmail and Drive drop folder...", "Ingesting", 5);
  runScheduleIngest();
  ss.toast("Ingestion complete.", "Complete", 5);
}

/**
 * Menu action: Refresh Google Form dropdown only.
 */
function menuSyncFormOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Pushing upcoming matches to Google Form...", "Updating Form", 5);
  syncFormDropdown();
  ss.toast("Google Form dropdown updated.", "Complete", 5);
}

/**
 * Menu action: Recalculate Field Setup and Takedown only.
 */
function menuSetupTakedownOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Calculating morning and afternoon duties...", "Setup/Takedown", 5);
  updateFieldSetupTakedownSheet();
  ss.toast("Setup/Takedown tab updated.", "Complete", 5);
}

/**
 * Menu action: Re-run referee audit only.
 */
function menuAuditOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Cross-referencing MatchTrak against Form submissions...", "Auditing", 5);
  runRefereeAudit();
  ss.toast("Referee Audit tab updated.", "Complete", 5);
}
