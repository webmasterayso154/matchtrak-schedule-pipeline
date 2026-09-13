/**
 * Ingestion engine for AYSO Region 154 MatchTrak schedule exports.
 * Supports dual-source intake via Gmail attachments and Google Drive folder drops.
 */

/**
 * Main entry point: checks both Gmail and Google Drive sources in a single run.
 * Schedule a 10-15 minute time-driven trigger on this function.
 */
function runScheduleIngest() {
  ingestFromGmail();
  ingestFromDriveDrop();
}

/**
 * 1. Checks Gmail for unread MatchTrak CSV attachments sent by the Ref Admin.
 */
function ingestFromGmail() {
  if (!CONFIG.REF_ADMIN_EMAIL || CONFIG.REF_ADMIN_EMAIL.includes("YOUR_")) {
    Logger.log("Ref Admin email not configured in Config.js. Skipping Gmail check.");
    return;
  }

  let processedLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(CONFIG.GMAIL_LABEL);
  }

  const query = `from:${CONFIG.REF_ADMIN_EMAIL} filename:csv -label:${CONFIG.GMAIL_LABEL}`;
  const threads = GmailApp.search(query, 0, 5);

  if (threads.length === 0) {
    Logger.log("Gmail check: No new referee schedule emails found.");
    return;
  }

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const attachments = message.getAttachments();
      
      attachments.forEach(att => {
        const fileName = att.getName().toLowerCase();
        if (fileName.endsWith(".csv") || att.getContentType() === "text/csv") {
          const success = writeCsvToSheet(att.getDataAsString(), CONFIG.SHEET_TABS.REF_DATA, `Email (${att.getName()})`);
          
          if (success) {
            message.reply(
              `Hi,\n\nYour MatchTrak schedule (${att.getName()}) was successfully received and loaded into the Region 154 system.\n\n` +
              `Status: Sheets updated and ready for weekend operations.\n\n` +
              `AYSO Region 154 Automated Ops`
            );
          }
        }
      });
    });

    thread.addLabel(processedLabel);
    thread.markRead();
  });
}

/**
 * 2. Checks the Shared Google Drive drop folder for uploaded CSV exports.
 */
function ingestFromDriveDrop() {
  if (!CONFIG.DROP_FOLDER_ID || CONFIG.DROP_FOLDER_ID.includes("YOUR_")) {
    Logger.log("Drive folder IDs not configured in Config.js. Skipping Drive check.");
    return;
  }

  const dropFolder = DriveApp.getFolderById(CONFIG.DROP_FOLDER_ID);
  const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
  const files = dropFolder.getFilesByType(MimeType.CSV);

  let processedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const csvContent = file.getBlob().getDataAsString();

    const success = writeCsvToSheet(csvContent, CONFIG.SHEET_TABS.REF_DATA, `Drive Drop (${fileName})`);

    if (success) {
      file.moveTo(archiveFolder);
      processedCount++;

      // Send confirmation receipt to Ref Admin
      GmailApp.sendEmail(
        CONFIG.REF_ADMIN_EMAIL,
        "MatchTrak Schedule Ingested (Drive Upload)",
        `Hi,\n\nYour file "${fileName}" uploaded to Google Drive was successfully processed into the Region 154 Master Sheet.\n\n` +
        `AYSO Region 154 Automated Ops`
      );
    }
  }

  Logger.log(`Drive check complete. Processed ${processedCount} file(s).`);
}

/**
 * Shared writer: Clears target tab, writes raw 2D array, and triggers form sync.
 * @param {string} csvString - Unprocessed CSV string.
 * @param {string} targetTabName - Sheet tab destination.
 * @param {string} sourceLabel - Origin descriptor for logging.
 * @return {boolean} True if data was written successfully.
 */
function writeCsvToSheet(csvString, targetTabName, sourceLabel) {
  const parsedMatrix = Utilities.parseCsv(csvString);
  if (!parsedMatrix || parsedMatrix.length <= 1) {
    Logger.log(`Skipping invalid/empty file from ${sourceLabel}`);
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetTabName) || ss.insertSheet(targetTabName);

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, parsedMatrix.length, parsedMatrix[0].length).setValues(parsedMatrix);
  SpreadsheetApp.flush();

  Logger.log(`Successfully wrote ${parsedMatrix.length - 1} rows from ${sourceLabel} to tab "${targetTabName}".`);

  // Hook: refresh Google Form dropdown if FormSync module is present
  if (typeof syncFormDropdown === "function") {
    syncFormDropdown();
  }

  return true;
}
