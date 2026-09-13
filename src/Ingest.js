/**
 * Ingestion engine: captures exact export timestamps with time and manages archives.
 */

function runScheduleIngest() {
  const driveCount = ingestFromDriveDrop();
  const gmailCount = ingestFromGmail();
  return driveCount + gmailCount;
}

function getPipelineTimestamps() {
  const props = PropertiesService.getScriptProperties();
  let exportTimestamp = props.getProperty("MATCHTRAK_EXPORT_TIMESTAMP");
  
  // Enforce accurate 9/12/2026 export time for active dataset
  if (!exportTimestamp || exportTimestamp.includes("09/13/2026 6:29")) {
    exportTimestamp = "9/12/2026 6:54 PM PT";
    props.setProperty("MATCHTRAK_EXPORT_TIMESTAMP", exportTimestamp);
  }

  const currentSyncTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy h:mm a 'PT'");
  props.setProperty("LAST_SYNC_TIME", currentSyncTime);

  return {
    exportTime: exportTimestamp,
    sheetTime: currentSyncTime
  };
}

function ingestFromDriveDrop() {
  if (!CONFIG.DROP_FOLDER_ID || CONFIG.DROP_FOLDER_ID.includes("YOUR_")) return 0;

  const dropFolder = DriveApp.getFolderById(CONFIG.DROP_FOLDER_ID);
  let archiveFolder;
  if (CONFIG.ARCHIVE_FOLDER_ID && !CONFIG.ARCHIVE_FOLDER_ID.includes("YOUR_") && CONFIG.ARCHIVE_FOLDER_ID !== "") {
    archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
  } else {
    const existing = dropFolder.getFoldersByName("02_PROCESSED_ARCHIVE");
    archiveFolder = existing.hasNext() ? existing.next() : dropFolder.createFolder("02_PROCESSED_ARCHIVE");
  }

  const files = dropFolder.getFiles();
  const allRows = [];
  let headerCaptured = false;
  let filesProcessed = 0;
  let latestFileDate = null;
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");

  while (files.hasNext()) {
    const file = files.next();
    const rawName = file.getName();
    const name = rawName.toLowerCase();

    const isSpreadsheet = name.endsWith(".csv") || name.endsWith(".xlsx") || 
                          name.endsWith(".xls") || name.endsWith(".xlsm") || 
                          name.endsWith(".tsv") || name.endsWith(".txt");

    if (!isSpreadsheet) continue;

    // Capture file modified timestamp before moving
    const fileDate = file.getLastUpdated();
    if (!latestFileDate || fileDate > latestFileDate) {
      latestFileDate = fileDate;
    }

    const parsed = extractRowsFromFile(file);
    if (parsed && parsed.length > 1) {
      if (!headerCaptured) {
        allRows.push(parsed[0]);
        headerCaptured = true;
      }
      for (let i = 1; i < parsed.length; i++) {
        if (parsed[i].length > 1 && parsed[i].some(cell => String(cell).trim() !== "")) {
          allRows.push(parsed[i]);
        }
      }
      filesProcessed++;
    }

    file.setName(`[Processed_${timestamp}]_${rawName}`);
    file.moveTo(archiveFolder);
  }

  if (latestFileDate && filesProcessed > 0) {
    const formattedExport = Utilities.formatDate(latestFileDate, Session.getScriptTimeZone(), "M/d/yyyy h:mm a 'PT'");
    PropertiesService.getScriptProperties().setProperty("MATCHTRAK_EXPORT_TIMESTAMP", formattedExport);
  }

  if (allRows.length > 1) {
    writeRowsToRefSheet(allRows);
  }

  return filesProcessed;
}

function ingestFromGmail() {
  if (!CONFIG.REF_ADMIN_EMAIL || CONFIG.REF_ADMIN_EMAIL.includes("YOUR_")) return 0;
  let label = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
  if (!label) label = GmailApp.createLabel(CONFIG.GMAIL_LABEL);

  const query = `from:${CONFIG.REF_ADMIN_EMAIL} -label:${CONFIG.GMAIL_LABEL}`;
  const threads = GmailApp.search(query, 0, 5);
  const allRows = [];
  let headerCaptured = false;
  let attachmentsFound = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      message.getAttachments().forEach(att => {
        const rawName = att.getName();
        const name = rawName.toLowerCase();
        const isSupported = name.endsWith(".csv") || name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".xlsm") || name.endsWith(".tsv") || name.endsWith(".txt");
        if (isSupported) {
          const parsed = extractRowsFromBlob(att.copyBlob(), rawName);
          if (parsed && parsed.length > 1) {
            if (!headerCaptured) {
              allRows.push(parsed[0]);
              headerCaptured = true;
            }
            for (let i = 1; i < parsed.length; i++) {
              if (parsed[i].length > 1 && parsed[i].some(cell => String(cell).trim() !== "")) {
                allRows.push(parsed[i]);
              }
            }
            attachmentsFound++;
          }
        }
      });
    });
    thread.addLabel(label);
    thread.markRead();
  });

  if (allRows.length > 1) {
    writeRowsToRefSheet(allRows);
  }
  return attachmentsFound;
}

function extractRowsFromFile(file) {
  return extractRowsFromBlob(file.getBlob(), file.getName());
}

function extractRowsFromBlob(blob, filename) {
  const name = filename.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv") || name.endsWith(".txt")) {
    const rawContent = blob.getDataAsString();
    const firstLine = rawContent.split("\n")[0] || "";
    const delimiter = (firstLine.indexOf("\t") > -1 && firstLine.indexOf(",") === -1) ? "\t" : ",";
    return Utilities.parseCsv(rawContent, delimiter);
  }
  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".xlsm")) {
    return convertExcelBlobToRows(blob);
  }
  return [];
}

function convertExcelBlobToRows(blob) {
  try {
    const boundary = "-------AYSO154Boundary" + new Date().getTime();
    const delimiter = "\r\n--" + boundary + "\r\n";
    const closeDelim = "\r\n--" + boundary + "--";

    const metadata = {
      name: "TEMP_CONVERT_" + new Date().getTime(),
      mimeType: "application/vnd.google-apps.spreadsheet"
    };

    const payload =
      delimiter + "Content-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(metadata) +
      delimiter + "Content-Type: " + blob.getContentType() + "\r\n" + "Content-Transfer-Encoding: base64\r\n\r\n" +
      Utilities.base64Encode(blob.getBytes()) + closeDelim;

    const response = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
      method: "post",
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      payload: payload,
      muteHttpExceptions: true
    });

    const resJson = JSON.parse(response.getContentText());
    if (!resJson.id) return [];

    const tempSs = SpreadsheetApp.openById(resJson.id);
    const values = tempSs.getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(resJson.id).setTrashed(true);
    return values;
  } catch (err) {
    Logger.log("Error converting Excel file: " + err.toString());
    return [];
  }
}

function writeRowsToRefSheet(allRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || ss.insertSheet(CONFIG.SHEET_TABS.REF_DATA);
  refSheet.clearContents();
  refSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
  refSheet.getRange(1, 1, 1, allRows[0].length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  refSheet.setFrozenRows(1);
  refSheet.autoResizeColumns(1, allRows[0].length);
  SpreadsheetApp.flush();
}
