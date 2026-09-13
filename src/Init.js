/**
 * Initialization and environment setup module for AYSO Region 154 Sheet.
 * Handles automatic tab provisioning, formatting, and test data seeding.
 */

/**
 * Creates all required sheet tabs with styled headers and frozen rows.
 */
function initSheetTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const tabSchemas = [
    {
      name: CONFIG.SHEET_TABS.RAW,
      headers: ["Circuit", "Date", "Time", "Game #", "Division", "Flight #", "Group", "Type", "Field", "Home Team", "Home Score", "Away Team", "Away Score", "Center Referee", "AR1", "AR2"]
    },
    {
      name: CONFIG.SHEET_TABS.MASTER_HOME,
      headers: ["Game #", "Date", "Time", "Division", "Field", "Venue", "Matchup", "Home Coach", "Away Coach", "Dropdown Label", "Circuit"]
    },
    {
      name: CONFIG.SHEET_TABS.REF_DATA,
      headers: ["Circuit", "Date", "Time", "Game #", "Division", "Flight #", "Group", "Type", "Field", "Home Team", "Home Score", "Away Team", "Away Score", "Center Referee", "AR1", "AR2"]
    },
    {
      name: CONFIG.SHEET_TABS.SETUP_TAKEDOWN,
      headers: ["Date", "Field", "Total Games", "Setup Time (First Kick-off)", "Setup Match (Flags & Nets)", "Takedown Time (Last Kick-off)", "Takedown Match (Pack Up & Lock)"]
    },
    {
      name: CONFIG.SHEET_TABS.AUDIT,
      headers: ["Game #", "Match Details", "Field", "Position", "Assigned in MatchTrak", "Checked-In Volunteer", "Audit Status", "Points Eligible"]
    }
  ];

  tabSchemas.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
    }
    
    // Set headers if blank
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
      sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
    }

    // Format header row (AYSO Navy Blue & White text)
    sheet.getRange(1, 1, 1, schema.headers.length)
      .setFontWeight("bold")
      .setBackground("#1b365d")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
      
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, schema.headers.length);
  });

  // Remove empty default Sheet1 if additional tabs exist
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  ss.toast("All tabs created and styled successfully.", "Setup Complete", 5);
}

/**
 * Seeds sample MatchTrak test records into the sheet to test transformations,
 * Setup/Takedown tables, and Master Home tabs without connecting Gmail or Drive.
 */
function seedTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || 
                   ss.getSheetByName(CONFIG.SHEET_TABS.RAW);

  if (!rawSheet) {
    initSheetTabs();
  }

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  const weekendDates = getActiveWeekendDates();
  const satDate = weekendDates[1] || "9/19/2026";

  const sampleRows = [
    ["Circuit", "Date", "Time", "Game #", "Division", "Flight #", "Group", "Type", "Field", "Home Team", "Home Score", "Away Team", "Away Score", "Center Referee", "AR1", "AR2"],
    ["Core", satDate, "08:00:00 AM", "22101", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Richardson_T", "", "11-Z106-Gastelum_G", "", "John Smith", "Mike Jones", ""],
    ["Core", satDate, "10:30:00 AM", "22102", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Mancilla_A", "", "11-E154-Zidan_H", "", "Dave Miller", "", ""],
    ["Core", satDate, "01:00:00 PM", "22103", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Hernandez_C", "", "11-E114-Baker_J", "", "Unassigned", "", ""],
    ["Core", satDate, "08:30:00 AM", "22104", "GU12", "1", "A", "REG", "E154-Arnold Elementary School Grass Field 10", "11-E154-Foster_D", "", "11-Q97-Smith_R", "", "Alice Wong", "", ""],
    ["Extra", satDate, "09:00:00 AM", "22105", "BU14", "1", "A", "REG", "E154-Park Lexington ARTIFICIAL TURF Fall 2026", "11-E154-Alvarez_M", "", "11-E114-Baker_J", "", "Carlos Ruiz", "", ""],
    ["Core", satDate, "12:01:00 AM", "22106", "BU10", "1", "A", "BYE", "E154-Lexington JHS U10 Field 7 Fall 2026", "#BYE", "", "11-E154-Richardson_T", "", "", "", ""],
    ["Area E", satDate, "09:00:00 AM", "22107", "BU12", "1", "A", "REG", "E114-Los Alamitos Field 2", "11-E114-Davis_M", "", "11-E154-Richardson_T", "", "", "", ""]
  ];

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, sampleRows.length, sampleRows[0].length).setValues(sampleRows);
  
  // Format header
  targetSheet.getRange(1, 1, 1, sampleRows[0].length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff");

  ss.toast("Sample test data seeded into MatchTrak_Ref_Data.", "Data Ready", 5);
}
