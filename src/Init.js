/**
 * Tab Provisioning Module.
 */

function initSheetTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const legacyTabs = ["Raw_Import", "Sheet1"];
  legacyTabs.forEach(tName => {
    const s = ss.getSheetByName(tName);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  const tabSchemas = [
    { name: CONFIG.SHEET_TABS.DASHBOARD },
    { name: CONFIG.SHEET_TABS.MASTER_HOME },
    { name: CONFIG.SHEET_TABS.REF_DATA },
    { name: CONFIG.SHEET_TABS.SETUP_TAKEDOWN },
    { name: CONFIG.SHEET_TABS.MISSING_SCORES },
    { name: CONFIG.SHEET_TABS.AUDIT }
  ];

  tabSchemas.forEach((schema, idx) => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name, idx);
    }
  });
}
