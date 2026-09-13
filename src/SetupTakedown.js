/**
 * Computes morning setup and evening takedown responsibilities by field
 * for AYSO Region 154 home facilities.
 */

/**
 * Main routine: Analyzes upcoming matches, computes first and last games
 * per field, and populates the Setup_Takedown tab in Google Sheets.
 */
function updateFieldSetupTakedownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || ss.getSheetByName(CONFIG.SHEET_TABS.RAW);

  if (!sourceSheet) {
    Logger.log("Source data sheet missing for setup/takedown calculation.");
    return;
  }

  const rawData = sourceSheet.getDataRange().getValues();
  if (rawData.length <= 1) return;

  const allGames = transformRawSchedule(rawData);
  const targetDates = getActiveWeekendDates();

  // Filter for upcoming Friday - Sunday games
  const weekendGames = allGames.filter(game => 
    targetDates.some(td => game.dateStr === td || game.dateStr.startsWith(td))
  );

  if (weekendGames.length === 0) {
    Logger.log("No weekend games found for setup/takedown.");
    return;
  }

  // Group games by unique Date + Field combination
  const fieldGroups = {};
  weekendGames.forEach(game => {
    const key = `${game.dateStr}___${game.cleanField}`;
    if (!fieldGroups[key]) {
      fieldGroups[key] = {
        dateStr: game.dateStr,
        field: game.cleanField,
        games: []
      };
    }
    fieldGroups[key].games.push(game);
  });

  // Calculate first and last game per field
  const summaryRows = [];
  const headers = [
    "Date", 
    "Field", 
    "Total Games", 
    "Setup Time (First Kick-off)", 
    "Setup Match (Flags & Nets)", 
    "Takedown Time (Last Kick-off)", 
    "Takedown Match (Pack Up & Lock)"
  ];
  summaryRows.push(headers);

  // Sort groups by date then field name
  const sortedKeys = Object.keys(fieldGroups).sort();

  sortedKeys.forEach(key => {
    const group = fieldGroups[key];
    const games = group.games;

    // Sort games chronologically by time
    games.sort((a, b) => {
      const timeA = new Date(a.rawDate + " " + a.timeStr).getTime() || 0;
      const timeB = new Date(b.rawDate + " " + b.timeStr).getTime() || 0;
      return timeA - timeB;
    });

    const firstGame = games[0];
    const lastGame = games[games.length - 1];

    summaryRows.push([
      group.dateStr,
      group.field,
      games.length,
      firstGame.timeStr,
      `${firstGame.division}: ${firstGame.matchup} [#${firstGame.gameId}]`,
      lastGame.timeStr,
      `${lastGame.division}: ${lastGame.matchup} [#${lastGame.gameId}]`
    ]);
  });

  // Write results to the Setup_Takedown tab
  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.SETUP_TAKEDOWN) || 
                      ss.insertSheet(CONFIG.SHEET_TABS.SETUP_TAKEDOWN);

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, summaryRows.length, summaryRows[0].length).setValues(summaryRows);
  
  // Apply visual styling
  targetSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff");
  
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, headers.length);

  Logger.log(`Setup/Takedown tab updated with ${summaryRows.length - 1} field entries.`);
}
