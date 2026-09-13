/**
 * Computes morning setup and evening takedown responsibilities by field
 * for AYSO Region 154 home facilities.
 */

function computeFieldSetupTakedown(games) {
  const fieldGroups = {};
  games.forEach(game => {
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

  const summary = [];
  const sortedKeys = Object.keys(fieldGroups).sort();

  sortedKeys.forEach(key => {
    const group = fieldGroups[key];
    const groupGames = group.games;

    // Chronological sort using absolute minutes from midnight
    groupGames.sort((a, b) => a.minutes - b.minutes);

    const firstGame = groupGames[0];
    const lastGame = groupGames[groupGames.length - 1];

    summary.push({
      dateStr: group.dateStr,
      field: group.field,
      totalGames: groupGames.length,
      setupTime: firstGame.timeStr,
      setupMatch: `${firstGame.division}: ${firstGame.matchup} [#${firstGame.gameId}]`,
      setupGameId: firstGame.gameId,
      takedownTime: lastGame.timeStr,
      takedownMatch: `${lastGame.division}: ${lastGame.matchup} [#${lastGame.gameId}]`,
      takedownGameId: lastGame.gameId
    });
  });

  return summary;
}

function updateFieldSetupTakedownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || ss.getSheetByName(CONFIG.SHEET_TABS.RAW);

  if (!sourceSheet) return;

  const rawData = sourceSheet.getDataRange().getValues();
  if (rawData.length <= 1) return;

  const allGames = transformRawSchedule(rawData);
  const targetDates = getActiveWeekendDates();
  const weekendGames = allGames.filter(g => 
    targetDates.some(td => g.dateStr === td || g.dateStr.startsWith(td))
  );

  const calculated = computeFieldSetupTakedown(weekendGames);
  const headers = ["Date", "Field", "Total Games", "Setup Time (First Kick-off)", "Setup Match (Flags & Nets)", "Takedown Time (Last Kick-off)", "Takedown Match (Pack Up & Lock)"];
  
  const rows = [headers];
  calculated.forEach(c => {
    rows.push([c.dateStr, c.field, c.totalGames, c.setupTime, c.setupMatch, c.takedownTime, c.takedownMatch]);
  });

  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.SETUP_TAKEDOWN) || ss.insertSheet(CONFIG.SHEET_TABS.SETUP_TAKEDOWN);
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  targetSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, headers.length);
}
