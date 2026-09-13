/**
 * Setup & Takedown calculation engine with auto-sizing and coach assignments.
 */

function computeFieldSetupTakedown(games) {
  const fieldGroups = {};
  games.forEach(game => {
    const key = `${game.dateStr}___${game.cleanField}`;
    if (!fieldGroups[key]) {
      fieldGroups[key] = {
        dateStr: game.dateStr,
        timestamp: game.timestamp,
        venue: game.venue,
        field: game.cleanField,
        games: []
      };
    }
    fieldGroups[key].games.push(game);
  });

  const groupKeys = Object.keys(fieldGroups);
  groupKeys.sort((a, b) => {
    if (fieldGroups[a].timestamp !== fieldGroups[b].timestamp) {
      return fieldGroups[a].timestamp - fieldGroups[b].timestamp;
    }
    return fieldGroups[a].field.localeCompare(fieldGroups[b].field);
  });

  const summary = [];
  groupKeys.forEach(key => {
    const group = fieldGroups[key];
    const groupGames = group.games;

    groupGames.sort((a, b) => a.minutes - b.minutes);

    const first = groupGames[0];
    const last = groupGames[groupGames.length - 1];

    summary.push({
      dateStr: group.dateStr,
      venue: group.venue,
      field: group.field,
      totalGames: groupGames.length,
      setupTime: first.timeStr,
      setupMatch: `${first.division}: ${first.matchup} [#${first.gameId}]`,
      setupDuty: `Home: Coach ${first.homeCoach} (Sets Flags/Nets)`,
      takedownTime: last.timeStr,
      takedownMatch: `${last.division}: ${last.matchup} [#${last.gameId}]`,
      takedownDuty: `Both: Coaches ${last.homeCoach} & ${last.awayCoach} (Pack & Lock)`
    });
  });

  return summary;
}

function updateFieldSetupTakedownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return;

  const rawData = sourceSheet.getDataRange().getValues();
  const allGames = transformRawSchedule(rawData);
  if (allGames.length === 0) return;

  const calculated = computeFieldSetupTakedown(allGames);
  const targetSheet = ss.getSheetByName(CONFIG.SHEET_TABS.SETUP_TAKEDOWN) || ss.insertSheet(CONFIG.SHEET_TABS.SETUP_TAKEDOWN);
  targetSheet.clear();

  const times = getPipelineTimestamps();
  targetSheet.getRange(1, 1).setValue(
    `⚡ AYSO 154 FIELD LOGISTICS | MatchTrak Schedule Export: ${times.exportTime} | Last Updated in Google Sheets: ${times.sheetTime} | Equipment: Setup 45 mins prior to kickoff; Takedown locks into bin.`
  );
  targetSheet.getRange(1, 1, 1, 12).merge()
    .setFontWeight("bold")
    .setFontColor("#1e293b")
    .setBackground("#e2e8f0")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(1, 28);

  const headers = [
    "Date", "Operational Venue", "Field", "Games", 
    "Setup Time", "Setup Game", "Setup Responsible", "Setup Done", 
    "Takedown Time", "Takedown Game", "Takedown Responsible", "Takedown Done"
  ];
  targetSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  targetSheet.getRange(2, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1b365d")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  targetSheet.setRowHeight(2, 32);

  const rows = calculated.map(c => [
    c.dateStr, c.venue, c.field, c.totalGames,
    c.setupTime, c.setupMatch, c.setupDuty, false,
    c.takedownTime, c.takedownMatch, c.takedownDuty, false
  ]);

  if (rows.length > 0) {
    const dataRange = targetSheet.getRange(3, 1, rows.length, headers.length);
    dataRange.setValues(rows);

    targetSheet.getRange(3, 4, rows.length, 1).setNumberFormat("0");
    targetSheet.getRange(3, 8, rows.length, 1).insertCheckboxes();
    targetSheet.getRange(3, 12, rows.length, 1).insertCheckboxes();

    targetSheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment("center");
    targetSheet.getRange(3, 4, rows.length, 2).setHorizontalAlignment("center");
    targetSheet.getRange(3, 9, rows.length, 1).setHorizontalAlignment("center");
    dataRange.setVerticalAlignment("middle");

    for (let i = 0; i < rows.length; i++) {
      targetSheet.setRowHeight(3 + i, 26);
      if (i % 2 === 1) targetSheet.getRange(3 + i, 1, 1, headers.length).setBackground("#f8fafc");
    }
  }

  const widths = [75, 200, 160, 65, 95, 240, 240, 85, 105, 240, 260, 100];
  widths.forEach((w, idx) => targetSheet.setColumnWidth(idx + 1, w));

  targetSheet.setFrozenRows(2);
}
