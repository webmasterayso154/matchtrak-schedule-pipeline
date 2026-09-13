/**
 * Data transformation and normalization engine for AYSO Region 154 schedules.
 */

function isE154HomeGame(row) {
  if (!row || row.length < 13) return false;
  
  const type = String(row[7] || "").trim().toUpperCase();
  const field = String(row[8] || "").trim();
  const homeTeam = String(row[9] || "").trim();

  if (type === "BYE" || field.includes("BYE") || homeTeam.includes("#BYE")) {
    return false;
  }

  return field.toUpperCase().startsWith(CONFIG.REGION_PREFIX);
}

function cleanFieldName(rawField) {
  if (!rawField) return "";
  return String(rawField)
    .replace(/^E154-/, "")
    .replace(/\s+Fall\s+20\d\d/gi, "")
    .replace(/\s+Fal\s+20\d\d/gi, "")
    .replace(/Elementary School/gi, "ES")
    .replace(/ARTIFICIAL TURF/gi, "Turf")
    .replace(/GRASS\s+/gi, "")
    .replace(/U\d+\s+/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getOperationalVenue(cleanField) {
  const f = String(cleanField).toLowerCase();
  if (f.includes("park lexington")) {
    return "Park Lexington";
  }
  if (f.includes("lexington") || f.includes("arnold")) {
    return "Lexington JHS / Arnold Complex";
  }
  return "Other Cypress Field";
}

function cleanTeamName(rawTeam) {
  if (!rawTeam) return "";
  const str = String(rawTeam).trim();
  const parts = str.split("-");
  
  if (parts.length >= 2) {
    const coachRaw = parts[parts.length - 1];
    const coachName = coachRaw.split("_")[0];
    
    if (!str.includes(CONFIG.REGION_PREFIX)) {
      const regionMatch = str.match(/[A-Z]\d+/);
      const region = regionMatch ? regionMatch[0] : parts[0];
      return `${coachName} (${region})`;
    }
    return coachName;
  }
  return str;
}

function cleanTime(rawTime) {
  if (!rawTime) return "";
  if (rawTime instanceof Date) {
    return Utilities.formatDate(rawTime, Session.getScriptTimeZone(), "h:mm a");
  }
  const str = String(rawTime).trim();
  const match = str.match(/^0?(\d{1,2}:\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (match) {
    const meridiem = match[2] ? ` ${match[2].toUpperCase()}` : "";
    return `${match[1]}${meridiem}`;
  }
  return str;
}

/**
 * Converts 12-hour AM/PM time strings into absolute minutes from midnight
 * to prevent alphabetical sort errors (e.g., 10:00 AM sorting before 8:00 AM).
 */
function parseTimeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const match = String(timeStr).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (!match) return 0;
  let hours = parseInt(match[1], 10);
  const minutes = parseInt(match[2], 10);
  const meridiem = match[3] ? match[3].toUpperCase() : null;
  if (meridiem === "PM" && hours < 12) hours += 12;
  if (meridiem === "AM" && hours === 12) hours = 0;
  return hours * 60 + minutes;
}

function buildFormDropdownLabel(game) {
  return `${game.dateStr} ${game.timeStr} | ${game.division} | ${game.homeCoach} vs ${game.awayCoach} | ${game.cleanField} [#${game.gameId}]`;
}

function transformRawSchedule(rawRows) {
  if (!rawRows || rawRows.length <= 1) return [];

  const cleanGames = [];

  for (let i = 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!isE154HomeGame(row)) continue;

    const gameId = String(row[3] || "").trim();
    const division = String(row[4] || "").trim();
    const rawDate = row[1];
    const dateStr = rawDate instanceof Date 
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M/d") 
      : String(rawDate).replace(/^0/, "");

    const timeStr = cleanTime(row[2]);
    const cleanField = cleanFieldName(row[8]);
    const homeCoach = cleanTeamName(row[9]);
    const awayCoach = cleanTeamName(row[11]);
    const venue = getOperationalVenue(cleanField);

    const gameObject = {
      gameId: gameId,
      circuit: String(row[0] || "").trim(),
      dateStr: dateStr,
      rawDate: rawDate,
      timeStr: timeStr,
      minutes: parseTimeToMinutes(timeStr),
      division: division,
      cleanField: cleanField,
      venue: venue,
      rawField: String(row[8] || "").trim(),
      homeCoach: homeCoach,
      rawHome: String(row[9] || "").trim(),
      awayCoach: awayCoach,
      rawAway: String(row[11] || "").trim(),
      matchup: `${homeCoach} vs ${awayCoach}`,
      type: String(row[7] || "").trim()
    };

    gameObject.dropdownLabel = buildFormDropdownLabel(gameObject);
    cleanGames.push(gameObject);
  }

  return cleanGames;
}
