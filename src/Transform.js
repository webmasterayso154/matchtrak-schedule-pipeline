/**
 * Data transformation engine with bidirectional BYE filtering and chronological sorting.
 */

function isE154HomeGame(row, headerMap) {
  if (!row || row.length < 10) return false;
  
  const type = String(row[headerMap.type] || "").trim().toUpperCase();
  const field = String(row[headerMap.field] || "").trim();
  const home = String(row[headerMap.home] || "").trim();
  const away = String(row[headerMap.away] || "").trim();
  const time = String(row[headerMap.time] || "").trim();

  // Exclude BYEs from either side and games without valid times
  if (type === "BYE" || field.includes("BYE") || home.includes("#BYE") || away.includes("#BYE") || !time) {
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
  if (f.includes("park lexington")) return "Park Lexington";
  if (f.includes("lexington") || f.includes("arnold")) return "Lexington JHS / Arnold Complex";
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

function parseDateToTimestamp(rawDate) {
  if (!rawDate) return 0;
  if (rawDate instanceof Date) return rawDate.getTime();
  const parts = String(rawDate).trim().split("/");
  if (parts.length >= 2) {
    const month = parseInt(parts[0], 10) - 1;
    const day = parseInt(parts[1], 10);
    const year = parts[2] ? parseInt(parts[2], 10) : 2026;
    return new Date(year, month, day).getTime();
  }
  return 0;
}

function buildHeaderMap(headers) {
  const map = {
    circuit: -1, date: -1, time: -1, gameId: -1, division: -1,
    type: -1, field: -1, home: -1, away: -1, centerRef: -1, ar1: -1, ar2: -1
  };
  headers.forEach((h, i) => {
    const col = String(h).trim().toLowerCase();
    if (col === "circuit") map.circuit = i;
    else if (col === "date") map.date = i;
    else if (col === "time") map.time = i;
    else if (col.includes("game #") || col === "game") map.gameId = i;
    else if (col === "division") map.division = i;
    else if (col === "type") map.type = i;
    else if (col === "field") map.field = i;
    else if (col.includes("home team") || col === "home") map.home = i;
    else if (col.includes("away team") || col === "away") map.away = i;
    else if (col.includes("center referee") || col === "referee" || col === "cr") map.centerRef = i;
    else if (col === "ar1") map.ar1 = i;
    else if (col === "ar2") map.ar2 = i;
  });
  return map;
}

function transformRawSchedule(rawRows) {
  if (!rawRows || rawRows.length <= 1) return [];

  const headerMap = buildHeaderMap(rawRows[0]);
  const cleanGames = [];

  for (let i = 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!isE154HomeGame(row, headerMap)) continue;

    const gameId = String(row[headerMap.gameId] || "").trim();
    const division = String(row[headerMap.division] || "").trim();
    const rawDate = row[headerMap.date];
    const dateStr = rawDate instanceof Date 
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M/d") 
      : String(rawDate).replace(/^0/, "").split(" ")[0];

    const timeStr = cleanTime(row[headerMap.time]);
    const cleanField = cleanFieldName(row[headerMap.field]);
    const homeCoach = cleanTeamName(row[headerMap.home]);
    const awayCoach = cleanTeamName(row[headerMap.away]);
    const venue = getOperationalVenue(cleanField);

    const centerRef = headerMap.centerRef > -1 ? String(row[headerMap.centerRef] || "").trim() : "";
    const ar1 = headerMap.ar1 > -1 ? String(row[headerMap.ar1] || "").trim() : "";
    const ar2 = headerMap.ar2 > -1 ? String(row[headerMap.ar2] || "").trim() : "";

    const gameObject = {
      gameId: gameId,
      circuit: String(row[headerMap.circuit] || "").trim(),
      dateStr: dateStr,
      timestamp: parseDateToTimestamp(rawDate),
      timeStr: timeStr,
      minutes: parseTimeToMinutes(timeStr),
      division: division,
      cleanField: cleanField,
      venue: venue,
      homeCoach: homeCoach,
      awayCoach: awayCoach,
      matchup: `${homeCoach} vs ${awayCoach}`,
      centerRef: (centerRef && centerRef !== "Unassigned") ? centerRef : "[Open]",
      ar1: ar1,
      ar2: ar2,
      dropdownLabel: `${dateStr} ${timeStr} | ${division} | ${homeCoach} vs ${awayCoach} | ${cleanField} [#${gameId}]`
    };

    cleanGames.push(gameObject);
  }

  // Unified Chronological Sort: Date -> Kick-Off Time -> Field
  cleanGames.sort((a, b) => {
    if (a.timestamp !== b.timestamp) return a.timestamp - b.timestamp;
    if (a.minutes !== b.minutes) return a.minutes - b.minutes;
    return a.cleanField.localeCompare(b.cleanField);
  });

  return cleanGames;
}
