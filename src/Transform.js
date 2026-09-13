/**
 * Data transformation and normalization engine for AYSO Region 154 schedules.
 */

/**
 * Validates whether a raw MatchTrak row represents an active Cypress home match.
 * @param {Array<string>} row - Raw CSV row from MatchTrak.
 * @return {boolean} True if it is an active E154 home game.
 */
function isE154HomeGame(row) {
  if (!row || row.length < 13) return false;
  
  const type = String(row[7] || "").trim().toUpperCase();
  const field = String(row[8] || "").trim();
  const homeTeam = String(row[9] || "").trim();

  // Exclude BYEs
  if (type === "BYE" || field.includes("BYE") || homeTeam.includes("#BYE")) {
    return false;
  }

  // Must have Cypress field prefix in Column I (index 8)
  return field.toUpperCase().startsWith(CONFIG.REGION_PREFIX);
}

/**
 * Cleans lengthy field strings for scannability on mobile screens and forms.
 * Example: "E154-Lexington JHS U10 Field 4 Fall 2026" -> "Lexington JHS Field 4"
 * @param {string} rawField - Raw field string from Column I.
 * @return {string} Cleaned field name.
 */
function cleanFieldName(rawField) {
  if (!rawField) return "";
  return String(rawField)
    .replace(/^E154-/, "")
    .replace(/\s+Fall\s+20\d\d/gi, "")
    .replace(/\s+Fal\s+20\d\d/gi, "")
    .replace(/Elementary School/gi, "ES")
    .replace(/ARTIFICIAL TURF/gi, "Turf")
    .replace(/GRASS Field/gi, "Grass")
    .replace(/U\d+\s+/gi, "") // Strips "U10 ", "U12 " embedded inside field name
    .trim();
}

/**
 * Normalizes coach/team strings and keeps region tags for visiting teams.
 * Examples: 
 *   "11-E154-Richardson_T" -> "Richardson"
 *   "11-Z106-Gastelum_G"   -> "Gastelum (Z106)"
 * @param {string} rawTeam - Raw team string.
 * @return {string} Cleaned coach name with visiting region if applicable.
 */
function cleanTeamName(rawTeam) {
  if (!rawTeam) return "";
  const str = String(rawTeam).trim();
  const parts = str.split("-");
  
  if (parts.length >= 2) {
    const coachRaw = parts[parts.length - 1];
    const coachName = coachRaw.split("_")[0]; // Takes "Richardson" from "Richardson_T"
    
    // Visiting team: extract region code (e.g. Z106, E114, Q97)
    if (!str.includes(CONFIG.REGION_PREFIX)) {
      const regionMatch = str.match(/[A-Z]\d+/);
      const region = regionMatch ? regionMatch[0] : parts[0];
      return `${coachName} (${region})`;
    }
    return coachName;
  }
  return str;
}

/**
 * Converts 24-hour or long time strings into standard 12-hour AM/PM format.
 * Example: "08:00:00 AM" -> "8:00 AM"
 * @param {string|Date} rawTime - Raw time value.
 * @return {string} Formatted time string.
 */
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
 * Builds a compact, uniform dropdown label for Google Forms.
 * Format: "9/19 8:00 AM | BU10 | Zidan vs Mancilla | Lexington JHS Field 7 [#22785]"
 * @param {Object} game - Clean game object.
 * @return {string} Formatted label string.
 */
function buildFormDropdownLabel(game) {
  return `${game.dateStr} ${game.timeStr} | ${game.division} | ${game.homeCoach} vs ${game.awayCoach} | ${game.cleanField} [#${game.gameId}]`;
}

/**
 * Processes an entire 2D array of raw CSV rows into an array of clean game objects.
 * @param {Array<Array<string>>} rawRows - Matrix of CSV records.
 * @return {Array<Object>} Filtered and parsed home game objects.
 */
function transformRawSchedule(rawRows) {
  if (!rawRows || rawRows.length <= 1) return [];

  const cleanGames = [];

  // Start at row 1 to skip CSV header
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

    const gameObject = {
      gameId: gameId,
      circuit: String(row[0] || "").trim(),
      dateStr: dateStr,
      rawDate: rawDate,
      timeStr: timeStr,
      division: division,
      cleanField: cleanField,
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
