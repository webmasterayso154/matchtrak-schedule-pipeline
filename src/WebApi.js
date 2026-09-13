/**
 * Web App API endpoint serving JSON feeds of AYSO Region 154 schedules
 * to public website pages (Game Day, Field Setup/Takedown, and Field Marshal).
 */

/**
 * HTTP GET endpoint handler.
 * Supported query parameters:
 *   - view: "gameday" (default), "setup", "grid", or "all"
 *   - date: specific date filter (e.g., "9/19" or "9/19/2026")
 *   - field: optional field substring filter (e.g., "Lexington" or "Turf")
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const view = (params.view || "gameday").toLowerCase();
  const dateFilter = params.date || null;
  const fieldFilter = params.field ? params.field.toLowerCase() : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || 
                      ss.getSheetByName(CONFIG.SHEET_TABS.RAW);

  if (!sourceSheet) {
    return createJsonResponse({ error: "Schedule source sheet not found" }, 404);
  }

  const rawData = sourceSheet.getDataRange().getValues();
  if (rawData.length <= 1) {
    return createJsonResponse({ error: "No schedule records found" }, 200);
  }

  const allGames = transformRawSchedule(rawData);

  // 1. View: Setup & Takedown Summary
  if (view === "setup") {
    const setupSheet = ss.getSheetByName(CONFIG.SHEET_TABS.SETUP_TAKEDOWN);
    if (!setupSheet || setupSheet.getLastRow() <= 1) {
      return createJsonResponse({ error: "Setup/takedown data not generated yet" }, 200);
    }
    const setupData = setupSheet.getDataRange().getValues();
    const headers = setupData[0];
    const rows = setupData.slice(1).map(r => {
      const item = {};
      headers.forEach((h, idx) => {
        item[toCamelCase(h)] = r[idx];
      });
      return item;
    });
    return createJsonResponse({ status: "success", count: rows.length, data: rows });
  }

  // 2. View: Game Day & General Schedule Filter
  let filteredGames = allGames;

  if (dateFilter) {
    filteredGames = filteredGames.filter(g => g.dateStr === dateFilter || g.dateStr.startsWith(dateFilter));
  } else if (view !== "all") {
    // Default to active weekend window
    const activeDates = getActiveWeekendDates();
    filteredGames = filteredGames.filter(g => 
      activeDates.some(ad => g.dateStr === ad || g.dateStr.startsWith(ad))
    );
  }

  if (fieldFilter) {
    filteredGames = filteredGames.filter(g => g.cleanField.toLowerCase().includes(fieldFilter));
  }

  // Sort chronologically
  filteredGames.sort((a, b) => {
    const timeA = new Date(a.rawDate + " " + a.timeStr).getTime() || 0;
    const timeB = new Date(b.rawDate + " " + b.timeStr).getTime() || 0;
    return timeA - timeB || a.cleanField.localeCompare(b.cleanField);
  });

  // 3. View: Field Marshal Matrix (Hourly Field Grid)
  if (view === "grid") {
    const timeSlots = [...new Set(filteredGames.map(g => g.timeStr))];
    const fields = [...new Set(filteredGames.map(g => g.cleanField))].sort();

    const grid = timeSlots.map(time => {
      const slotRow = { time: time };
      fields.forEach(field => {
        const match = filteredGames.find(g => g.timeStr === time && g.cleanField === field);
        slotRow[field] = match ? `${match.division}: ${match.homeCoach} vs ${match.awayCoach}` : "—";
      });
      return slotRow;
    });

    return createJsonResponse({
      status: "success",
      fields: fields,
      timeSlots: timeSlots,
      grid: grid
    });
  }

  // Default output: Normalized Game Day Array
  return createJsonResponse({
    status: "success",
    count: filteredGames.length,
    games: filteredGames
  });
}

/**
 * Serializes data into a JSON response with proper MIME headers.
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Converts header strings into JSON-friendly camelCase keys.
 */
function toCamelCase(str) {
  return String(str)
    .replace(/[^a-zA-Z0-9 ]/g, "")
    .toLowerCase()
    .split(" ")
    .map((word, i) => i === 0 ? word : word.charAt(0).toUpperCase() + word.slice(1))
    .join("");
}
