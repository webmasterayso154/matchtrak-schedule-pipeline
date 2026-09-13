/**
 * Synchronizes cleaned AYSO Region 154 home match schedules
 * with the Game Day check-in Google Form.
 */

/**
 * Main Form sync routine. Reads processed games, determines the active weekend
 * window, and pushes formatted match choices into the Form dropdown.
 */
function syncFormDropdown() {
  if (!CONFIG.FORM_ID || CONFIG.FORM_ID.includes("YOUR_")) {
    Logger.log("Form ID not set in Config.js. Skipping Form sync.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_TABS.REF_DATA) || ss.getSheetByName(CONFIG.SHEET_TABS.RAW);

  if (!sourceSheet) {
    Logger.log("Source data sheet not found for Form sync.");
    return;
  }

  const rawData = sourceSheet.getDataRange().getValues();
  if (rawData.length <= 1) {
    Logger.log("No data available in source sheet to populate Form.");
    return;
  }

  // 1. Transform raw matrix into normalized game objects
  const allGames = transformRawSchedule(rawData);
  if (allGames.length === 0) {
    Logger.log("No valid Cypress home games detected.");
    return;
  }

  // 2. Identify the active weekend window (Friday - Sunday)
  const targetDateStrings = getActiveWeekendDates();
  Logger.log(`Active date filter: ${targetDateStrings.join(", ")}`);

  // 3. Filter and sort matches for this weekend
  const weekendGames = allGames.filter(game => {
    // Matches against standard "M/d" or "M/d/yyyy" strings
    return targetDateStrings.some(td => game.dateStr === td || game.dateStr.startsWith(td));
  });

  // Sort chronologically by kick-off time and field
  weekendGames.sort((a, b) => {
    const timeA = new Date(a.rawDate + " " + a.timeStr).getTime() || 0;
    const timeB = new Date(b.rawDate + " " + b.timeStr).getTime() || 0;
    return timeA - timeB || a.cleanField.localeCompare(b.cleanField);
  });

  // 4. Construct dropdown array
  const choices = weekendGames.map(g => g.dropdownLabel);

  // 5. Push choices to Google Form
  updateFormListItem(CONFIG.FORM_ID, "Select Your Game", choices);
}

/**
 * Calculates upcoming match dates (Friday, Saturday, Sunday) relative to today.
 * If run Monday-Thursday, targets the upcoming weekend.
 * If run Friday-Sunday, targets the current weekend.
 * @return {Array<string>} Array of date strings in "M/d" format.
 */
function getActiveWeekendDates() {
  const today = new Date();
  const currentDayOfWeek = today.getDay(); // 0 = Sun, 5 = Fri, 6 = Sat
  
  // Calculate distance to this week's Friday
  let daysUntilFriday = 5 - currentDayOfWeek;
  if (currentDayOfWeek === 0) {
    // If today is Sunday, keep the weekend window from Friday through today
    daysUntilFriday = -2;
  } else if (daysUntilFriday < 0) {
    // If today is Saturday (6), Friday was yesterday (-1)
    daysUntilFriday = -1;
  }

  const friday = new Date(today);
  friday.setDate(today.getDate() + daysUntilFriday);

  const dates = [];
  for (let i = 0; i < 3; i++) {
    const matchDate = new Date(friday);
    matchDate.setDate(friday.getDate() + i);
    dates.push(`${matchDate.getMonth() + 1}/${matchDate.getDate()}`);
  }

  return dates;
}

/**
 * Targets a specific Dropdown/ListItem in the Form and applies values.
 * @param {string} formId - Google Form ID.
 * @param {string} itemTitle - Title of the dropdown question.
 * @param {Array<string>} choices - Array of string values.
 */
function updateFormListItem(formId, itemTitle, choices) {
  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.LIST);
  let targetItem = null;

  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle().trim().toLowerCase() === itemTitle.trim().toLowerCase()) {
      targetItem = items[i].asListItem();
      break;
    }
  }

  if (!targetItem) {
    Logger.log(`Warning: Dropdown question "${itemTitle}" not found in Form.`);
    return;
  }

  if (choices.length > 0) {
    targetItem.setChoiceValues(choices);
    Logger.log(`Updated "${itemTitle}" with ${choices.length} games.`);
  } else {
    targetItem.setChoiceValues(["No games scheduled for this weekend"]);
    Logger.log(`No games matched. Set default placeholder.`);
  }
}
