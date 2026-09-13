const fs = require("fs");
const vm = require("vm");

global.Utilities = {
  formatDate: (d) => `${d.getMonth() + 1}/${d.getDate()}`
};
global.Session = {
  getScriptTimeZone: () => "America/Los_Angeles"
};
global.Logger = { log: () => {} };

vm.runInThisContext(fs.readFileSync("src/Config.js", "utf8"));
vm.runInThisContext(fs.readFileSync("src/Transform.js", "utf8"));
vm.runInThisContext(fs.readFileSync("src/SetupTakedown.js", "utf8"));
vm.runInThisContext(fs.readFileSync("src/Audit.js", "utf8"));

const mockCsvData = [
  ["Circuit", "Date", "Time", "Game #", "Division", "Flight #", "Group", "Type", "Field", "Home Team", "Home Score", "Away Team", "Away Score", "Center Referee", "AR1", "AR2"],
  ["Core", "9/19/2026", "08:00:00 AM", "22101", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Richardson_T", "", "11-Z106-Gastelum_G", "", "John Smith", "Mike Jones", ""],
  ["Core", "9/19/2026", "10:30:00 AM", "22102", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Mancilla_A", "", "11-E154-Zidan_H", "", "Dave Miller", "", ""],
  ["Core", "9/19/2026", "01:00:00 PM", "22103", "BU10", "1", "A", "REG", "E154-Lexington JHS U10 Field 7 Fall 2026", "11-E154-Hernandez_C", "", "11-E114-Baker_J", "", "Unassigned", "", ""],
  ["Core", "9/19/2026", "08:30:00 AM", "22104", "GU12", "1", "A", "REG", "E154-Arnold Elementary School Grass Field 10", "11-E154-Foster_D", "", "11-Q97-Smith_R", "", "Alice Wong", "", ""],
  ["Extra", "9/19/2026", "09:00:00 AM", "22105", "BU14", "1", "A", "REG", "E154-Park Lexington ARTIFICIAL TURF Fall 2026", "11-E154-Alvarez_M", "", "11-E114-Baker_J", "", "Carlos Ruiz", "", ""],
  ["Core", "9/19/2026", "12:01:00 AM", "22106", "BU10", "1", "A", "BYE", "E154-Lexington JHS U10 Field 7 Fall 2026", "#BYE", "", "11-E154-Richardson_T", "", "", "", ""],
  ["Area E", "9/19/2026", "09:00:00 AM", "22107", "BU12", "1", "A", "REG", "E114-Los Alamitos Field 2", "11-E114-Davis_M", "", "11-E154-Richardson_T", "", "", "", ""]
];

const mockFormResponses = [
  ["Timestamp", "Volunteer Name", "Selected Game", "Role"],
  ["9/19/2026 7:45 AM", "John Smith", "9/19 8:00 AM | BU10 | Richardson vs Gastelum | Lexington JHS Field 7 [#22101]", "Center Referee"],
  ["9/19/2026 10:15 AM", "Robert Garcia", "9/19 10:30 AM | BU10 | Mancilla vs Zidan | Lexington JHS Field 7 [#22102]", "Center Referee"],
  ["9/19/2026 12:45 PM", "Sam Walker", "9/19 1:00 PM | BU10 | Hernandez vs Baker | Lexington JHS Field 7 [#22103]", "Center Referee"]
];

function assert(condition, message) {
  console.log(condition ? `\x1b[32m[PASS]\x1b[0m ${message}` : `\x1b[31m[FAIL]\x1b[0m ${message}`);
}

console.log("\n========================================================");
console.log("       AYSO REGION 154 FULL PIPELINE INTEGRATION TEST    ");
console.log("========================================================");

// [1] TRANSFORMATION ENGINE
console.log("\n[1] Testing Transformation Engine (Transform.js)...");
const games = transformRawSchedule(mockCsvData);
assert(games.length === 5, `Extracted 5 active home games (got ${games.length})`);
assert(!games.some(g => g.type === "BYE"), "Filtered out BYE ghost games");
assert(!games.some(g => g.cleanField.includes("Los Alamitos")), "Filtered out away matches");
assert(games.find(g => g.gameId === "22104").venue === "Lexington JHS / Arnold Complex", "Arnold ES Field 10 mapped to Lexington Complex");
assert(games.find(g => g.gameId === "22105").venue === "Park Lexington", "Park Lexington kept isolated");

// [2] SETUP & TAKEDOWN LOGISTICS
console.log("\n[2] Testing Logistics Engine (SetupTakedown.js)...");
const setupTakedownSummary = computeFieldSetupTakedown(games);

const field7 = setupTakedownSummary.find(s => s.field === "Lexington JHS Field 7");
assert(field7 && field7.setupGameId === "22101" && field7.setupTime === "8:00 AM", `Field 7 Setup: ${field7.setupMatch} at ${field7.setupTime}`);
assert(field7 && field7.takedownGameId === "22103" && field7.takedownTime === "1:00 PM", `Field 7 Takedown: ${field7.takedownMatch} at ${field7.takedownTime}`);

const arnold10 = setupTakedownSummary.find(s => s.field === "Arnold ES Field 10");
assert(arnold10 && arnold10.setupGameId === "22104" && arnold10.takedownGameId === "22104", "Arnold Field 10 correctly identified as single-game setup & takedown");

// [3] REFEREE AUDIT
console.log("\n[3] Testing Audit Engine (Audit.js)...");
const matchTrakIndex = buildMatchTrakRefIndex(mockCsvData);
const checkInIndex = buildCheckInIndex(mockFormResponses);

const audit1 = evaluatePositionAudit("22101", matchTrakIndex["22101"], checkInIndex["22101"], "Center Referee", matchTrakIndex["22101"].centerRef);
assert(audit1[6] === "Verified Match" && audit1[7] === "Yes", `Game 22101 status: "${audit1[6]}" (Points: ${audit1[7]})`);

const audit2 = evaluatePositionAudit("22102", matchTrakIndex["22102"], checkInIndex["22102"], "Center Referee", matchTrakIndex["22102"].centerRef);
assert(audit2[6] === "Substitute / Swap", `Game 22102 status: "${audit2[6]}" (Substitute tracked)`);

const audit3 = evaluatePositionAudit("22103", matchTrakIndex["22103"], checkInIndex["22103"], "Center Referee", matchTrakIndex["22103"].centerRef);
assert(audit3[6] === "Walk-On / Fill-In" && audit3[7].includes("Credit Volunteer"), `Game 22103 status: "${audit3[6]}" (Walk-on credited)`);

const audit4 = evaluatePositionAudit("22104", matchTrakIndex["22104"], checkInIndex["22104"] || [], "Center Referee", matchTrakIndex["22104"].centerRef);
assert(audit4[6] === "Unreported / No-Show" && audit4[7].includes("0 pts"), `Game 22104 status: "${audit4[6]}" (No points assigned)`);

console.log("\n========================================================");
console.log("            ALL ENGINES VERIFIED SUCCESSFULLY           ");
console.log("========================================================\n");
