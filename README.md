# AYSO Region 154 MatchTrak Schedule Pipeline

Automated data pipeline and operational engine for AYSO Region 154 (Cypress, CA). Ingests raw multi-circuit CSV exports from MatchTrak, filters for active Cypress home fixtures, and populates downstream operations across Google Forms, Sheets, and regional web pages.

---

## 🏗️ Architecture & Data Flow

---

## 📁 Repository Structure

* `src/Config.js`: Central IDs, season labels, sheet tab names, and email targets.
* `src/Transform.js`: Data normalization, coach/venue sanitization, and regex Game ID embedding.
* `src/Ingest.js`: Dual intake engine checking Gmail attachments and Google Drive folder drops.
* `src/FormSync.js`: Dynamic weekend windowing (Fri–Sun) and Google Form dropdown publisher.
* `src/SetupTakedown.js`: Aggregates first/last kick-offs per field to assign setup & lock-up crews.
* `src/Audit.js`: Reconciles MatchTrak referee rosters against actual Game Day Form submissions.
* `src/WebApi.js`: `doGet(e)` JSON endpoint serving live feeds to public site pages.
* `src/Main.js`: Custom `⚽ AYSO 154 Ops` spreadsheet menu controller.
* `appsscript.json`: Manifest declaring America/Los_Angeles timezone and Google OAuth scopes.

---

## ⚙️ Google Apps Script Triggers

To automate operations hands-free, configure the following installable triggers under **Apps Script > Triggers**:

| Function | Event Source | Type / Interval | Purpose |
| :--- | :--- | :--- | :--- |
| `runScheduleIngest` | Time-driven | Every 10–15 minutes | Ingests new files from Gmail/Drive and syncs Form |
| `menuRunFullSync` | Time-driven | Weekly (Thursday at 8:00 PM) | Prepares full weekend schedule, setup duties, and forms |
| `runRefereeAudit` | Time-driven | Hourly on Saturdays (8 AM – 8 PM) | Live referee check-in and point reconciliation |

---

## 🏟️ Field Mapping Reference (Lexington JHS & Arnold ES)

* **Fields #1, #2, #3:** U5 / U6 Divisions (West side near campus buildings)
* **Fields #4, #5, #6:** U8 Divisions (Center field strip)
* **Fields #7 & #8:** U10 Divisions (South grass facing Orange Ave)
* **Field #9:** U12 Division (Inside Denni St track)
* **Field #10:** U12 Division (Arnold Elementary School grass)
* **Park Lexington Turf & Grass:** U14, U16, U19, and Section 11 Extra fixtures
