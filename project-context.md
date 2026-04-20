# Project Context: Business Trips Agent

> This file provides comprehensive context for AI assistants and developers working on this project.

## Overview

This is a **Google Apps Script** project that automates the generation of monthly business trip reports. It runs as a container-bound script inside a Google Sheets spreadsheet.

## Architecture

```
Google Calendar (read-only)
       │
       ▼
┌─────────────────────────┐
│   Apps Script (Code.js) │
│                         │
│  1. Read calendar events│
│  2. Parse & pair trips  │
│  3. Write to Sheets     │
│  4. Send notification   │
└─────────────────────────┘
       │              │
       ▼              ▼
Google Sheets    Gmail (notification)
```

## Technology Stack

- **Runtime**: Google Apps Script (V8 engine)
- **Language**: JavaScript (ES6+, transpiled by V8 runtime)
- **APIs Used**:
  - `CalendarApp` — reading calendar events
  - `SpreadsheetApp` — creating/formatting sheets
  - `MailApp` — sending HTML email notifications
  - `Maps` (DirectionFinder) — automatic distance calculation for car trips
  - `Session` — getting current user email
- **v2.9.0 Features**:
- Support for both **Train** and **Car** travel.
- Automatic **Distance Calculation** via Google Maps API for car trips.
- **Intelligent Main Customer Selection**: Scoring system based on frequency and participants.
- **Vacation/Absence Detection**: Support for multi-day events and intelligent date formatting (date-only for all-day).
- **External Configuration**: Robust management in a "Konfigurace" sheet (with auto-repair).
- **Improved UI**: Auto-resized columns and premium sheet design.
- **Deployment**: Container-bound script (lives inside the Google Sheet)
- **Local Dev**: `clasp` CLI for push/pull between local files and Apps Script

## Key Files

| File | Purpose |
|------|---------|
| `src/Code.js` | Main script — all runtime logic (290 lines) |
| `src/appsscript.json` | Manifest: timezone, OAuth scopes, runtime version |
| `.clasp.json` | Clasp config: scriptId, rootDir=`src` |

## Key Design Decisions

### 1. Container-bound vs Standalone
The script is **container-bound** (attached to a specific Google Sheet). This means:
- `SpreadsheetApp.getActiveSpreadsheet()` works without hardcoding IDs
- The script has inherent access to the parent spreadsheet
- Triggers run in the context of the sheet owner

### 2. Trip Pairing Logic
The core algorithm pairs departure and arrival events:
- Events are sorted chronologically
- A departure ("z Ostrava") creates an open trip
- An arrival ("do Ostrava") closes the most recent open trip
- Unmatched arrivals are logged as "missing departure"
- Uses **fuzzy matching** — doesn't require exact city name match

### 3. Transport Types & Buffer Time
- **Train**: Keyword `vlakem`. Uses `HODINY_BUFFER` (loaded from "Konfigurace") before and after trip for station transfers.
- **Car**: Keyword `autem`. No buffer time added. Automatically calculates distance using Google Maps (Driving Mode).

### 4. External Configuration (v2.9.0)
The script uses a dedicated tab **"Konfigurace"** to manage settings without touching the code.
- **Parameters**:
    - `Domovské město`: Used for travel pairing and KM calculation.
    - `Časová rezerva - vlak [hod]`: Buffer for train trips.
    - `Ignorované domény`: Comma-separated list of domains to skip during client identification.
    - `Email pro report`: Recipient of the notification.
- **Auto-Initialization**: If the sheet is missing, the script creates it with default values and a premium design.

### 4. Client Identification (Scoring System)
The script identifies and ranks clients based on calendar events overlapping with the trip timeframe.
- **Scoring Algorithm**:
    - Each domain gets a score: `(Event Count * 10) + Unique Participants Count`.
    - Domains are sorted by score descending.
    - **Output**: The winner is displayed in UPPERCASE. Up to two other domains are listed as "others".
- **Filtering**: Only meetings where the user is `OWNER` or status is `YES`.
- **Exclusion**:
    - **Domains**: Blacklist in `CONFIG.IGNOROVANE_DOMENY`.
    - **Keywords**: Events with titles containing "Zrušeno", "Canceled", etc., are skipped.
    - **Status Override**: If the user's status is explicitly `NO` (declined), the event is skipped even if they are the `OWNER`.

### 5. City Name Standardization
Helper function `formatovatMesto()` ensures consistent naming (e.g., `praha hl.n.` -> `Praha`). It title-cases the name and strips technical suffixes.

### 6. Vacation & Absence Detection
The script scans for keywords `"vacation OR dovolená"`.
- **v2.9.0 Logic**:
    - Each calendar event becomes one row (respecting multi-day ranges).
    - **Full Day Visibility**: All-day events (or >=8h) are automatically detected and formatted.
    - **Formatting**: All-day entries are formatted as `dd.MM.yyyy` (date only), while timed entries include `HH:mm`.
- **Reporting**: Entries are inserted chronologically into the main report. Destination column for vacations is left as "-".

### 7. Sheet Naming & Structure
Sheets are named `{CzechMonthName} - Vlaky`. New "Doprava" (Transport) column added in v2.4.0. Columns are auto-resized for optimal fit.

## Function Reference

| Function | Description |
|----------|-------------|
| `main_generovatCestovniPrikazy()` | Entry point — orchestrates the entire flow |
| `ziskatObdobiMinulehoMesice()` | Calculates previous month date range + Czech name |
| `ziskatCestyZKalendare(start, end)` | Fetches and parses calendar events into trip objects |
| `zpracovatOdjezd(udalost, title, cestyRef, idsRef)` | Handles departure events |
| `zpracovatPrijezd(udalost, title, cestyRef, idsRef)` | Handles arrival events with fuzzy pairing |
| `zpracovatCestuMeziMesty` | Handles inter-city trips |
| `ziskatKlientaZPrekryvu` | Analyzes overlapping events and returns unique client domains |
| `ziskatKm(start, end)` | Calls Google Maps API for distance calculation |
| `formatovatMesto(text)` | Normalizes city names (capitalization, cleaning) |
| `zapsatDoTabulky` | Creates/clears sheet and writes formatted data |
| `odeslatNotifikaci` | Sends HTML email notification |

## Trigger Configuration

- **Function**: `main_generovatCestovniPrikazy`
- **Type**: Time-driven
- **Frequency**: Monthly
- **Day**: 1st day of month
- **Time**: Midnight to 1 AM (Europe/Prague timezone)

## OAuth Scopes

| Scope | Purpose |
|-------|---------|
| `calendar.readonly` | Read calendar events |
| `spreadsheets` | Create and format sheets |
| `script.send_mail` | Send email notifications |
| `userinfo.email` | Get current user's email for notifications |

## Calendar Event Format

The script searches for events containing `"vlakem OR autem"` (case-insensitive) and parses the title:

| Event Title Pattern | Interpretation |
|---|---|
| `Vlakem z Ostrava do Praha` | Train Trip |
| `Autem z Ostrava do Čeladná` | Car Trip (distance calc) |

## Configuration Reference

```javascript
const CONFIG = {
  DOMOVSKE_MESTO: "Ostrava",     // Home city — used for pairing logic
  HLEDANY_TEXT: "vlakem",         // Calendar search query
  HODINY_BUFFER: 1,               // Hours buffer before/after train times
  EMAIL_PREDMET: "Cestovní report připraven: ",
  EMAIL_PRIJEMCE: Session.getActiveUser().getEmail(),
  SHEET_HEADER: ["Popis cesty", "Odjezd", "Příjezd", "Destinace", "Doprava", "km autem", "Klient"],
  IGNOROVANE_DOMENY: [...],                 // Blacklist for client matching
  COLORS: {
    HEADER_BG: "#4c1130",   // Dark burgundy
    HEADER_TEXT: "#ffffff",
    ROW_BANDING: "#f3f3f3",
    BORDER: "#000000"
  }
};
```

## Development Workflow

1. **Pull latest**: `clasp pull` (downloads from Apps Script to `src/`)
2. **Edit locally**: Modify files in `src/`
3. **Push changes**: `clasp push` (uploads `src/` to Apps Script)
4. **Test**: Run `main_generovatCestovniPrikazy` from Apps Script editor or `clasp run`
5. **Commit**: `git add . && git commit -m "description"`
6. **Deploy**: Changes take effect immediately after `clasp push` (no build step)

### Important: `.clasp.json` Configuration
- `rootDir` is set to `"src"` — only files in `src/` are pushed to Apps Script
- `scriptId` points to the production Apps Script project
- **Never commit `.clasprc.json`** (contains auth tokens)

## Apps Script Links

- **Script Editor**: https://script.google.com/u/0/home/projects/1T1l5R7t42X6HgGIYGcMPXAKDAFhqHxGz6NMhy0EpVBOrUVoCXHQFmeVd/edit
- **Parent Spreadsheet**: Linked via container binding (no hardcoded ID)

## Known Limitations

1. **Single calendar only**: Reads from the default calendar. Multi-calendar support not implemented.
2. **Czech locale dependency**: Month names are generated in Czech via `toLocaleString('cs-CZ')`. The V8 runtime supports this, but behavior may vary.
3. **No duplicate detection**: If the trigger runs twice in a month, the sheet is overwritten (cleared + rewritten), which is safe but loses manual edits.
4. **Regex fragility**: City name extraction relies on `z {city} do {city}` pattern.
5. **Ghost Events**: API may return events delete or moved instances from recurring series. Tracing logging in `ziskatKlientaZPrekryvu` is used to identify these cases manually.

## Future Enhancement Ideas

- [x] Configurable transport types (vlak/auto)
- [x] Dry-run mode for testing (`dev_dryRunTest`)
- [ ] Summary row with totals (total km, etc.)
- [ ] Multi-calendar support
- [ ] Unit tests with mocked GAS APIs
