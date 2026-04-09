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
  - `Session` — getting current user email
- **Deployment**: Container-bound script (lives inside the Google Sheet)
- **Local Dev**: `clasp` CLI for push/pull between local files and Apps Script

## Key Files

| File | Purpose |
|------|---------|
| `src/Code.js` | Main script — all runtime logic (276 lines) |
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

### 3. Buffer Time
`HODINY_BUFFER` (default: 1 hour) is subtracted from departure time and added to arrival time. This accounts for travel to/from the train station.

### 4. Sheet Naming Convention
Sheets are named `{CzechMonthName} - Vlaky` (e.g., `Říjen - Vlaky`). If a sheet with the same name already exists, its content is cleared (not deleted) to preserve the sheet ID for external links.

## Function Reference

| Function | Description |
|----------|-------------|
| `main_generovatCestovniPrikazy()` | Entry point — orchestrates the entire flow |
| `ziskatObdobiMinulehoMesice()` | Calculates previous month date range + Czech name |
| `ziskatCestyZKalendare(start, end)` | Fetches and parses calendar events into trip objects |
| `zpracovatOdjezd(udalost, title, cestyRef, idsRef)` | Handles departure events |
| `zpracovatPrijezd(udalost, title, cestyRef, idsRef)` | Handles arrival events with fuzzy pairing |
| `zpracovatCestuMeziMesty(udalost, title, cestyRef, idsRef)` | Handles inter-city trips |
| `zapsatDoTabulky(ss, data, nazevZaklad)` | Creates/clears sheet and writes formatted data |
| `odeslatNotifikaci(url, mesic, pocetCest)` | Sends HTML email notification |

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

The script searches for events containing `"vlakem"` (case-insensitive) and parses the title:

| Event Title Pattern | Interpretation |
|---|---|
| `Vlakem z Ostrava do Praha` | Departure from home city |
| `Vlakem z Praha do Ostrava` | Arrival to home city |
| `Vlakem z Brno do Praha` | Inter-city trip (not involving home) |

The regex `/(z|do) ([^,]+)/i` extracts city names. The comma acts as a delimiter.

## Configuration Reference

```javascript
const CONFIG = {
  DOMOVSKE_MESTO: "Ostrava",     // Home city — used for pairing logic
  HLEDANY_TEXT: "vlakem",         // Calendar search query
  HODINY_BUFFER: 1,               // Hours buffer before/after train times
  EMAIL_PREDMET: "Cestovní report připraven: ",
  EMAIL_PRIJEMCE: Session.getActiveUser().getEmail(),
  SHEET_HEADER: ["Popis cesty", "Odjezd (Datum a čas)", "Příjezd (Datum a čas)", "Destinace"],
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
4. **Regex fragility**: City name extraction relies on `z {city} do {city}` pattern. Non-standard event titles will be missed or misparsed.
5. **No tests**: GAS APIs (`CalendarApp`, `SpreadsheetApp`) cannot be mocked locally without a testing framework.

## Future Enhancement Ideas

- [ ] Multi-calendar support
- [ ] Configurable transport types (bus, car, etc.)
- [ ] Summary row with totals (trip count, total days)
- [ ] Integration with expense reporting
- [ ] Unit tests with mocked GAS APIs
- [ ] Dry-run mode for testing without writing to sheets
