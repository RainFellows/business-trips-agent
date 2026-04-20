# 🚄 Business Trips Agent

Google Apps Script automating monthly business trip reports from Google Calendar into Google Sheets. 

**v3.1.0 Features**:
- Support for both **Train** and **Car** travel.
- Automatic **Distance Calculation** via Google Maps API for car trips.
- **Intelligent Main Customer Selection**: Scoring system based on frequency and participants.
- **Vacation/Absence Detection**: Premium UI highlighting (light blue), multi-day support, and smart formatting.
- **External Configuration**: Robust management in a "Konfigurace" sheet (with auto-repair).
- **Universal Sheet Naming**: Reports are named `{Month} - Report` for better clarity.
- **Premium UI**: Zebra striping, frozen headers, and automated bold highlighting for keys.

## How it works

1. **Trigger**: Runs automatically on the 1st day of each month (midnight).
2. **Calendar scan**: Reads events matching `"vlakem OR autem"`.
3. **Trip pairing**: Pairs departure and arrival events into round trips.
4. **Client Matching**: Scans meetings during the trip to identify client domains (filtering out internal and common domains).
5. **Sheet generation**: Creates a formatted sheet with trip data, including transport mode and calculated distance.
6. **Email notification**: Sends an HTML email with a direct link to the generated report.

## Project Structure

```
├── .clasp.json          # Clasp deployment config (scriptId, rootDir -> src/)
├── src/
│   ├── Code.js          # Main script logic
│   ├── Checklist.js     # Development checklist (no runtime impact)
│   └── appsscript.json  # Apps Script manifest (timezone, scopes, runtime)
├── project-context.md   # Full project documentation for AI/dev context
└── README.md
```

## Development Workflow

```bash
# Pull latest from Apps Script
clasp pull

# Push local changes to Apps Script
clasp push

# Open Apps Script editor in browser
clasp open
```

The script expects travel events in this format:
- **Train (vlakem)**: `Vlakem z Ostrava do Praha`
- **Car (autem)**: `Autem z Ostrava do Čeladná`
- **Inter-city**: `Vlakem z Brno do Praha`

## Configuration

Since v2.7.0, the script is managed via the **"Konfigurace"** sheet in your Google Spreadsheet. The first time you run the script, this sheet is automatically created with default values.

| Key | Description |
|-----|-------------|
| **Domovské město** | Home city for trip pairing and distance calculation. |
| **Časová rezerva - vlak [hod]** | Buffer time added before/after train trips. |
| **Ignorované domény** | Comma-separated list of domains to exclude from client identification. |
| **Email pro report** | The recipient of the generated trip report notifications. |

To change which calendar events are scanned (e.g., adding keywords for vacations), edit the `HLEDANY_TEXT` variable in `src/Code.js` (for advanced users).
