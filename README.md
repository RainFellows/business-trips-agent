# 🚄 Business Trips Agent

Google Apps Script automating monthly business trip reports from Google Calendar into Google Sheets. 

**v2.8.0 Features**:
- Support for both **Train** and **Car** travel.
- Automatic **Distance Calculation** via Google Maps API for car trips.
- **Intelligent Main Customer Selection**: Scoring system based on frequency and participants.
- **Vacation/Absence Detection**: Support for multi-day events and intelligent date formatting (date-only for all-day).
- **External Configuration**: Robust management in a "Konfigurace" sheet (with auto-repair).
- **Improved UI**: Auto-resized columns and premium sheet design.

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

Edit `CONFIG` object in `src/Code.js`:
```

| Key | Default | Description |
|-----|---------|-------------|
| `DOMOVSKE_MESTO` | `"Ostrava"` | Home city for trip pairing |
| `HLEDANY_TEXT` | `"vlakem OR autem"` | Calendar search keywords |
| `HODINY_BUFFER` | `1` | Buffer time for train trips (not used for cars) |
| `EMAIL_PRIJEMCE` | Auto-detected | Notification recipient |
| `IGNOROVANE_DOMENY` | List of strings | Domains to exclude from client matching |
