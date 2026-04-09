# 🚄 Business Trips Agent

Google Apps Script automating monthly business trip reports from Google Calendar into Google Sheets.

## How it works

1. **Trigger**: Runs automatically on the 1st day of each month (midnight)
2. **Calendar scan**: Reads events from the default calendar matching the keyword `"vlakem"` (by train)
3. **Trip pairing**: Intelligently pairs departure and arrival events using fuzzy matching
4. **Sheet generation**: Creates a formatted sheet named `{Month} - Vlaky` with trip data
5. **Email notification**: Sends an HTML email with a direct link to the generated sheet

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

## Calendar Event Format

The script expects train events in this format:
- **Departure**: `Vlakem z Ostrava do Praha` (case-insensitive)
- **Arrival**: `Vlakem z Praha do Ostrava`
- **Inter-city**: `Vlakem z Brno do Praha` (trips not involving home city)

## Configuration

Edit `CONFIG` object in `src/Code.js`:

| Key | Default | Description |
|-----|---------|-------------|
| `DOMOVSKE_MESTO` | `"Ostrava"` | Home city for trip pairing |
| `HLEDANY_TEXT` | `"vlakem"` | Calendar search keyword |
| `HODINY_BUFFER` | `1` | Hours added before departure / after arrival |
| `EMAIL_PRIJEMCE` | Auto-detected | Notification recipient |
