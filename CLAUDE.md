# CLAUDE.md - Project Context for AI Assistants

## Project Overview
A **Google Apps Script** container-bound web app for tracking hospital **Discharge Counseling**, **Discharge Counts**, and **Prescribing Errors**. Deployed from within a Google Sheet via Extensions > Apps Script.

## Architecture
- **Container-bound** to Google Sheets (uses `SpreadsheetApp.getActiveSpreadsheet()`, NOT `openById`)
- **SPA-style** web app served via `HtmlService` with `doGet()` entry point
- **4 pages** in a single HTML shell with CSS `display` toggle: Dashboard, Counseling, DC, PE
- **Chart.js** for all dashboard charts (all horizontal bar charts)
- **Google Fonts: Athiti** for Thai text rendering

## File Structure
```
Code.gs              # All server-side logic (doGet, CRUD, aggregation, sheet setup)
Index.html           # Main HTML shell with navbar, page containers, CDN links
Stylesheet.html      # All CSS (included via <?!= include('Stylesheet') ?>)
Page_Counseling.html # Hierarchical checkbox form (36 checkboxes)
Page_DC.html         # Discharge count form (date + numeric)
Page_PE.html         # Prescribing error form (date + numeric)
Page_Dashboard.html  # Month selector, stat cards, summary table, 4 chart canvases
JavaScript.html      # All client-side JS (included via <?!= include('JavaScript') ?>)
```

## Key Design Decisions
- **HTML files are included via `<?!= include('filename') ?>`** scriptlet tags in Index.html. The `include()` function is defined in Code.gs.
- **COUNSELING_FIELDS array** in Code.gs is the single source of truth for checkbox data-keys, labels, categories, and subgroups. The HTML `data-key` attributes must match the `key` values in this array.
- **DC and PE use upsert logic**: if a row exists for the same date, update the count; otherwise append.
- **Report Summary** uses spreadsheet formulas (COUNTIFS, SUMIFS, UNIQUE, FILTER) referencing raw sheets. Cell B1 is the month selector date. Run `setupReportSummary()` to regenerate.
- **AN (Admission Number)** is validated as exactly 9 digits and stored as string (text format) to preserve leading zeros.

## Data Sheets
| Sheet | Purpose | Key Columns |
|-------|---------|-------------|
| `Counseling_Raw` | One row per patient counseling | Timestamp, CounselingDate, AN, 36 boolean columns |
| `DC_Raw` | One row per date (upsert) | Timestamp, DischargeDate, DischargeCount |
| `PE_Raw` | One row per date (upsert) | Timestamp, ErrorDate, ErrorCount |
| `Report Summary` | Formula-driven monthly report | B1=month selector, all counts via COUNTIFS/SUMIFS |

## Color Palette (use only these 3)
- **Honey** `#FEB737` - primary actions, active nav, chart accent
- **Periwinkle** `#A7B5FE` - navbar, card borders, secondary elements
- **Sea Breeze** `#7BB8C0` - section headers, success states, hover

## Common Tasks

### Adding a new drug/checkbox item
1. Add entry to `COUNSELING_FIELDS` in Code.gs (with key, label, category, subgroup)
2. Add corresponding `<div class="checkbox-item">` in Page_Counseling.html with matching `data-key`
3. Run `setupSheets()` to update Counseling_Raw headers (this will NOT delete existing data, but the new column will be added to new rows only - consider migrating)
4. Run `setupReportSummary()` to regenerate formulas

### Changing labels
- Update in BOTH Code.gs `COUNSELING_FIELDS` and the corresponding HTML file
- JavaScript.html `TOP_LEVEL_LABELS` object must also match for dashboard charts

### Regenerating Report Summary
- Run `setupReportSummary()` from the Apps Script editor. This DELETES and recreates the Report Summary sheet with fresh formulas.

### First-time setup on a new spreadsheet
1. Open the sheet > Extensions > Apps Script
2. Create all files (Code.gs + 7 HTML files) with matching names
3. Run `setupSheets()` (creates raw data sheets + Report Summary with formulas)
4. Deploy > New deployment > Web app

## Important Constraints
- `google.script.run` does NOT support Promises - uses `withSuccessHandler`/`withFailureHandler`
- Date objects passed via `google.script.run` are serialized - pass as ISO strings when possible
- Included HTML files (`<?!= include() ?>`) cannot contain scriptlet tags themselves
- Apps Script has a 6-minute execution timeout per function call
- `sheet.appendRow()` is atomic for concurrent access
- Thai Buddhist Era year = Gregorian + 543 (display only, dates stored in Gregorian)

## Testing
- After deployment, test form submissions and verify data appears in raw sheets
- Change B1 date in Report Summary to check formulas recalculate
- Dashboard charts require at least one data row to display
