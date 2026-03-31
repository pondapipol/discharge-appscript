# Discharge Counseling Tracker

A Google Apps Script web application for tracking hospital discharge counseling sessions, discharge counts, and prescribing errors. Built as a container-bound script deployed from Google Sheets.

## Features

- **Counseling Form** - Record patient discharge counseling with hierarchical checkboxes for 36 medication/technique categories
- **Discharge Count (DC)** - Track daily patient discharge numbers (upsert: one record per date)
- **Prescribing Error (PE)** - Track daily prescribing error counts (upsert: one record per date)
- **Dashboard** - Interactive charts and summary table with month selector
  - Top-level category comparison (horizontal bar)
  - Detailed drug breakdown under category 1 (horizontal bar)
  - Daily discharge counts (horizontal bar)
  - Prescribing errors - last 6 months (horizontal bar)
- **Report Summary** - Auto-calculating spreadsheet with formulas (no manual aggregation needed)

## Setup

### Prerequisites
- A Google Spreadsheet
- Access to Apps Script (Extensions > Apps Script)

### Installation

1. Open your Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Create the following files in the Apps Script editor:

| File Name | Type |
|-----------|------|
| `Code` | Script (.gs) |
| `Index` | HTML |
| `Stylesheet` | HTML |
| `Page_Counseling` | HTML |
| `Page_DC` | HTML |
| `Page_PE` | HTML |
| `Page_Dashboard` | HTML |
| `JavaScript` | HTML |

4. Copy-paste content from each local file into the corresponding Apps Script file
5. In the Apps Script editor, select `setupSheets` from the function dropdown and click **Run**
   - This creates `Counseling_Raw`, `DC_Raw`, `PE_Raw` sheets with headers
   - This creates `Report Summary` sheet with auto-calculating formulas
6. Go to **Deploy > New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (or your organization)
7. Click **Deploy** and authorize when prompted

### Updating

After code changes:
1. Update the files in Apps Script editor
2. Go to **Deploy > Manage deployments**
3. Click the edit (pencil) icon on your deployment
4. Set version to **New version**
5. Click **Deploy**

## Data Structure

### Counseling_Raw
Each row represents one patient counseling session with 39 columns:
- `Timestamp` - Auto-generated server timestamp
- `CounselingDate` - Date of counseling
- `AN` - 9-digit admission number (stored as text)
- 36 boolean columns for each counseling item

### DC_Raw
One row per date (upserted on duplicate dates):
- `Timestamp`, `DischargeDate`, `DischargeCount`

### PE_Raw
One row per date (upserted on duplicate dates):
- `Timestamp`, `ErrorDate`, `ErrorCount`

### Report Summary
Formula-driven sheet. Change the date in cell **B1** to view different months. All values auto-calculate from the raw data sheets.

To regenerate: run `setupReportSummary()` from the Apps Script editor.

## Tech Stack

- **Google Apps Script** (server-side)
- **Chart.js** (CDN) - dashboard visualizations
- **Font Awesome 6** (CDN) - icons
- **Google Fonts: Athiti** - Thai text rendering

## Color Palette

| Color | Hex | Usage |
|-------|-----|-------|
| Honey | `#FEB737` | Primary actions, active nav |
| Periwinkle | `#A7B5FE` | Navbar, borders, secondary |
| Sea Breeze | `#7BB8C0` | Section headers, success |
