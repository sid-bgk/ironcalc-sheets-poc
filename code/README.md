# Google Sheets Financial Calculator

Google Sheets API-based calculation engine for DSCR (Debt Service Coverage Ratio) and rental income calculations.

> **Note:** Consider using the [IronCalc version](../code-ironcalc/) instead - it's faster, has no quota limits, and works offline.

## Overview

This project uses Google Sheets API to perform financial calculations by:
1. Copying a template spreadsheet
2. Inserting input data
3. Reading calculated results
4. Exporting to XLSX

## Features

- Copy Google Sheets templates
- Batch update cell values
- Batch read calculated results
- Export to XLSX format
- Delete temporary sheets

## Prerequisites

- Node.js 18+
- npm
- Google Cloud Project with:
  - Google Drive API enabled
  - Google Sheets API enabled
  - Service Account with credentials

## Installation

```bash
npm install
```

## Setup

### 1. Google Cloud Setup

1. Create project at [Google Cloud Console](https://console.cloud.google.com/)
2. Enable Google Drive API
3. Enable Google Sheets API
4. Create Service Account
5. Download credentials JSON

### 2. Configuration

1. Place `credentials.json` in this folder
2. Update `helpers/google.js`:
   - Line 15: Set `templateSheetID`
   - Line 49: Set output folder ID in `parents`

### 3. Template Setup

1. Create Google Sheet with formulas
2. Share with service account email (Editor access)
3. Share output folder with service account

## Usage

```bash
npm start
```

## Input Cells

| Range | Sheet | Description |
|-------|-------|-------------|
| H6:H13 | RentalCalc | 8 rental property values |
| B3 | DscrCalc | Gross Monthly Income |
| B5:B13 | DscrCalc | 9 expense line items |

## Output Cells

| Cell | Sheet | Description |
|------|-------|-------------|
| G12 | DscrCalc | Total Expenses |
| G17 | DscrCalc | Net Operating Income |
| G29 | DscrCalc | Annual NOI |
| G31 | DscrCalc | Annual Debt Service |
| G34 | DscrCalc | DSCR Ratio |

## Output

Results are saved to `../result/<spreadsheetId>.xlsx`

## API Functions

| Function | Description |
|----------|-------------|
| `uploadInitialSheet(name)` | Copy template sheet |
| `insertDataIntoSheet(id)` | Batch update values |
| `fetchDataFromSheet(id)` | Batch read results |
| `downloadSpreadsheet(id)` | Export to XLSX |
| `deleteSpreadsheetByID(id)` | Delete sheet |

## Known Issues

- **Quota limits:** Service accounts have 15GB storage limit
- **Network dependency:** Requires internet connection
- **Setup complexity:** Requires Google Cloud configuration

## Dependencies

- [googleapis](https://www.npmjs.com/package/googleapis) - Google APIs client

## License

ISC
