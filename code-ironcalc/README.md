# IronCalc Financial Calculator

Local spreadsheet calculation engine for DSCR (Debt Service Coverage Ratio) and rental income calculations.

## Overview

This project uses [IronCalc](https://www.ironcalc.com/) to perform financial calculations locally, without requiring Google Sheets API or any cloud services.

## Features

- Load XLSX templates with formulas
- Insert input data programmatically
- Evaluate all formulas locally
- Read calculated results
- Export to XLSX

## Prerequisites

- Node.js 18+
- npm

## Installation

```bash
npm install
```

## Setup

1. Place your template file at `../documents/template.xlsx`
2. Ensure the template has two sheets:
   - `DscrCalc` (index 0) - DSCR calculations
   - `RentalCalc` (index 1) - Rental calculations

## Usage

### Run with default data

```bash
npm start
```

### Programmatic usage

```javascript
import { loadTemplate, insertDataIntoSheet, fetchDataFromSheet, saveSpreadsheet } from './helpers/ironcalc.js';

// Load template
const model = loadTemplate("calculation_1");

// Insert custom data
insertDataIntoSheet(model, {
    rentalValues: ["$3,200", "$2,700", "$2,850", "$3,500", "$2,800", "$2,900", "$3,100", "$4,200"],
    grossIncome: "$21,155",
    expenseValues: ["$16,000", "$18,500", "$8,500", "$2,500", "$1,800", "$800", "$1,200", "$2,100", "$750"]
});

// Get results
const results = fetchDataFromSheet(model);
console.log(results);

// Save output
saveSpreadsheet(model);
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

Results are saved to `../result/result_<timestamp>.xlsx`

## API Reference

| Method | Description |
|--------|-------------|
| `loadTemplate(name)` | Load XLSX template |
| `insertDataIntoSheet(model, data)` | Set input values and evaluate |
| `fetchDataFromSheet(model)` | Read calculated results |
| `saveSpreadsheet(model, path?)` | Export to XLSX |

## Dependencies

- [@ironcalc/nodejs](https://www.npmjs.com/package/@ironcalc/nodejs) - Local spreadsheet engine

## License

ISC
