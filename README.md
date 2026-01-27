# Iron Calc POC

Proof of concept for spreadsheet-based financial calculations (DSCR & Rental Income).

## Overview

This project explores two approaches for automating financial calculations using spreadsheet templates:

| Approach | Folder | Status | Recommendation |
|----------|--------|--------|----------------|
| **IronCalc** (Local) | `code-ironcalc/` | Working | Recommended |
| **Google Sheets API** | `code/` | Working | Legacy |

## Why IronCalc?

| Aspect | Google Sheets API | IronCalc |
|--------|-------------------|----------|
| Network | Required | None (local) |
| Quotas | 15GB limit | Unlimited |
| Setup | Complex (credentials, sharing) | `npm install` |
| Speed | ~3-6s per calculation | ~25-120ms |
| Cost | Potential API costs | Free |
| Offline | No | Yes |

## Project Structure

```
iron_calc_poc/
├── code-ironcalc/       # IronCalc implementation (recommended)
│   ├── helpers/
│   │   └── ironcalc.js  # IronCalc API wrapper
│   ├── run_script.js    # Main entry point
│   └── package.json
├── code/                # Google Sheets implementation (legacy)
│   ├── helpers/
│   │   └── google.js    # Google API wrapper
│   ├── run_script.js    # Main entry point
│   └── package.json
├── documents/
│   └── template.xlsx    # Spreadsheet template with formulas
└── result/              # Generated output files
```

## Getting Started

### Option 1: IronCalc (Recommended)

1. **Clone the repo**
   ```bash
   git clone <repo-url>
   cd iron_calc_poc
   ```

2. **Add your template**
   - Place your Excel template at `documents/template.xlsx`
   - Template must have `DscrCalc` and `RentalCalc` sheets with formulas

3. **Install dependencies**
   ```bash
   cd code-ironcalc
   npm install
   ```

4. **Run**
   ```bash
   npm start
   ```

5. **Check results**
   - Console shows calculated values
   - Output saved to `result/result_<timestamp>.xlsx`

### Option 2: Google Sheets API (Legacy)

1. **Google Cloud Setup**
   - Create project at [console.cloud.google.com](https://console.cloud.google.com/)
   - Enable Google Drive API & Google Sheets API
   - Create Service Account & download credentials

2. **Configure**
   - Place `credentials.json` in `code/` folder
   - Update IDs in `code/helpers/google.js`
   - Share template with service account email

3. **Install & Run**
   ```bash
   cd code
   npm install
   npm start
   ```

## Template Requirements

The template spreadsheet must have:

**Sheet 1: DscrCalc**
- Input: B3 (Gross Income), B5:B13 (Expenses)
- Output: G12, G17, G29, G31, G34 (Calculated values)

**Sheet 2: RentalCalc**
- Input: H6:H13 (Rental values)

## Documentation

See `_bmad-output/implementation-artifacts/` for detailed documentation:
- `ironcalc-poc-documentation.md` - Full IronCalc documentation
- `iron-calc-poc-analysis.md` - Code analysis and architecture

## Tech Stack

- **Runtime:** Node.js 18+
- **IronCalc:** [@ironcalc/nodejs](https://www.npmjs.com/package/@ironcalc/nodejs)
- **Google APIs:** [googleapis](https://www.npmjs.com/package/googleapis) (legacy)

## License

ISC
