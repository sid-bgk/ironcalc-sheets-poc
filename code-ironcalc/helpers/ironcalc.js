import { Model } from '@ironcalc/nodejs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Template path - local XLSX file
const TEMPLATE_PATH = path.join(__dirname, '../../documents/template.xlsx');

// Sheet indices (0-indexed)
const SHEETS = {
    DSCR_CALC: 0,     // First sheet: DscrCalc
    RENTAL_CALC: 1   // Second sheet: RentalCalc
};

// Column mapping (A=1, B=2, ..., G=7, H=8)
const COL = {
    A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8
};

/**
 * Load template XLSX and return a new Model instance
 * Equivalent to: Google Sheets copy template
 */
const loadTemplate = (name) => {
    if (!fs.existsSync(TEMPLATE_PATH)) {
        throw new Error(`Template not found at: ${TEMPLATE_PATH}`);
    }

    console.log(`Loading template: ${TEMPLATE_PATH}`);
    const model = Model.fromXlsx(TEMPLATE_PATH, "en", "UTC", "en");
    return model;
};

/**
 * Insert data into the spreadsheet
 * Equivalent to: sheets.spreadsheets.values.batchUpdate
 *
 * Input data structure mirrors Google Sheets POC:
 * - RentalCalc!H6:H13 - 8 rental values
 * - DscrCalc!B3 - Gross monthly income
 * - DscrCalc!B5:B13 - 9 expense values
 */
const insertDataIntoSheet = (model, inputData) => {
    // RentalCalc!H6:H13 - Rental values (8 properties)
    const rentalValues = inputData.rentalValues || [
        "$3,200", "$2,700", "$2,850", "$3,500",
        "$2,800", "$2,900", "$3,100", "$4,200"
    ];

    rentalValues.forEach((value, index) => {
        model.setUserInput(SHEETS.RENTAL_CALC, 6 + index, COL.H, value);
    });

    // DscrCalc!B3 - Gross Monthly Income
    const grossIncome = inputData.grossIncome || "$21,155";
    model.setUserInput(SHEETS.DSCR_CALC, 3, COL.B, grossIncome);

    // DscrCalc!B5:B13 - Expense values (9 items)
    const expenseValues = inputData.expenseValues || [
        "$16,000", "$18,500", "$8,500", "$2,500", "$1,800",
        "$800", "$1,200", "$2,100", "$750"
    ];

    expenseValues.forEach((value, index) => {
        model.setUserInput(SHEETS.DSCR_CALC, 5 + index, COL.B, value);
    });

    // Evaluate all formulas after setting inputs
    model.evaluate();
};

/**
 * Fetch calculated results from the spreadsheet
 * Equivalent to: sheets.spreadsheets.values.batchGet
 *
 * Output cells:
 * - DscrCalc!G12 - Total Expenses
 * - DscrCalc!G17 - Net Operating Income
 * - DscrCalc!G29 - Annual NOI
 * - DscrCalc!G31 - Annual Debt Service
 * - DscrCalc!G34 - DSCR Ratio
 */
const fetchDataFromSheet = (model) => {
    const outputRanges = [
        { name: "DscrCalc!G12", sheet: SHEETS.DSCR_CALC, row: 12, col: COL.G },
        { name: "DscrCalc!G17", sheet: SHEETS.DSCR_CALC, row: 17, col: COL.G },
        { name: "DscrCalc!G29", sheet: SHEETS.DSCR_CALC, row: 29, col: COL.G },
        { name: "DscrCalc!G31", sheet: SHEETS.DSCR_CALC, row: 31, col: COL.G },
        { name: "DscrCalc!G34", sheet: SHEETS.DSCR_CALC, row: 34, col: COL.G },
    ];

    const results = {};

    outputRanges.forEach(range => {
        const value = model.getFormattedCellValue(range.sheet, range.row, range.col);
        console.log(`${range.name}:`, value);
        results[range.name] = value;
    });

    return results;
};

/**
 * Save the spreadsheet to an XLSX file
 * Equivalent to: drive.files.export
 */
const saveSpreadsheet = (model, outputPath) => {
    const resultDir = path.join(__dirname, '../../result');

    // Create result directory if it doesn't exist
    if (!fs.existsSync(resultDir)) {
        fs.mkdirSync(resultDir, { recursive: true });
    }

    const fullPath = outputPath || path.join(resultDir, `result_${Date.now()}.xlsx`);
    model.saveToXlsx(fullPath);
    console.log(`Saved to: ${fullPath}`);
    return fullPath;
};

export {
    loadTemplate,
    insertDataIntoSheet,
    fetchDataFromSheet,
    saveSpreadsheet,
    SHEETS,
    COL
};
