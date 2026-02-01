import { Model } from "@ironcalc/nodejs";
import fs from "fs";

const DEFAULT_EXCEL_PATH =
  process.env.EXCEL_FILE_PATH || "./DSCR_NoArrayFormulas_Testing.xlsx";

/**
 * Load an Excel workbook using IronCalc
 * @param {string} filePath - Path to the XLSX file
 * @returns {Model} - IronCalc Model instance
 * @throws {Error} - If file not found or invalid
 */
export function loadWorkbook(filePath = DEFAULT_EXCEL_PATH) {
  console.log(`[workbookLoader] Loading workbook: ${filePath}`);

  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }

  try {
    const model = Model.fromXlsx(filePath, "en", "UTC", "en");
    console.log(`[workbookLoader] Workbook loaded successfully`);
    return model;
  } catch (error) {
    throw new Error(`Failed to load Excel file: ${error.message}`);
  }
}

export { DEFAULT_EXCEL_PATH };
