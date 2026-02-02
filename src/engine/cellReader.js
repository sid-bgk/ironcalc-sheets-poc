/**
 * Get the formatted display value of a cell
 * @param {Model} model - IronCalc Model instance
 * @param {number} sheet - Sheet index (0-indexed)
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed, A=1)
 * @returns {string} - Formatted cell value
 */
export function getCellValue(model, sheet, row, col) {
  return model.getFormattedCellValue(sheet, row, col);
}

/**
 * Get the raw content of a cell (formula string or literal value)
 * @param {Model} model - IronCalc Model instance
 * @param {number} sheet - Sheet index (0-indexed)
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed, A=1)
 * @returns {string} - Raw cell content (formulas start with "=")
 */
export function getCellRawContent(model, sheet, row, col) {
  return model.getCellContent(sheet, row, col);
}

/**
 * Check if a cell contains a formula
 * @param {Model} model - IronCalc Model instance
 * @param {number} sheet - Sheet index (0-indexed)
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed, A=1)
 * @returns {boolean} - True if cell contains a formula
 */
export function isFormulaCell(model, sheet, row, col) {
  const content = model.getCellContent(sheet, row, col);
  return content ? content.startsWith("=") : false;
}

/**
 * Coerce a string value from IronCalc to appropriate JS type
 * @param {string} value - String value from getFormattedCellValue()
 * @returns {any} - Coerced value (number, boolean, null, or string)
 */
export function coerceValue(value) {
  // Empty or null → null
  if (value === '' || value === null || value === undefined) {
    return null;
  }

  // Excel error values → null
  if (value.startsWith('#') && (
    value === '#N/A' ||
    value === '#REF!' ||
    value === '#VALUE!' ||
    value === '#DIV/0!' ||
    value === '#NAME?' ||
    value === '#NULL!' ||
    value === '#NUM!'
  )) {
    return null;
  }

  // Boolean values
  const lowerValue = value.toLowerCase();
  if (lowerValue === 'true') return true;
  if (lowerValue === 'false') return false;

  // Numeric values - try to parse as number
  // Only convert if it looks like a number (avoid converting strings like "123abc")
  const trimmed = value.trim();
  if (trimmed !== '' && !isNaN(Number(trimmed))) {
    return Number(trimmed);
  }

  // Everything else stays as string
  return value;
}

/**
 * Get the typed value of a cell (with automatic type coercion)
 * @param {Model} model - IronCalc Model instance
 * @param {number} sheet - Sheet index (0-indexed)
 * @param {number} row - Row number (1-indexed)
 * @param {number} col - Column number (1-indexed, A=1)
 * @returns {any} - Typed cell value (number, boolean, null, or string)
 */
export function getCellValueTyped(model, sheet, row, col) {
  const raw = model.getFormattedCellValue(sheet, row, col);
  return coerceValue(raw);
}
