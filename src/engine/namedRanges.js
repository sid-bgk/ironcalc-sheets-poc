import { isFormulaCell } from './cellReader.js';

/**
 * Get all named ranges from a workbook
 * @param {Model} model - IronCalc Model instance
 * @returns {Array} - Array of named range objects { name, formula, scope }
 */
export function getNamedRanges(model) {
  return model.getDefinedNameList();
}

/**
 * Create a new named range
 * @param {Model} model - IronCalc Model instance
 * @param {string} name - Name for the range
 * @param {string} formula - Cell reference (e.g., "Input!$A$1")
 * @param {number|null} scope - Sheet scope (null for workbook scope)
 */
export function createNamedRange(model, name, formula, scope = null) {
  model.newDefinedName(name, scope, formula);
}

/**
 * Parse a cell reference string to extract sheet and address
 * @param {string} reference - Cell reference (e.g., "Sheet1!$A$1")
 * @returns {object|null} - { sheet, startCol, startRow, endCol, endRow, isRange }
 */
export function parseCellReference(reference) {
  if (!reference) return null;

  // Remove leading "=" if present
  const ref = reference.startsWith('=') ? reference.slice(1) : reference;

  // Handle "Sheet1!$A$1" or "'Sheet Name'!$A$1:$B$10" format
  const match = ref.match(/^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);

  if (!match) return null;

  const [, sheet, startCol, startRow, endCol, endRow] = match;

  return {
    sheet,
    startCol,
    startRow: parseInt(startRow),
    endCol: endCol || startCol,
    endRow: endRow ? parseInt(endRow) : parseInt(startRow),
    isRange: !!endCol
  };
}

/**
 * Get sheet index by name
 * @param {Model} model - IronCalc Model instance
 * @param {string} sheetName - Name of the sheet
 * @returns {number} - Sheet index (0-indexed) or -1 if not found
 */
export function getSheetIndex(model, sheetName) {
  const sheets = model.getWorksheetsProperties();
  const index = sheets.findIndex(s => s.name === sheetName);
  return index;
}

/**
 * Convert column letter to number (A=1, B=2, etc.)
 * @param {string} col - Column letter(s)
 * @returns {number} - Column number
 */
export function colToNum(col) {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Determine the type of a named range based on its dimensions
 * @param {object} parsed - Parsed cell reference from parseCellReference()
 * @returns {string} - 'SINGLE'|'HORIZONTAL'|'VERTICAL'|'TABLE'
 */
export function getRangeType(parsed) {
  if (!parsed) return 'UNKNOWN';

  const rows = parsed.endRow - parsed.startRow + 1;
  const cols = colToNum(parsed.endCol) - colToNum(parsed.startCol) + 1;

  if (rows === 1 && cols === 1) return 'SINGLE';
  if (rows === 1 && cols > 1) return 'HORIZONTAL';
  if (rows > 1 && cols === 1) return 'VERTICAL';
  if (rows > 1 && cols > 1) return 'TABLE';
  return 'UNKNOWN';
}

/**
 * Classify a named range as INPUT or OUTPUT based on target cell content
 * @param {Model} model - IronCalc Model instance
 * @param {object} namedRange - Named range object { name, formula }
 * @returns {object} - { name, type, rangeType, reference, sheet, sheetIndex, row, col, endRow, endCol, rows, cols, isRange }
 */
export function classifyNamedRange(model, namedRange) {
  const parsed = parseCellReference(namedRange.formula);

  if (!parsed) {
    return { name: namedRange.name, type: 'UNKNOWN', rangeType: 'UNKNOWN', reference: namedRange.formula };
  }

  const sheetIndex = getSheetIndex(model, parsed.sheet);
  if (sheetIndex === -1) {
    return { name: namedRange.name, type: 'UNKNOWN', rangeType: 'UNKNOWN', reference: namedRange.formula };
  }

  const startCol = colToNum(parsed.startCol);
  const endCol = colToNum(parsed.endCol);
  const rangeType = getRangeType(parsed);

  // Calculate dimensions
  const rows = parsed.endRow - parsed.startRow + 1;
  const cols = endCol - startCol + 1;

  // TABLE ranges are always OUTPUT (per strategy doc)
  // For SINGLE, HORIZONTAL, VERTICAL: check first cell for formula
  let type;
  if (rangeType === 'TABLE') {
    type = 'OUTPUT';
  } else {
    const hasFormula = isFormulaCell(model, sheetIndex, parsed.startRow, startCol);
    type = hasFormula ? 'OUTPUT' : 'INPUT';
  }

  return {
    name: namedRange.name,
    type,
    rangeType,
    reference: namedRange.formula,
    sheet: parsed.sheet,
    sheetIndex,
    row: parsed.startRow,
    col: startCol,
    endRow: parsed.endRow,
    endCol,
    rows,
    cols,
    isRange: parsed.isRange
  };
}

/**
 * Classify all named ranges in a workbook
 * @param {Model} model - IronCalc Model instance
 * @returns {object} - { inputs: [], outputs: [], unknown: [] }
 */
export function classifyAllNamedRanges(model) {
  const namedRanges = getNamedRanges(model);

  const result = {
    inputs: [],
    outputs: [],
    unknown: []
  };

  for (const nr of namedRanges) {
    const classified = classifyNamedRange(model, nr);

    if (classified.type === 'INPUT') {
      result.inputs.push(classified);
    } else if (classified.type === 'OUTPUT') {
      result.outputs.push(classified);
    } else {
      result.unknown.push(classified);
    }
  }

  return result;
}
