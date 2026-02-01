import { classifyAllNamedRanges } from './namedRanges.js';
import { getCellValue, getCellValueTyped } from './cellReader.js';
import { v4 as uuidv4 } from 'uuid';
import fs from 'fs';
import path from 'path';

// Cache for classified ranges (optimization)
let cachedClassification = null;

/**
 * Get classified named ranges (with caching)
 * @param {Model} model - IronCalc Model instance
 * @param {boolean} refresh - Force refresh cache
 * @returns {object} - { inputs: [], outputs: [] }
 */
function getClassifiedRanges(model, refresh = false) {
  if (!cachedClassification || refresh) {
    cachedClassification = classifyAllNamedRanges(model);
  }
  return cachedClassification;
}

/**
 * Set array values for a HORIZONTAL or VERTICAL input range
 * Flexible sizing: shorter arrays pad with empty, longer arrays truncate
 * @param {Model} model - IronCalc Model instance
 * @param {object} input - Classified input object from classifyNamedRange
 * @param {Array} values - Array of values to set
 */
function setInputArrayValue(model, input, values) {
  const { sheetIndex, row, col, endRow, endCol, rangeType, rows, cols } = input;

  // HORIZONTAL: map values left-to-right across columns
  if (rangeType === 'HORIZONTAL') {
    for (let i = 0; i < cols; i++) {
      const value = i < values.length ? values[i] : '';
      model.setUserInput(sheetIndex, row, col + i, String(value));
    }
    return;
  }

  // VERTICAL: map values top-to-bottom across rows
  if (rangeType === 'VERTICAL') {
    for (let i = 0; i < rows; i++) {
      const value = i < values.length ? values[i] : '';
      model.setUserInput(sheetIndex, row + i, col, String(value));
    }
    return;
  }

  // SINGLE with array - error
  throw new Error(`Cannot set array value for SINGLE input range: ${input.name}`);
}

/**
 * Set a single input value by named range name (supports scalars and arrays)
 * @param {Model} model - IronCalc Model instance
 * @param {string} inputName - Name of the INPUT named range
 * @param {any} value - Value to set (scalar or array for HORIZONTAL/VERTICAL inputs)
 * @throws {Error} - If input named range not found or array/range mismatch
 */
export function setInputValue(model, inputName, value) {
  const { inputs } = getClassifiedRanges(model);
  const input = inputs.find(i => i.name === inputName);

  if (!input) {
    throw new Error(`Input named range not found: ${inputName}`);
  }

  // If value is array, delegate to array handler
  if (Array.isArray(value)) {
    setInputArrayValue(model, input, value);
    return;
  }

  // Scalar value - set single cell (existing behavior)
  model.setUserInput(input.sheetIndex, input.row, input.col, String(value));
}

/**
 * Set multiple input values from an object
 * @param {Model} model - IronCalc Model instance
 * @param {object} inputsObj - Object of { inputName: value }
 * @throws {Error} - If any input named range not found
 */
export function setInputValues(model, inputsObj) {
  for (const [name, value] of Object.entries(inputsObj)) {
    setInputValue(model, name, value);
  }
}

/**
 * Clear the classification cache (call after modifying named ranges)
 */
export function clearCache() {
  cachedClassification = null;
}

/**
 * Get list of all required INPUT named range names
 * @param {Model} model - IronCalc Model instance
 * @returns {string[]} - Array of input names
 */
export function getRequiredInputNames(model) {
  const { inputs } = getClassifiedRanges(model);
  return inputs.map(i => i.name);
}

/**
 * Trigger recalculation of all formulas
 * @param {Model} model - IronCalc Model instance
 */
export function calculate(model) {
  model.evaluate();
}

/**
 * Get output value for a classified output range (handles all range types)
 * @param {Model} model - IronCalc Model instance
 * @param {object} output - Classified output object from classifyNamedRange
 * @returns {any} - Single value, 1D array, or TABLE object depending on rangeType
 */
function getOutputRangeValue(model, output) {
  const { sheetIndex, row, col, endRow, endCol, rangeType, rows, cols } = output;

  // SINGLE: return scalar value (existing behavior)
  if (rangeType === 'SINGLE') {
    return getCellValueTyped(model, sheetIndex, row, col);
  }

  // HORIZONTAL: return 1D array (iterate columns)
  if (rangeType === 'HORIZONTAL') {
    const values = [];
    for (let c = col; c <= endCol; c++) {
      values.push(getCellValueTyped(model, sheetIndex, row, c));
    }
    return values;
  }

  // VERTICAL: return 1D array (iterate rows)
  if (rangeType === 'VERTICAL') {
    const values = [];
    for (let r = row; r <= endRow; r++) {
      values.push(getCellValueTyped(model, sheetIndex, r, col));
    }
    return values;
  }

  // TABLE: return structured object with headers and data
  if (rangeType === 'TABLE') {
    // First row = headers
    const headers = [];
    for (let c = col; c <= endCol; c++) {
      headers.push(getCellValueTyped(model, sheetIndex, row, c));
    }

    // Remaining rows = data
    const data = [];
    for (let r = row + 1; r <= endRow; r++) {
      const rowData = [];
      for (let c = col; c <= endCol; c++) {
        rowData.push(getCellValueTyped(model, sheetIndex, r, c));
      }
      data.push(rowData);
    }

    return {
      type: 'TABLE',
      rows: rows - 1, // data rows only (excluding header)
      cols,
      headers,
      data
    };
  }

  // Fallback for unknown types - return single cell value
  return getCellValueTyped(model, sheetIndex, row, col);
}

/**
 * Get a single output value by named range name
 * @param {Model} model - IronCalc Model instance
 * @param {string} outputName - Name of the OUTPUT named range
 * @returns {any} - Single value, 1D array, or TABLE object depending on range type
 * @throws {Error} - If output named range not found
 */
export function getOutputValue(model, outputName) {
  const { outputs } = getClassifiedRanges(model);
  const output = outputs.find(o => o.name === outputName);

  if (!output) {
    throw new Error(`Output named range not found: ${outputName}`);
  }

  return getOutputRangeValue(model, output);
}

/**
 * Get all output values as an object
 * @param {Model} model - IronCalc Model instance
 * @returns {object} - Object of { outputName: value }
 */
export function getAllOutputValues(model) {
  const { outputs } = getClassifiedRanges(model);
  const result = {};

  for (const output of outputs) {
    result[output.name] = getOutputRangeValue(model, output);
  }

  return result;
}

// Results folder path
const RESULTS_DIR = './results';

/**
 * Ensure results directory exists
 */
function ensureResultsDir() {
  if (!fs.existsSync(RESULTS_DIR)) {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
    console.log(`[calculator] Created results directory: ${RESULTS_DIR}`);
  }
}

/**
 * Save the model to results folder with unique filename
 * @param {Model} model - IronCalc Model instance
 * @returns {string} - Path to saved file
 */
function saveResult(model) {
  ensureResultsDir();
  const filename = `result_${uuidv4()}.xlsx`;
  const filepath = path.join(RESULTS_DIR, filename);
  model.saveToXlsx(filepath);
  console.log(`[calculator] Saved result to: ${filepath}`);
  return filepath;
}

/**
 * Execute full calculation workflow: set inputs → calculate → save → return outputs
 * @param {Model} model - IronCalc Model instance
 * @param {object} inputs - Object of { inputName: value }
 * @returns {object} - Object of { outputs, resultFile }
 */
export function executeCalculation(model, inputs) {
  console.log(`[calculator] Executing calculation with inputs:`, inputs);

  // 1. Set all input values
  console.log(`[calculator] Setting input values...`);
  setInputValues(model, inputs);

  // 2. Trigger recalculation
  console.log(`[calculator] Triggering recalculation...`);
  calculate(model);

  // 3. Save result to file
  const resultFile = saveResult(model);

  // 4. Return all output values
  console.log(`[calculator] Retrieving output values...`);
  const outputs = getAllOutputValues(model);
  console.log(`[calculator] Outputs generated:`, outputs);

  return { outputs, resultFile };
}
