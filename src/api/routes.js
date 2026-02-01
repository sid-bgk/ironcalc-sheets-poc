import { Router } from 'express';
import { loadWorkbook } from '../engine/workbookLoader.js';
import { executeCalculation, getRequiredInputNames } from '../engine/calculator.js';

const router = Router();
let model = null;

/**
 * Initialize workbook (call on server start)
 * @param {string} excelPath - Path to Excel file
 */
export function initWorkbook(excelPath) {
  model = loadWorkbook(excelPath);
  console.log('Workbook loaded:', excelPath);
}

/**
 * POST /api/v1/calculate/dscr
 *
 * Request body: { "inputs": { "InputName": value, ... } }
 * Response: { "outputs": { "OutputName": value, ... } }
 */
router.post('/api/v1/calculate/dscr', (req, res) => {
  console.log(`[routes] API called: POST /api/v1/calculate/dscr`);
  console.log(`[routes] Request body:`, JSON.stringify(req.body));

  try {
    const { inputs } = req.body;

    if (!inputs || typeof inputs !== 'object') {
      console.log(`[routes] Invalid request - missing inputs`);
      return res.status(400).json({
        error: {
          code: 'INVALID_REQUEST',
          message: 'Request body must contain "inputs" object'
        }
      });
    }

    // Check for missing inputs (warning only)
    const requiredInputs = getRequiredInputNames(model);
    const providedInputs = Object.keys(inputs);
    const missingInputs = requiredInputs.filter(r => !providedInputs.includes(r));

    if (missingInputs.length > 0) {
      console.log(`[routes] Warning: Missing inputs: ${missingInputs.join(', ')}`);
    }

    const { outputs, resultFile } = executeCalculation(model, inputs);

    // Build response with optional warnings
    const response = { outputs, resultFile };
    if (missingInputs.length > 0) {
      response.warnings = missingInputs.map(name => `Input '${name}' was not provided`);
    }

    console.log(`[routes] Sending response with outputs, saved to: ${resultFile}`);
    res.json(response);
  } catch (error) {
    console.error(`[routes] Error:`, error.message);
    res.status(500).json({
      error: {
        code: 'CALCULATION_ERROR',
        message: error.message
      }
    });
  }
});

export { router };
