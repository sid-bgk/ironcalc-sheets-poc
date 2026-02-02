import { Router } from 'express';
import { loadWorkbook } from '../engine/workbookLoader.js';
import { executeCalculation, getRequiredInputNames } from '../engine/calculator.js';
import { importSpreadsheet, checkAuthStatus } from '../services/googleImport.js';

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

/**
 * POST /api/v1/import
 *
 * Request body: { "spreadsheetId": "<google-spreadsheet-id>" }
 * Response: { "modelId": "...", "status": "ready", "path": "..." }
 */
router.post('/api/v1/import', async (req, res) => {
  console.log(`[routes] API called: POST /api/v1/import`);

  try {
    const { spreadsheetId } = req.body;

    // Validate spreadsheetId is provided
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.log(`[routes] Invalid request - missing or invalid spreadsheetId`);
      return res.status(400).json({
        error: {
          code: 'INVALID_REQUEST',
          message: 'Request body must contain "spreadsheetId" string'
        }
      });
    }

    // Log truncated ID for debugging (security - don't log full ID)
    const truncatedId = spreadsheetId.length > 10
      ? `${spreadsheetId.substring(0, 10)}...`
      : spreadsheetId;
    console.log(`[routes] Importing spreadsheet: ${truncatedId}`);

    // Import the spreadsheet
    const result = await importSpreadsheet(spreadsheetId);

    console.log(`[routes] Import successful: modelId=${result.modelId}`);
    res.json(result);
  } catch (error) {
    console.error(`[routes] Import error:`, error.message);

    // Handle structured import errors
    if (error.code && error.code.startsWith('IMPORT_FAILED')) {
      return res.status(error.statusCode || 500).json({
        error: {
          code: error.code,
          subcode: error.subcode,
          message: error.message
        }
      });
    }

    // Generic error
    res.status(500).json({
      error: {
        code: 'IMPORT_FAILED',
        message: error.message
      }
    });
  }
});

/**
 * GET /api/v1/import/health
 *
 * Check if Google API credentials are configured
 * Response: { "status": "ok"|"error", "auth": { "configured": boolean, "message": string } }
 */
router.get('/api/v1/import/health', async (req, res) => {
  console.log(`[routes] API called: GET /api/v1/import/health`);

  try {
    const authStatus = await checkAuthStatus();

    res.json({
      status: authStatus.configured ? 'ok' : 'error',
      auth: authStatus
    });
  } catch (error) {
    console.error(`[routes] Health check error:`, error.message);
    res.status(500).json({
      status: 'error',
      auth: {
        configured: false,
        message: error.message
      }
    });
  }
});

export { router };
