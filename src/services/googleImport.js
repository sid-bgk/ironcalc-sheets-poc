/**
 * Google Sheets Import Service
 *
 * This module handles importing Google Sheets as xlsx files for use with IronCalc.
 *
 * ## Setup Instructions
 *
 * ### 1. Create Service Account in Google Cloud Console
 * - Go to https://console.cloud.google.com/
 * - Navigate to APIs & Services → Credentials
 * - Create Service Account (e.g., "iron-calc-import")
 * - Download JSON key file as `credentials.json`
 *
 * ### 2. Enable Required APIs
 * - Google Drive API
 * - Google Sheets API
 *
 * ### 3. Share Spreadsheets with Service Account
 * - Find the service account email in credentials.json: "client_email" field
 * - Example: iron-calc@project-id.iam.gserviceaccount.com
 * - Share target spreadsheets with this email address (Viewer access is sufficient)
 *
 * ### 4. Configure Environment Variables
 * - GOOGLE_CREDENTIALS_PATH: Path to credentials.json (default: ./credentials.json)
 * - MODELS_DIR: Directory for downloaded models (default: ./models)
 *
 * ### 5. Shared Drive Support
 * - All API calls use `supportsAllDrives: true` for shared drive compatibility
 * - Service account must be added to the shared drive or have file-level access
 *
 * @module googleImport
 */

import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { loadWorkbook } from '../engine/workbookLoader.js';
import { classifyAllNamedRanges } from '../engine/namedRanges.js';

// Configuration
// GOOGLE_CREDENTIALS_PATH: Path to service account JSON credentials
// MODELS_DIR: Directory where imported xlsx files are stored
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || './credentials.json';
const MODELS_DIR = process.env.MODELS_DIR || './models';

// Google API clients (initialized lazily)
let auth = null;
let drive = null;
let sheets = null;

/**
 * Custom error class for import failures
 */
class ImportError extends Error {
  constructor(message, subcode, statusCode = 500) {
    super(message);
    this.code = 'IMPORT_FAILED';
    this.subcode = subcode;
    this.statusCode = statusCode;
  }
}

/**
 * Initialize Google Auth with service account credentials
 * @returns {Promise<void>}
 * @throws {ImportError} If credentials are missing or invalid
 */
async function initGoogleAuth() {
  if (auth) return; // Already initialized

  console.log(`[googleImport] Initializing Google Auth...`);

  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new ImportError(
      `Service account credentials not found. Please place credentials.json at ${CREDENTIALS_PATH} or set GOOGLE_CREDENTIALS_PATH environment variable.`,
      'AUTH_ERROR',
      500
    );
  }

  try {
    auth = new google.auth.GoogleAuth({
      keyFile: CREDENTIALS_PATH,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/spreadsheets.readonly'
      ]
    });

    drive = google.drive({ version: 'v3', auth });
    sheets = google.sheets({ version: 'v4', auth });
    console.log(`[googleImport] Google Auth initialized successfully`);
  } catch (error) {
    throw new ImportError(
      `Failed to initialize Google Auth: ${error.message}`,
      'AUTH_ERROR',
      500
    );
  }
}

/**
 * Ensure models directory exists
 */
function ensureModelsDir() {
  if (!fs.existsSync(MODELS_DIR)) {
    fs.mkdirSync(MODELS_DIR, { recursive: true });
    console.log(`[googleImport] Created models directory: ${MODELS_DIR}`);
  }
}

/**
 * Get file metadata from Google Drive
 * @param {string} spreadsheetId - Google spreadsheet ID
 * @returns {Promise<{name: string, mimeType: string}>}
 */
async function getFileMetadata(spreadsheetId) {
  try {
    const response = await drive.files.get({
      fileId: spreadsheetId,
      fields: 'id,name,mimeType',
      supportsAllDrives: true
    });
    return response.data;
  } catch (error) {
    if (error.code === 404) {
      throw new ImportError(
        `Spreadsheet not found. Check that the spreadsheet ID is correct and the file exists.`,
        'NOT_FOUND',
        404
      );
    }
    if (error.code === 403) {
      throw new ImportError(
        `Permission denied. Share the spreadsheet with your service account email address.`,
        'PERMISSION_DENIED',
        403
      );
    }
    throw new ImportError(
      `Failed to get file metadata: ${error.message}`,
      'NETWORK_ERROR',
      500
    );
  }
}

/**
 * Download file from Google Drive
 * @param {string} spreadsheetId - Google spreadsheet ID
 * @param {string} mimeType - File MIME type
 * @param {string} outputPath - Local path to save file
 * @returns {Promise<void>}
 */
async function downloadFile(spreadsheetId, mimeType, outputPath) {
  const isNativeGoogleSheet = mimeType === 'application/vnd.google-apps.spreadsheet';
  const isUploadedXlsx = mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  let response;

  try {
    if (isNativeGoogleSheet) {
      // Export native Google Sheet as xlsx
      console.log(`[googleImport] Exporting native Google Sheet as xlsx...`);
      response = await drive.files.export({
        fileId: spreadsheetId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        supportsAllDrives: true
      }, {
        responseType: 'arraybuffer'
      });
    } else if (isUploadedXlsx) {
      // Download uploaded xlsx directly
      console.log(`[googleImport] Downloading uploaded xlsx file...`);
      response = await drive.files.get({
        fileId: spreadsheetId,
        alt: 'media',
        supportsAllDrives: true
      }, {
        responseType: 'arraybuffer'
      });
    } else {
      throw new ImportError(
        `Unsupported file type: ${mimeType}. Only Google Sheets and xlsx files are supported.`,
        'INVALID_FILE_TYPE',
        400
      );
    }

    // Write file to disk
    fs.writeFileSync(outputPath, Buffer.from(response.data));
    console.log(`[googleImport] File saved to: ${outputPath}`);
  } catch (error) {
    if (error instanceof ImportError) throw error;

    if (error.code === 'ECONNRESET' || error.code === 'ETIMEDOUT') {
      throw new ImportError(
        `Network error during download. Please try again.`,
        'NETWORK_ERROR',
        500
      );
    }

    throw new ImportError(
      `Failed to download file: ${error.message}`,
      'NETWORK_ERROR',
      500
    );
  }
}

/**
 * Validate downloaded file
 * @param {string} filePath - Path to downloaded file
 * @throws {ImportError} If file is invalid
 */
function validateFile(filePath) {
  const stats = fs.statSync(filePath);

  if (stats.size < 1000) {
    // Remove corrupted file
    fs.unlinkSync(filePath);
    throw new ImportError(
      `Downloaded file appears to be corrupted (size: ${stats.size} bytes). Please try again.`,
      'INVALID_FILE',
      500
    );
  }

  console.log(`[googleImport] File validated: ${stats.size} bytes`);
}

/**
 * Import a Google Spreadsheet
 * @param {string} spreadsheetId - Google spreadsheet ID
 * @returns {Promise<{modelId: string, status: string, path: string}>}
 */
export async function importSpreadsheet(spreadsheetId) {
  console.log(`[googleImport] Starting import for spreadsheet...`);

  // Initialize auth
  await initGoogleAuth();

  // Ensure models directory exists
  ensureModelsDir();

  // Get file metadata
  const metadata = await getFileMetadata(spreadsheetId);
  console.log(`[googleImport] File: ${metadata.name}, Type: ${metadata.mimeType}`);

  // Generate unique model ID
  const modelId = uuidv4();
  const outputPath = path.join(MODELS_DIR, `${modelId}.xlsx`);

  // Download file
  await downloadFile(spreadsheetId, metadata.mimeType, outputPath);

  // Validate file
  validateFile(outputPath);

  // Validate with IronCalc and extract metadata
  // This ensures the file can actually be used for calculations
  console.log(`[googleImport] Validating with IronCalc and extracting metadata...`);

  // Validate file loads in IronCalc and extract named ranges
  // This throws ImportError if IronCalc can't process the file (e.g., unsupported formulas)
  const namedRanges = validateAndExtractNamedRanges(outputPath);

  // Get sheet tab names from Sheets API (works for native Google Sheets)
  const sheetMeta = await getSheetMetadata(spreadsheetId);

  console.log(`[googleImport] Import complete: modelId=${modelId}`);

  return {
    modelId,
    status: 'ready',
    path: outputPath,
    metadata: {
      sheetName: sheetMeta.sheetName || metadata.name,
      sheets: sheetMeta.sheets,
      inputs: namedRanges.inputs,
      outputs: namedRanges.outputs
    }
  };
}

/**
 * Get sheet metadata (tab names) from Google Sheets API
 * @param {string} spreadsheetId - Google spreadsheet ID
 * @returns {Promise<{sheetName: string, sheets: string[]}>}
 */
async function getSheetMetadata(spreadsheetId) {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'properties.title,sheets.properties.title'
    });

    const sheetName = response.data.properties?.title || 'Unknown';
    const tabNames = response.data.sheets?.map(s => s.properties.title) || [];

    console.log(`[googleImport] Sheet metadata: ${sheetName}, tabs: ${tabNames.join(', ')}`);
    return { sheetName, sheets: tabNames };
  } catch (error) {
    console.log(`[googleImport] Could not fetch sheet metadata: ${error.message}`);
    // Return defaults if metadata fetch fails (file might be xlsx, not native sheet)
    return { sheetName: 'Unknown', sheets: [] };
  }
}

/**
 * Validate that IronCalc can load the downloaded xlsx file and extract named ranges
 * This catches incompatible formulas or corrupted files early at import time
 * @param {string} filePath - Path to the xlsx file
 * @returns {{inputs: Array, outputs: Array}}
 * @throws {ImportError} If IronCalc cannot load or process the file
 */
function validateAndExtractNamedRanges(filePath) {
  console.log(`[googleImport] Validating file with IronCalc: ${filePath}`);

  let model;
  try {
    model = loadWorkbook(filePath);
  } catch (error) {
    // IronCalc failed to load the file - this is a critical error
    // Could be unsupported formulas, corrupted file, or incompatible Excel features
    console.error(`[googleImport] IronCalc failed to load file: ${error.message}`);

    // Clean up the downloaded file since it's unusable
    try {
      fs.unlinkSync(filePath);
      console.log(`[googleImport] Removed unusable file: ${filePath}`);
    } catch (unlinkError) {
      // Ignore cleanup errors
    }

    throw new ImportError(
      `IronCalc cannot process this spreadsheet: ${error.message}. The file may contain unsupported formulas or features. Please simplify the spreadsheet and try again.`,
      'IRONCALC_LOAD_ERROR',
      400
    );
  }

  // File loaded successfully, now extract named ranges
  try {
    const classified = classifyAllNamedRanges(model);

    // Format for API response
    const inputs = classified.inputs.map(r => ({
      name: r.name,
      cellReference: r.reference.replace(/^=/, '')
    }));

    const outputs = classified.outputs.map(r => ({
      name: r.name,
      cellReference: r.reference.replace(/^=/, ''),
      type: r.rangeType.toLowerCase()
    }));

    console.log(`[googleImport] IronCalc validation passed. Found ${inputs.length} inputs, ${outputs.length} outputs`);
    return { inputs, outputs };
  } catch (error) {
    console.error(`[googleImport] Failed to extract named ranges: ${error.message}`);
    throw new ImportError(
      `Failed to analyze spreadsheet structure: ${error.message}`,
      'IRONCALC_ANALYSIS_ERROR',
      500
    );
  }
}

/**
 * Check if Google Auth is configured and can be initialized
 * @returns {Promise<{configured: boolean, message: string}>}
 */
export async function checkAuthStatus() {
  try {
    await initGoogleAuth();
    return {
      configured: true,
      message: 'Google API credentials configured and initialized successfully'
    };
  } catch (error) {
    return {
      configured: false,
      message: error.message
    };
  }
}

// Export for testing
export { initGoogleAuth, ImportError };
