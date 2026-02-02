import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Usage: node scripts/test-native-export.js <spreadsheetId>
const spreadsheetId = process.argv[2];

if (!spreadsheetId) {
  console.error('Usage: node scripts/test-native-export.js <spreadsheetId>');
  process.exit(1);
}

// Timeout helper
const timeout = (ms, promise, label) => {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms)
    )
  ]);
};

async function testNativeExport() {
  console.log(`\n📊 Testing Google Sheet/xlsx download: ${spreadsheetId}\n`);

  // Step 0: Auth
  console.log('0️⃣ Authenticating...');
  const credPath = path.join(__dirname, '../code/credentials.json');

  if (!fs.existsSync(credPath)) {
    console.error(`   ❌ credentials.json not found at: ${credPath}`);
    process.exit(1);
  }
  console.log(`   ✓ Found credentials at: ${credPath}`);

  const auth = new google.auth.GoogleAuth({
    keyFile: credPath,
    scopes: [
      'https://www.googleapis.com/auth/drive.readonly',
      'https://www.googleapis.com/auth/spreadsheets.readonly'
    ]
  });

  const drive = google.drive({ version: 'v3', auth });
  const sheets = google.sheets({ version: 'v4', auth });
  console.log('   ✓ Auth initialized');

  // Step 1: Get file metadata from Drive to determine file type
  console.log('\n1️⃣ Fetching file metadata from Drive...');
  const fileMetadata = await timeout(
    15000,
    drive.files.get({ fileId: spreadsheetId, fields: 'id,name,mimeType', supportsAllDrives: true }),
    'File metadata fetch'
  );

  const mimeType = fileMetadata.data.mimeType;
  const fileName = fileMetadata.data.name;
  console.log(`   ✓ File name: ${fileName}`);
  console.log(`   ✓ MIME type: ${mimeType}`);

  const isNativeGoogleSheet = mimeType === 'application/vnd.google-apps.spreadsheet';
  const isUploadedXlsx = mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  if (isNativeGoogleSheet) {
    console.log('   📋 Type: Native Google Sheet (or converted xlsx)');
  } else if (isUploadedXlsx) {
    console.log('   📋 Type: Uploaded xlsx (NOT converted to Google Sheets)');
  } else {
    console.log(`   ⚠ Unknown type: ${mimeType}`);
  }

  // Step 2: Get sheet metadata (only works for Google Sheets)
  if (isNativeGoogleSheet) {
    console.log('\n2️⃣ Fetching sheet metadata via Sheets API...');
    try {
      const metadata = await timeout(
        15000,
        sheets.spreadsheets.get({ spreadsheetId }),
        'Sheet metadata fetch'
      );
      console.log(`   ✓ Sheets: ${metadata.data.sheets.map(s => s.properties.title).join(', ')}`);

      // Check for named ranges
      const namedRanges = metadata.data.namedRanges || [];
      if (namedRanges.length > 0) {
        console.log(`   ✓ Named ranges: ${namedRanges.map(nr => nr.name).join(', ')}`);
      } else {
        console.log('   ⚠ No named ranges found');
      }

      // Check for conditional formatting
      let cfCount = 0;
      metadata.data.sheets.forEach(sheet => {
        const rules = sheet.conditionalFormats || [];
        cfCount += rules.length;
      });
      console.log(`   ✓ Conditional formatting rules: ${cfCount}`);
    } catch (err) {
      console.log(`   ⚠ Could not fetch sheet metadata: ${err.message}`);
    }
  } else {
    console.log('\n2️⃣ Skipping Sheets API (not a Google Sheet format)');
  }

  // Step 3: Download the file
  console.log('\n3️⃣ Downloading file...');

  // Ensure result directory exists
  const resultDir = path.join(__dirname, '../result');
  if (!fs.existsSync(resultDir)) fs.mkdirSync(resultDir, { recursive: true });
  const outputPath = path.join(resultDir, `download-test-${Date.now()}.xlsx`);

  let exportRes;

  if (isNativeGoogleSheet) {
    // Use export for Google Sheets
    console.log('   Using drive.files.export() for Google Sheet...');
    exportRes = await timeout(
      60000,
      drive.files.export({
        fileId: spreadsheetId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        supportsAllDrives: true
      }, {
        responseType: 'arraybuffer'
      }),
      'Export'
    );
  } else {
    // Use direct download for uploaded xlsx
    console.log('   Using drive.files.get() for uploaded xlsx...');
    exportRes = await timeout(
      60000,
      drive.files.get({
        fileId: spreadsheetId,
        alt: 'media',
        supportsAllDrives: true
      }, {
        responseType: 'arraybuffer'
      }),
      'Download'
    );
  }

  fs.writeFileSync(outputPath, Buffer.from(exportRes.data));

  const stats = fs.statSync(outputPath);
  console.log(`   ✓ Downloaded: ${outputPath}`);
  console.log(`   ✓ File size: ${stats.size} bytes`);

  // Step 4: Validation
  console.log('\n4️⃣ Validation:');
  if (stats.size < 1000) {
    console.log('   ⚠ WARNING: File seems too small - may be corrupted');
  } else {
    console.log('   ✓ File size looks reasonable');
  }

  console.log('\n✅ DONE!');
  console.log(`   File: ${outputPath}`);

  if (!isNativeGoogleSheet) {
    console.log('\n⚠ NOTE: This was an uploaded xlsx (not converted).');
    console.log('   - Sheets API metadata (named ranges, conditional formatting) not available via API');
    console.log('   - The xlsx file itself contains this data - open it to verify');
    console.log('   - Consider: Have ops team "Open with Google Sheets" to convert for full API access');
  }
}

testNativeExport().catch(err => {
  console.error('\n❌ Error:', err.message);
  if (err.message.includes('timed out')) {
    console.error('   Possible causes: network issue, large file, or auth problem');
  }
  if (err.code === 404) {
    console.error('   File not found - check the file ID');
  }
  if (err.code === 403) {
    console.error('   Permission denied - share the file with your service account email');
  }
  process.exit(1);
});
