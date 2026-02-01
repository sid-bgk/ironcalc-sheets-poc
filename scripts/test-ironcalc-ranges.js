/**
 * Task 2: Analyze IronCalc Range Handling
 * Tests how IronCalc reports and handles named ranges that span multiple cells
 */

import { Model } from '@ironcalc/nodejs';
import fs from 'fs';

const XLSX_PATH = './DSCR_NoArrayFormulas_Testing.xlsx';

// ============================================================
// Helper functions (copied from namedRanges.js for standalone testing)
// ============================================================

function colToNum(col) {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result;
}

function numToCol(num) {
  let result = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

function parseCellReference(reference) {
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

function getRangeType(parsed) {
  if (!parsed) return 'UNKNOWN';

  const rows = parsed.endRow - parsed.startRow + 1;
  const cols = colToNum(parsed.endCol) - colToNum(parsed.startCol) + 1;

  if (rows === 1 && cols === 1) return 'SINGLE';
  if (rows === 1 && cols > 1) return 'HORIZONTAL';
  if (rows > 1 && cols === 1) return 'VERTICAL';
  if (rows > 1 && cols > 1) return 'TABLE';
  return 'UNKNOWN';
}

function getSheetIndex(model, sheetName) {
  const sheets = model.getWorksheetsProperties();
  return sheets.findIndex(s => s.name === sheetName);
}

// ============================================================
// Main Test Functions
// ============================================================

function printSeparator(title) {
  console.log('\n' + '='.repeat(80));
  console.log(title);
  console.log('='.repeat(80));
}

function testNamedRangeRetrieval(model) {
  printSeparator('TEST 1: Named Range Retrieval via getDefinedNameList()');

  const namedRanges = model.getDefinedNameList();
  console.log(`\nFound ${namedRanges.length} named ranges:\n`);

  console.log(`${'Name'.padEnd(30)} ${'Formula'.padEnd(35)} ${'Scope'}`);
  console.log('-'.repeat(80));

  for (const nr of namedRanges) {
    console.log(`${nr.name.padEnd(30)} ${nr.formula.padEnd(35)} ${nr.scope ?? 'workbook'}`);
  }

  return namedRanges;
}

function testParseCellReference(namedRanges) {
  printSeparator('TEST 2: parseCellReference() with All Range Types');

  console.log(`\n${'Name'.padEnd(25)} ${'Type'.padEnd(12)} ${'Dims'.padEnd(10)} ${'Parsed Details'}`);
  console.log('-'.repeat(80));

  for (const nr of namedRanges) {
    const parsed = parseCellReference(nr.formula);
    const rangeType = getRangeType(parsed);

    if (parsed) {
      const rows = parsed.endRow - parsed.startRow + 1;
      const cols = colToNum(parsed.endCol) - colToNum(parsed.startCol) + 1;
      const dims = `${rows}x${cols}`;

      console.log(
        `${nr.name.padEnd(25)} ${rangeType.padEnd(12)} ${dims.padEnd(10)} ` +
        `${parsed.sheet}!${parsed.startCol}${parsed.startRow}:${parsed.endCol}${parsed.endRow}`
      );
    } else {
      console.log(`${nr.name.padEnd(25)} PARSE_FAIL  N/A        Could not parse: ${nr.formula}`);
    }
  }
}

function testReadingSingleCell(model, namedRanges) {
  printSeparator('TEST 3: Reading SINGLE Cell Values');

  const singles = namedRanges.filter(nr => {
    const parsed = parseCellReference(nr.formula);
    return getRangeType(parsed) === 'SINGLE';
  });

  console.log(`\nFound ${singles.length} SINGLE named ranges:\n`);

  for (const nr of singles) {
    const parsed = parseCellReference(nr.formula);
    const sheetIndex = getSheetIndex(model, parsed.sheet);
    const col = colToNum(parsed.startCol);

    const value = model.getFormattedCellValue(sheetIndex, parsed.startRow, col);
    const content = model.getCellContent(sheetIndex, parsed.startRow, col);
    const isFormula = content?.startsWith('=');

    console.log(`${nr.name}:`);
    console.log(`  Location: ${parsed.sheet}!${parsed.startCol}${parsed.startRow} (sheet=${sheetIndex}, row=${parsed.startRow}, col=${col})`);
    console.log(`  Value: ${value}`);
    console.log(`  Formula: ${isFormula ? content : '(not a formula)'}`);
    console.log(`  Classification: ${isFormula ? 'OUTPUT' : 'INPUT'}`);
    console.log();
  }
}

function testReadingHorizontalArray(model, namedRanges) {
  printSeparator('TEST 4: Reading HORIZONTAL Array Values');

  const horizontals = namedRanges.filter(nr => {
    const parsed = parseCellReference(nr.formula);
    return getRangeType(parsed) === 'HORIZONTAL';
  });

  console.log(`\nFound ${horizontals.length} HORIZONTAL named ranges:\n`);

  for (const nr of horizontals) {
    const parsed = parseCellReference(nr.formula);
    const sheetIndex = getSheetIndex(model, parsed.sheet);
    const startCol = colToNum(parsed.startCol);
    const endCol = colToNum(parsed.endCol);
    const row = parsed.startRow;

    console.log(`${nr.name}:`);
    console.log(`  Range: ${parsed.sheet}!${parsed.startCol}${row}:${parsed.endCol}${row}`);
    console.log(`  Dimensions: 1 row x ${endCol - startCol + 1} cols`);

    const values = [];
    const formulas = [];

    for (let c = startCol; c <= endCol; c++) {
      values.push(model.getFormattedCellValue(sheetIndex, row, c));
      const content = model.getCellContent(sheetIndex, row, c);
      formulas.push(content?.startsWith('=') ? 'F' : 'V');
    }

    console.log(`  Values: [${values.join(', ')}]`);
    console.log(`  Cell types: [${formulas.join(', ')}] (F=formula, V=value)`);

    const firstCellHasFormula = formulas[0] === 'F';
    console.log(`  Classification: ${firstCellHasFormula ? 'OUTPUT' : 'INPUT'} (based on first cell)`);
    console.log();
  }
}

function testReadingVerticalArray(model, namedRanges) {
  printSeparator('TEST 5: Reading VERTICAL Array Values');

  const verticals = namedRanges.filter(nr => {
    const parsed = parseCellReference(nr.formula);
    return getRangeType(parsed) === 'VERTICAL';
  });

  console.log(`\nFound ${verticals.length} VERTICAL named ranges:\n`);

  for (const nr of verticals) {
    const parsed = parseCellReference(nr.formula);
    const sheetIndex = getSheetIndex(model, parsed.sheet);
    const col = colToNum(parsed.startCol);
    const startRow = parsed.startRow;
    const endRow = parsed.endRow;

    console.log(`${nr.name}:`);
    console.log(`  Range: ${parsed.sheet}!${parsed.startCol}${startRow}:${parsed.endCol}${endRow}`);
    console.log(`  Dimensions: ${endRow - startRow + 1} rows x 1 col`);

    const values = [];
    const formulas = [];

    for (let r = startRow; r <= endRow; r++) {
      values.push(model.getFormattedCellValue(sheetIndex, r, col));
      const content = model.getCellContent(sheetIndex, r, col);
      formulas.push(content?.startsWith('=') ? 'F' : 'V');
    }

    console.log(`  Values: [${values.join(', ')}]`);
    console.log(`  Cell types: [${formulas.join(', ')}] (F=formula, V=value)`);

    const firstCellHasFormula = formulas[0] === 'F';
    console.log(`  Classification: ${firstCellHasFormula ? 'OUTPUT' : 'INPUT'} (based on first cell)`);
    console.log();
  }
}

function testReadingTable(model, namedRanges) {
  printSeparator('TEST 6: Reading TABLE Values');

  const tables = namedRanges.filter(nr => {
    const parsed = parseCellReference(nr.formula);
    return getRangeType(parsed) === 'TABLE';
  });

  console.log(`\nFound ${tables.length} TABLE named ranges:\n`);

  for (const nr of tables) {
    const parsed = parseCellReference(nr.formula);
    const sheetIndex = getSheetIndex(model, parsed.sheet);
    const startCol = colToNum(parsed.startCol);
    const endCol = colToNum(parsed.endCol);
    const startRow = parsed.startRow;
    const endRow = parsed.endRow;

    const rows = endRow - startRow + 1;
    const cols = endCol - startCol + 1;

    console.log(`${nr.name}:`);
    console.log(`  Range: ${parsed.sheet}!${parsed.startCol}${startRow}:${parsed.endCol}${endRow}`);
    console.log(`  Dimensions: ${rows} rows x ${cols} cols = ${rows * cols} cells`);

    // Read header row (first row of table)
    const headers = [];
    for (let c = startCol; c <= endCol; c++) {
      headers.push(model.getFormattedCellValue(sheetIndex, startRow, c));
    }
    console.log(`  Headers (row ${startRow}): [${headers.slice(0, 5).join(', ')}${cols > 5 ? `, ... (${cols} total)` : ''}]`);

    // Count formulas in header vs data
    let headerFormulas = 0;
    let dataFormulas = 0;
    let totalDataCells = (rows - 1) * cols;

    for (let c = startCol; c <= endCol; c++) {
      const headerContent = model.getCellContent(sheetIndex, startRow, c);
      if (headerContent?.startsWith('=')) headerFormulas++;
    }

    for (let r = startRow + 1; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const content = model.getCellContent(sheetIndex, r, c);
        if (content?.startsWith('=')) dataFormulas++;
      }
    }

    console.log(`  Header row formulas: ${headerFormulas}/${cols}`);
    console.log(`  Data rows formulas: ${dataFormulas}/${totalDataCells} (${((dataFormulas/totalDataCells)*100).toFixed(1)}%)`);
    console.log(`  Classification: OUTPUT (TABLE type is always OUTPUT)`);

    // Read first data row as sample
    if (rows > 1) {
      const firstDataRow = [];
      for (let c = startCol; c <= endCol && c < startCol + 5; c++) {
        firstDataRow.push(model.getFormattedCellValue(sheetIndex, startRow + 1, c));
      }
      console.log(`  First data row (row ${startRow + 1}): [${firstDataRow.join(', ')}${cols > 5 ? ', ...' : ''}]`);
    }
    console.log();
  }
}

function testIronCalcQuirks(model) {
  printSeparator('TEST 7: IronCalc Quirks & Edge Cases');

  const sheets = model.getWorksheetsProperties();
  console.log(`\n1. Worksheet Properties:`);
  for (let i = 0; i < sheets.length; i++) {
    console.log(`   Sheet ${i}: "${sheets[i].name}"`);
  }

  console.log(`\n2. Testing cell content types:`);
  const inputSheetIdx = sheets.findIndex(s => s.name === 'Input');

  // Test various cell types
  const testCells = [
    { desc: 'Number cell', row: 6, col: 4 },      // D6 - FicoScore input
    { desc: 'Formula cell', row: 69, col: 4 },    // D69 - LiabilitiesVerticalSum
    { desc: 'Empty cell', row: 1, col: 1 },       // A1 - likely empty
    { desc: 'Text cell', row: 7, col: 3 },        // C7 - Label
  ];

  for (const tc of testCells) {
    const formatted = model.getFormattedCellValue(inputSheetIdx, tc.row, tc.col);
    const content = model.getCellContent(inputSheetIdx, tc.row, tc.col);
    console.log(`   ${tc.desc} (row ${tc.row}, col ${tc.col}):`);
    console.log(`      Formatted: "${formatted}" (type: ${typeof formatted})`);
    console.log(`      Content: "${content}" (type: ${typeof content})`);
  }

  console.log(`\n3. Empty/null handling in arrays:`);
  // Get a row that might have empty cells
  const apiOutputIdx = sheets.findIndex(s => s.name === 'API_Output');
  console.log(`   Testing API_Output row 32 (last row of RateStackTable):`);

  for (let c = 7; c <= 11; c++) {
    const val = model.getFormattedCellValue(apiOutputIdx, 32, c);
    const content = model.getCellContent(apiOutputIdx, 32, c);
    console.log(`      Col ${numToCol(c)}: formatted="${val}", content="${content}"`);
  }
}

function testClassificationLogic(model, namedRanges) {
  printSeparator('TEST 8: Classification Summary');

  console.log('\nClassification results for all named ranges:\n');
  console.log(`${'Name'.padEnd(30)} ${'Type'.padEnd(12)} ${'Class'.padEnd(10)} ${'Notes'}`);
  console.log('-'.repeat(80));

  const summary = { INPUT: 0, OUTPUT: 0, UNKNOWN: 0 };

  for (const nr of namedRanges) {
    const parsed = parseCellReference(nr.formula);
    const rangeType = getRangeType(parsed);

    let classification = 'UNKNOWN';
    let notes = '';

    if (parsed) {
      const sheetIndex = getSheetIndex(model, parsed.sheet);

      if (rangeType === 'TABLE') {
        classification = 'OUTPUT';
        notes = 'TABLE always OUTPUT';
      } else if (sheetIndex >= 0) {
        const col = colToNum(parsed.startCol);
        const content = model.getCellContent(sheetIndex, parsed.startRow, col);
        const hasFormula = content?.startsWith('=');
        classification = hasFormula ? 'OUTPUT' : 'INPUT';
        notes = hasFormula ? `formula: ${content.slice(0, 30)}...` : 'no formula';
      } else {
        notes = 'sheet not found';
      }
    } else {
      notes = 'parse failed';
    }

    summary[classification]++;
    console.log(`${nr.name.padEnd(30)} ${rangeType.padEnd(12)} ${classification.padEnd(10)} ${notes}`);
  }

  console.log('\n' + '-'.repeat(80));
  console.log(`Total: ${namedRanges.length} named ranges`);
  console.log(`  INPUT: ${summary.INPUT}`);
  console.log(`  OUTPUT: ${summary.OUTPUT}`);
  console.log(`  UNKNOWN: ${summary.UNKNOWN}`);
}

// ============================================================
// Main
// ============================================================

async function main() {
  console.log('='.repeat(80));
  console.log('IRONCALC RANGE HANDLING ANALYSIS');
  console.log('Task 2 of Story 2.1');
  console.log('='.repeat(80));

  if (!fs.existsSync(XLSX_PATH)) {
    console.error(`ERROR: File not found: ${XLSX_PATH}`);
    process.exit(1);
  }

  console.log(`\nLoading: ${XLSX_PATH}`);
  const model = Model.fromXlsx(XLSX_PATH, "en", "UTC", "en");
  console.log('Workbook loaded successfully.');

  // Run all tests
  const namedRanges = testNamedRangeRetrieval(model);
  testParseCellReference(namedRanges);
  testReadingSingleCell(model, namedRanges);
  testReadingHorizontalArray(model, namedRanges);
  testReadingVerticalArray(model, namedRanges);
  testReadingTable(model, namedRanges);
  testIronCalcQuirks(model);
  testClassificationLogic(model, namedRanges);

  printSeparator('FINDINGS SUMMARY');
  console.log(`
Key Findings:

1. getDefinedNameList() returns all named ranges with:
   - name: string
   - formula: string (e.g., "Input!$D$6" or "API_Output!$G$5:$AK$32")
   - scope: number|null (sheet index or null for workbook scope)

2. parseCellReference() correctly handles:
   - Single cells: "Sheet!$A$1" -> {startCol:'A', startRow:1, endCol:'A', endRow:1, isRange:false}
   - Ranges: "Sheet!$A$1:$B$10" -> {startCol:'A', startRow:1, endCol:'B', endRow:10, isRange:true}

3. Range type detection works correctly:
   - SINGLE: 1x1
   - HORIZONTAL: 1xN
   - VERTICAL: Nx1
   - TABLE: NxM

4. Cell reading with getFormattedCellValue() and getCellContent():
   - Works with formula cells (returns calculated value / formula string)
   - Works with value cells (returns value / value string)
   - Empty cells return "" for both

5. Classification logic:
   - TABLE type: Always OUTPUT (no formula check needed)
   - SINGLE/HORIZONTAL/VERTICAL: Check first cell for formula
`);
}

main().catch(console.error);
