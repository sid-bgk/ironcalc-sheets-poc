import XLSX from 'xlsx';
import { resolve } from 'path';

const inputFile = process.argv[2] || '../DSCR_Complete_Pricing_Engine_Dev_NoArrayTesting.xlsx';
const outputFile = process.argv[3] || '../DSCR_Complete_Pricing_Engine_Dev_NoArray_FIXED.xlsx';

console.log(`\n=== Remove CSE Array Flags ===`);
console.log(`Input:  ${inputFile}`);
console.log(`Output: ${outputFile}\n`);

try {
  const workbook = XLSX.readFile(resolve(inputFile));

  let fixedCount = 0;
  const fixedCells = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    for (const cellAddress in sheet) {
      if (cellAddress.startsWith('!')) continue;

      const cell = sheet[cellAddress];

      // Check for array formula (cell.F indicates array range)
      if (cell && cell.f && cell.F) {
        // Store info before fixing
        fixedCells.push({
          sheet: sheetName,
          cell: cellAddress,
          formula: cell.f,
          oldValue: cell.v
        });

        // Remove the array flag (F property)
        delete cell.F;
        fixedCount++;
      }
    }
  }

  // Write the modified workbook
  XLSX.writeFile(workbook, resolve(outputFile));

  console.log(`✅ Fixed ${fixedCount} array formulas:\n`);

  for (const item of fixedCells) {
    console.log(`  ${item.sheet}!${item.cell} - Value: ${item.oldValue}`);
  }

  console.log(`\n✅ Saved to: ${outputFile}`);
  console.log(`\nYou can now compare both files in Excel to verify the values are identical.`);

} catch (error) {
  console.error(`Error: ${error.message}`);
  process.exit(1);
}
