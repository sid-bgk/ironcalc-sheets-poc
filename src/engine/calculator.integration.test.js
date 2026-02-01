import { describe, it, expect, beforeAll } from 'vitest';
import { Model } from '@ironcalc/nodejs';
import {
  setInputValue,
  setInputValues,
  getOutputValue,
  getAllOutputValues,
  executeCalculation,
  clearCache
} from './calculator.js';
import { classifyAllNamedRanges } from './namedRanges.js';

const XLSX_PATH = './DSCR_NoArrayFormulas_Testing.xlsx';

describe('Integration Tests with Real Excel Data', () => {
  let model;

  beforeAll(() => {
    model = Model.fromXlsx(XLSX_PATH, 'en', 'UTC', 'en');
    clearCache();
  });

  describe('Named Range Classification', () => {
    it('classifies all named ranges correctly', () => {
      const result = classifyAllNamedRanges(model);

      expect(result.inputs.length).toBeGreaterThan(0);
      expect(result.outputs.length).toBeGreaterThan(0);
      expect(result.unknown.length).toBe(0);

      // Check that we have the expected INPUT ranges
      const inputNames = result.inputs.map(i => i.name);
      expect(inputNames).toContain('FicoScore');
      expect(inputNames).toContain('LoanAmount');
      expect(inputNames).toContain('LiabilitiesHorizontal');
      expect(inputNames).toContain('LiabilitiesVertical');

      // Check that we have the expected OUTPUT ranges
      const outputNames = result.outputs.map(o => o.name);
      expect(outputNames).toContain('LoanEligiblity');
      expect(outputNames).toContain('RateStackTable');
      expect(outputNames).toContain('PriceAdjustmentTable');
      expect(outputNames).toContain('EligiblityFailureReason');
    });

    it('classifies TABLE ranges correctly', () => {
      const result = classifyAllNamedRanges(model);

      const rateStackTable = result.outputs.find(o => o.name === 'RateStackTable');
      expect(rateStackTable).toBeDefined();
      expect(rateStackTable.rangeType).toBe('TABLE');
      expect(rateStackTable.type).toBe('OUTPUT'); // TABLE is always OUTPUT
      expect(rateStackTable.rows).toBe(28);
      expect(rateStackTable.cols).toBe(31);

      const priceAdjTable = result.outputs.find(o => o.name === 'PriceAdjustmentTable');
      expect(priceAdjTable).toBeDefined();
      expect(priceAdjTable.rangeType).toBe('TABLE');
      expect(priceAdjTable.rows).toBe(16);
      expect(priceAdjTable.cols).toBe(3);
    });

    it('classifies HORIZONTAL and VERTICAL INPUT ranges correctly', () => {
      const result = classifyAllNamedRanges(model);

      const horizontal = result.inputs.find(i => i.name === 'LiabilitiesHorizontal');
      expect(horizontal).toBeDefined();
      expect(horizontal.rangeType).toBe('HORIZONTAL');
      expect(horizontal.type).toBe('INPUT');
      expect(horizontal.rows).toBe(1);
      expect(horizontal.cols).toBe(4);

      const vertical = result.inputs.find(i => i.name === 'LiabilitiesVertical');
      expect(vertical).toBeDefined();
      expect(vertical.rangeType).toBe('VERTICAL');
      expect(vertical.type).toBe('INPUT');
      expect(vertical.rows).toBe(4);
      expect(vertical.cols).toBe(1);
    });
  });

  describe('Task 5.1: RateStackTable (28×31) Output', () => {
    it('returns complete structured TABLE output', () => {
      const result = getOutputValue(model, 'RateStackTable');

      expect(result).toBeDefined();
      expect(result.type).toBe('TABLE');
      expect(result.rows).toBe(27); // 28 rows - 1 header = 27 data rows
      expect(result.cols).toBe(31);
      expect(result.headers).toHaveLength(31);
      expect(result.data).toHaveLength(27);

      // Each data row should have 31 columns
      for (const row of result.data) {
        expect(row).toHaveLength(31);
      }
    });
  });

  describe('Task 5.2: PriceAdjustmentTable (16×3) Output', () => {
    it('returns complete structured TABLE output', () => {
      const result = getOutputValue(model, 'PriceAdjustmentTable');

      expect(result).toBeDefined();
      expect(result.type).toBe('TABLE');
      expect(result.rows).toBe(15); // 16 rows - 1 header = 15 data rows
      expect(result.cols).toBe(3);
      expect(result.headers).toHaveLength(3);
      expect(result.data).toHaveLength(15);

      // Each data row should have 3 columns
      for (const row of result.data) {
        expect(row).toHaveLength(3);
      }
    });
  });

  describe('Task 5.3: EligiblityFailureReason (63×2) Output', () => {
    it('returns complete structured TABLE output', () => {
      const result = getOutputValue(model, 'EligiblityFailureReason');

      expect(result).toBeDefined();
      expect(result.type).toBe('TABLE');
      expect(result.rows).toBe(62); // 63 rows - 1 header = 62 data rows
      expect(result.cols).toBe(2);
      expect(result.headers).toHaveLength(2);
      expect(result.data).toHaveLength(62);

      // Each data row should have 2 columns
      for (const row of result.data) {
        expect(row).toHaveLength(2);
      }
    });
  });

  describe('Task 5.4: LiabilitiesHorizontal INPUT with Array', () => {
    it('sets HORIZONTAL array values correctly', () => {
      const testValues = [100, 200, 300, 400];

      // Set array input
      setInputValue(model, 'LiabilitiesHorizontal', testValues);

      // Recalculate
      model.evaluate();

      // LiabilitiesHorizontalSum should reflect the sum
      const sum = getOutputValue(model, 'LiabilitiesHorizontalSum');
      expect(sum).toBe(1000); // 100 + 200 + 300 + 400
    });

    it('handles shorter arrays by padding with empty values', () => {
      // Set only 3 values for a 4-cell range
      setInputValue(model, 'LiabilitiesHorizontal', [100, 200, 300]);
      model.evaluate();

      // Sum should be 600 (100+200+300+0)
      const sum = getOutputValue(model, 'LiabilitiesHorizontalSum');
      expect(sum).toBe(600);
    });

    it('handles longer arrays by truncating to range size', () => {
      // Set 5 values for a 4-cell range
      setInputValue(model, 'LiabilitiesHorizontal', [100, 200, 300, 400, 500]);
      model.evaluate();

      // Sum should be 1000 (100+200+300+400), 500 is ignored
      const sum = getOutputValue(model, 'LiabilitiesHorizontalSum');
      expect(sum).toBe(1000);
    });
  });

  describe('Task 5.5: LiabilitiesVertical INPUT with Array', () => {
    it('sets VERTICAL array values correctly', () => {
      const testValues = [50, 100, 150, 200];

      // Set array input
      setInputValue(model, 'LiabilitiesVertical', testValues);

      // Recalculate
      model.evaluate();

      // LiabilitiesVerticalSum should reflect the sum
      const sum = getOutputValue(model, 'LiabilitiesVerticalSum');
      expect(sum).toBe(500); // 50 + 100 + 150 + 200
    });

    it('handles shorter arrays by padding with empty values', () => {
      // Set only 3 values for a 4-cell range
      setInputValue(model, 'LiabilitiesVertical', [50, 100, 150]);
      model.evaluate();

      // Sum should be 300 (50+100+150+0)
      const sum = getOutputValue(model, 'LiabilitiesVerticalSum');
      expect(sum).toBe(300);
    });

    it('handles longer arrays by truncating to range size', () => {
      // Set 5 values for a 4-cell range
      setInputValue(model, 'LiabilitiesVertical', [50, 100, 150, 200, 250]);
      model.evaluate();

      // Sum should be 500 (50+100+150+200), 250 is ignored
      const sum = getOutputValue(model, 'LiabilitiesVerticalSum');
      expect(sum).toBe(500);
    });
  });

  describe('Task 5.6: Regression - Single Cell Behavior Unchanged', () => {
    it('gets SINGLE OUTPUT values correctly', () => {
      const result = getOutputValue(model, 'LoanEligiblity');
      // Should be a scalar value, not an array or object
      expect(typeof result === 'string' || typeof result === 'number' || result === null).toBe(true);
    });

    it('sets SINGLE INPUT values correctly', () => {
      // Set a single input
      setInputValue(model, 'FicoScore', 750);
      setInputValue(model, 'LoanAmount', 500000);

      model.evaluate();

      // Should not throw and should complete calculation
      const outputs = getAllOutputValues(model);
      expect(outputs).toBeDefined();
      expect(outputs.LoanEligiblity).toBeDefined();
    });
  });

  describe('Task 5.7: End-to-End API Flow with Arrays', () => {
    it('executeCalculation works with array inputs and saves result', () => {
      const inputs = {
        FicoScore: 720,
        LoanAmount: 450000,
        LiabilitiesHorizontal: [150, 250, 350, 250],
        LiabilitiesVertical: [75, 125, 175, 125]
      };

      const { outputs, resultFile } = executeCalculation(model, inputs);

      // Verify outputs exist
      expect(outputs).toBeDefined();
      expect(outputs.LoanEligiblity).toBeDefined();

      // Verify sum outputs reflect array inputs
      expect(outputs.LiabilitiesHorizontalSum).toBe(1000); // 150+250+350+250
      expect(outputs.LiabilitiesVerticalSum).toBe(500); // 75+125+175+125

      // Verify TABLE outputs are structured
      expect(outputs.RateStackTable).toBeDefined();
      expect(outputs.RateStackTable.type).toBe('TABLE');
      expect(outputs.RateStackTable.headers).toBeDefined();
      expect(outputs.RateStackTable.data).toBeDefined();

      // Verify result file was saved
      expect(resultFile).toBeDefined();
      expect(resultFile).toMatch(/results[\/\\]result_.*\.xlsx$/);
    });
  });

  describe('AC4: Type Coercion in Outputs', () => {
    it('coerces numeric values to numbers', () => {
      const result = getOutputValue(model, 'LiabilitiesHorizontalSum');
      expect(typeof result).toBe('number');
    });

    it('handles mixed types in TABLE output', () => {
      const result = getOutputValue(model, 'RateStackTable');

      // Headers should contain strings or numbers
      for (const header of result.headers) {
        expect(header === null || typeof header === 'string' || typeof header === 'number').toBe(true);
      }

      // Data cells should contain strings, numbers, or null
      for (const row of result.data) {
        for (const cell of row) {
          expect(cell === null || typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean').toBe(true);
        }
      }
    });
  });
});
