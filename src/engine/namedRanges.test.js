import { describe, it, expect } from 'vitest';
import { colToNum, getRangeType, parseCellReference } from './namedRanges.js';

describe('colToNum', () => {
  it('converts single letter columns', () => {
    expect(colToNum('A')).toBe(1);
    expect(colToNum('B')).toBe(2);
    expect(colToNum('Z')).toBe(26);
  });

  it('converts double letter columns', () => {
    expect(colToNum('AA')).toBe(27);
    expect(colToNum('AB')).toBe(28);
    expect(colToNum('AZ')).toBe(52);
    expect(colToNum('BA')).toBe(53);
  });

  it('converts triple letter columns', () => {
    expect(colToNum('AAA')).toBe(703);
    expect(colToNum('AK')).toBe(37); // Used in RateStackTable G5:AK32
  });
});

describe('getRangeType', () => {
  it('returns SINGLE for 1x1 range', () => {
    const parsed = parseCellReference('Input!$D$6');
    expect(getRangeType(parsed)).toBe('SINGLE');
  });

  it('returns HORIZONTAL for 1xN range', () => {
    const parsed = parseCellReference('Input!$D$65:$G$65');
    expect(getRangeType(parsed)).toBe('HORIZONTAL');
  });

  it('returns VERTICAL for Nx1 range', () => {
    const parsed = parseCellReference('Input!$C$60:$C$63');
    expect(getRangeType(parsed)).toBe('VERTICAL');
  });

  it('returns TABLE for NxM range', () => {
    const parsed = parseCellReference('API_Output!$G$5:$AK$32');
    expect(getRangeType(parsed)).toBe('TABLE');
  });

  it('returns UNKNOWN for null input', () => {
    expect(getRangeType(null)).toBe('UNKNOWN');
  });

  it('handles various TABLE dimensions', () => {
    // PriceAdjustmentTable 16x3
    const parsed1 = parseCellReference('API_Output!$A$7:$C$22');
    expect(getRangeType(parsed1)).toBe('TABLE');

    // EligiblityFailureReason 63x2
    const parsed2 = parseCellReference('API_Output!$A$30:$B$92');
    expect(getRangeType(parsed2)).toBe('TABLE');
  });
});

describe('parseCellReference', () => {
  it('parses single cell reference', () => {
    const parsed = parseCellReference('Input!$D$6');
    expect(parsed).toEqual({
      sheet: 'Input',
      startCol: 'D',
      startRow: 6,
      endCol: 'D',
      endRow: 6,
      isRange: false
    });
  });

  it('parses range reference', () => {
    const parsed = parseCellReference('API_Output!$G$5:$AK$32');
    expect(parsed).toEqual({
      sheet: 'API_Output',
      startCol: 'G',
      startRow: 5,
      endCol: 'AK',
      endRow: 32,
      isRange: true
    });
  });

  it('handles sheet names with underscores', () => {
    const parsed = parseCellReference('API_Output!$A$1');
    expect(parsed.sheet).toBe('API_Output');
  });

  it('returns null for invalid references', () => {
    expect(parseCellReference('')).toBe(null);
    expect(parseCellReference(null)).toBe(null);
    expect(parseCellReference('invalid')).toBe(null);
  });

  it('handles references with leading equals sign', () => {
    const parsed = parseCellReference('=Input!$D$6');
    expect(parsed).toEqual({
      sheet: 'Input',
      startCol: 'D',
      startRow: 6,
      endCol: 'D',
      endRow: 6,
      isRange: false
    });
  });
});
