import { describe, it, expect } from 'vitest';
import { coerceValue, columnToNumber } from './cellReader.js';

describe('coerceValue', () => {
  describe('empty/null handling', () => {
    it('converts empty string to null', () => {
      expect(coerceValue('')).toBe(null);
    });

    it('converts null to null', () => {
      expect(coerceValue(null)).toBe(null);
    });

    it('converts undefined to null', () => {
      expect(coerceValue(undefined)).toBe(null);
    });
  });

  describe('numeric conversion', () => {
    it('converts integer strings to numbers', () => {
      expect(coerceValue('123')).toBe(123);
      expect(coerceValue('0')).toBe(0);
      expect(coerceValue('-42')).toBe(-42);
    });

    it('converts decimal strings to numbers', () => {
      expect(coerceValue('123.45')).toBe(123.45);
      expect(coerceValue('0.5')).toBe(0.5);
      expect(coerceValue('-0.001')).toBe(-0.001);
    });

    it('handles whitespace around numbers', () => {
      expect(coerceValue(' 123 ')).toBe(123);
      expect(coerceValue('  45.67  ')).toBe(45.67);
    });
  });

  describe('boolean conversion', () => {
    it('converts TRUE to boolean true (case insensitive)', () => {
      expect(coerceValue('true')).toBe(true);
      expect(coerceValue('TRUE')).toBe(true);
      expect(coerceValue('True')).toBe(true);
    });

    it('converts FALSE to boolean false (case insensitive)', () => {
      expect(coerceValue('false')).toBe(false);
      expect(coerceValue('FALSE')).toBe(false);
      expect(coerceValue('False')).toBe(false);
    });
  });

  describe('Excel error handling', () => {
    it('converts #N/A to null', () => {
      expect(coerceValue('#N/A')).toBe(null);
    });

    it('converts #REF! to null', () => {
      expect(coerceValue('#REF!')).toBe(null);
    });

    it('converts #VALUE! to null', () => {
      expect(coerceValue('#VALUE!')).toBe(null);
    });

    it('converts #DIV/0! to null', () => {
      expect(coerceValue('#DIV/0!')).toBe(null);
    });

    it('converts #NAME? to null', () => {
      expect(coerceValue('#NAME?')).toBe(null);
    });

    it('converts #NULL! to null', () => {
      expect(coerceValue('#NULL!')).toBe(null);
    });

    it('converts #NUM! to null', () => {
      expect(coerceValue('#NUM!')).toBe(null);
    });
  });

  describe('string preservation', () => {
    it('keeps text strings as strings', () => {
      expect(coerceValue('YES')).toBe('YES');
      expect(coerceValue('Hello World')).toBe('Hello World');
      expect(coerceValue('Loan Amount')).toBe('Loan Amount');
    });

    it('does not convert strings that look partially numeric', () => {
      expect(coerceValue('123abc')).toBe('123abc');
      expect(coerceValue('$100')).toBe('$100');
      expect(coerceValue('100%')).toBe('100%');
    });
  });
});

describe('columnToNumber', () => {
  it('converts single letter columns', () => {
    expect(columnToNumber('A')).toBe(1);
    expect(columnToNumber('B')).toBe(2);
    expect(columnToNumber('Z')).toBe(26);
  });

  it('converts double letter columns', () => {
    expect(columnToNumber('AA')).toBe(27);
    expect(columnToNumber('AK')).toBe(37);
  });
});
