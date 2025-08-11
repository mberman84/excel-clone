import { describe, test, expect } from 'vitest';
import {
  columnIndexToLabel,
  columnLabelToIndex,
  parseCellAddress,
  makeCellAddress,
  expandRange
} from '../cellAddresses';

describe('columnIndexToLabel', () => {
  test('converts single-letter column indices correctly', () => {
    expect(columnIndexToLabel(0)).toBe('A');
    expect(columnIndexToLabel(1)).toBe('B');
    expect(columnIndexToLabel(25)).toBe('Z');
  });

  test('converts double-letter column indices correctly', () => {
    expect(columnIndexToLabel(26)).toBe('AA');
    expect(columnIndexToLabel(27)).toBe('AB');
    expect(columnIndexToLabel(51)).toBe('AZ');
    expect(columnIndexToLabel(52)).toBe('BA');
    expect(columnIndexToLabel(701)).toBe('ZZ');
  });

  test('converts triple-letter column indices correctly', () => {
    expect(columnIndexToLabel(702)).toBe('AAA');
    expect(columnIndexToLabel(703)).toBe('AAB');
    expect(columnIndexToLabel(16383)).toBe('XFD'); // Excel's max column
  });
});

describe('columnLabelToIndex', () => {
  test('converts single-letter column labels correctly', () => {
    expect(columnLabelToIndex('A')).toBe(0);
    expect(columnLabelToIndex('B')).toBe(1);
    expect(columnLabelToIndex('Z')).toBe(25);
  });

  test('converts double-letter column labels correctly', () => {
    expect(columnLabelToIndex('AA')).toBe(26);
    expect(columnLabelToIndex('AB')).toBe(27);
    expect(columnLabelToIndex('AZ')).toBe(51);
    expect(columnLabelToIndex('BA')).toBe(52);
    expect(columnLabelToIndex('ZZ')).toBe(701);
  });

  test('converts triple-letter column labels correctly', () => {
    expect(columnLabelToIndex('AAA')).toBe(702);
    expect(columnLabelToIndex('AAB')).toBe(703);
    expect(columnLabelToIndex('XFD')).toBe(16383); // Excel's max column
  });

  test('handles lowercase labels correctly', () => {
    expect(columnLabelToIndex('a')).toBe(0);
    expect(columnLabelToIndex('aa')).toBe(26);
    expect(columnLabelToIndex('xfd')).toBe(16383);
  });

  test('throws error for invalid labels', () => {
    // Empty label returns -1 (no columns parsed)
    expect(columnLabelToIndex('')).toBe(-1);
    expect(() => columnLabelToIndex('123')).toThrow();
    expect(() => columnLabelToIndex('A1')).toThrow();
    expect(() => columnLabelToIndex('A!')).toThrow();
  });
});

describe('round-trip column conversions', () => {
  test('index to label to index works correctly', () => {
    for (const index of [0, 1, 25, 26, 27, 51, 52, 701, 702, 16383]) {
      const label = columnIndexToLabel(index);
      expect(columnLabelToIndex(label)).toBe(index);
    }
  });

  test('label to index to label works correctly', () => {
    for (const label of ['A', 'B', 'Z', 'AA', 'AB', 'AZ', 'BA', 'ZZ', 'AAA', 'XFD']) {
      const index = columnLabelToIndex(label);
      expect(columnIndexToLabel(index)).toBe(label);
    }
  });
});

describe('parseCellAddress', () => {
  test('parses simple cell addresses correctly', () => {
    expect(parseCellAddress('A1')).toEqual({ col: 0, row: 0 });
    expect(parseCellAddress('B2')).toEqual({ col: 1, row: 1 });
    expect(parseCellAddress('Z26')).toEqual({ col: 25, row: 25 });
  });

  test('parses addresses with multi-letter columns correctly', () => {
    expect(parseCellAddress('AA1')).toEqual({ col: 26, row: 0 });
    expect(parseCellAddress('AB10')).toEqual({ col: 27, row: 9 });
    expect(parseCellAddress('XFD1048576')).toEqual({ col: 16383, row: 1048575 }); // Excel's max cell
  });

  test('handles addresses with $ anchors correctly', () => {
    expect(parseCellAddress('$A$1')).toEqual({ col: 0, row: 0 });
    expect(parseCellAddress('$AA$100')).toEqual({ col: 26, row: 99 });
  });

  test('handles mixed case addresses correctly', () => {
    expect(parseCellAddress('a1')).toEqual({ col: 0, row: 0 });
    expect(parseCellAddress('aA1')).toEqual({ col: 26, row: 0 });
  });

  test('throws error for invalid addresses', () => {
    expect(() => parseCellAddress('')).toThrow();
    expect(() => parseCellAddress('A')).toThrow();
    expect(() => parseCellAddress('1')).toThrow();
    expect(() => parseCellAddress('A0')).toThrow(); // Row must be >= 1
    expect(() => parseCellAddress('A-1')).toThrow();
    expect(() => parseCellAddress('!1')).toThrow();
  });
});

describe('makeCellAddress', () => {
  test('creates simple cell addresses correctly', () => {
    expect(makeCellAddress(0, 0)).toBe('A1');
    expect(makeCellAddress(1, 1)).toBe('B2');
    expect(makeCellAddress(25, 25)).toBe('Z26');
  });

  test('creates addresses with multi-letter columns correctly', () => {
    expect(makeCellAddress(26, 0)).toBe('AA1');
    expect(makeCellAddress(27, 9)).toBe('AB10');
    expect(makeCellAddress(16383, 1048575)).toBe('XFD1048576'); // Excel's max cell
  });
});

describe('round-trip address conversions', () => {
  test('parse to make works correctly', () => {
    for (const addr of ['A1', 'B2', 'Z26', 'AA1', 'AB10', 'XFD1048576']) {
      const { col, row } = parseCellAddress(addr);
      expect(makeCellAddress(col, row)).toBe(addr);
    }
  });

  test('make to parse works correctly', () => {
    const testCases = [
      { col: 0, row: 0 },
      { col: 1, row: 1 },
      { col: 25, row: 25 },
      { col: 26, row: 0 },
      { col: 27, row: 9 },
      { col: 16383, row: 1048575 }
    ];
    
    for (const { col, row } of testCases) {
      const addr = makeCellAddress(col, row);
      const parsed = parseCellAddress(addr);
      expect(parsed.col).toBe(col);
      expect(parsed.row).toBe(row);
    }
  });
});

describe('expandRange', () => {
  test('expands single cell range correctly', () => {
    expect(expandRange('A1', 'A1')).toEqual(['A1']);
    expect(expandRange('B2', 'B2')).toEqual(['B2']);
  });

  test('expands horizontal ranges correctly', () => {
    expect(expandRange('A1', 'C1')).toEqual(['A1', 'B1', 'C1']);
    expect(expandRange('C1', 'A1')).toEqual(['A1', 'B1', 'C1']); // Order doesn't matter
  });

  test('expands vertical ranges correctly', () => {
    expect(expandRange('A1', 'A3')).toEqual(['A1', 'A2', 'A3']);
    expect(expandRange('A3', 'A1')).toEqual(['A1', 'A2', 'A3']); // Order doesn't matter
  });

  test('expands rectangular ranges correctly', () => {
    expect(expandRange('A1', 'B2')).toEqual(['A1', 'B1', 'A2', 'B2']);
    expect(expandRange('B2', 'A1')).toEqual(['A1', 'B1', 'A2', 'B2']); // Order doesn't matter
  });

  test('expands larger ranges correctly', () => {
    expect(expandRange('A1', 'C3')).toEqual([
      'A1', 'B1', 'C1',
      'A2', 'B2', 'C2',
      'A3', 'B3', 'C3'
    ]);
  });

  test('handles ranges with $ anchors correctly', () => {
    expect(expandRange('$A$1', '$C$3')).toEqual([
      'A1', 'B1', 'C1',
      'A2', 'B2', 'C2',
      'A3', 'B3', 'C3'
    ]);
  });

  test('handles ranges with mixed case correctly', () => {
    expect(expandRange('a1', 'c3')).toEqual([
      'A1', 'B1', 'C1',
      'A2', 'B2', 'C2',
      'A3', 'B3', 'C3'
    ]);
  });
});
