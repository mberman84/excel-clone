import { describe, test, expect, beforeEach, vi, afterEach } from 'vitest';
import { useStore } from '../store';
import { makeCellAddress } from '../utils/cellAddresses';

// @vitest-environment jsdom

// Mock localStorage
const localStorageMock = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: vi.fn((key: string) => store[key] || null),
    setItem: vi.fn((key: string, value: string) => {
      store[key] = value.toString();
    }),
    removeItem: vi.fn((key: string) => {
      delete store[key];
    }),
    clear: vi.fn(() => {
      store = {};
    }),
    store
  };
})();

Object.defineProperty(window, 'localStorage', {
  value: localStorageMock
});

// Helper to reset store between tests
const resetStore = () => {
  // Clear localStorage
  localStorageMock.clear();
  
  // Reset store to initial state by deleting all but one sheet
  const store = useStore.getState();
  
  // First, make sure we're on sheet 0
  store.setActiveSheet(0);
  
  // Delete all sheets except the first one
  while (store.workbook.sheets.length > 1) {
    store.deleteSheet(store.workbook.sheets.length - 1);
  }
  
  // Clear the first sheet
  store.clearSheet();
  
  // Reset selection and editing state
  store.selectCell(1, 1);
  if (store.editing.addr) {
    store.cancelEdit();
  }
};

// Helper to set up test data
const setupTestData = () => {
  const store = useStore.getState();
  
  // Set up markers in column A to track row movement
  store.setCellValue(makeCellAddress(0, 0), 'r1'); // A1
  store.setCellValue(makeCellAddress(0, 1), 'r2'); // A2
  store.setCellValue(makeCellAddress(0, 2), 'r3'); // A3
  
  // Set up test values in column C
  store.setCellValue(makeCellAddress(2, 0), '2');  // C1
  store.setCellValue(makeCellAddress(2, 1), '1');  // C2
  // C3 is intentionally left blank
};

describe('Sorting functionality', () => {
  beforeEach(() => {
    resetStore();
    vi.clearAllMocks();
    setupTestData();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  test('sortByColumn with ascending order (A→Z) places blanks at bottom', () => {
    const store = useStore.getState();
    
    // Sort column C (index 2) in ascending order
    store.sortByColumn(2, 'asc');
    
    // Get the current state after sorting
    const state = useStore.getState();
    const sheet = state.workbook.sheets[0];
    
    // Expected order: '1', '2', blank
    // Check markers in column A to verify row movement
    expect(sheet.cells[makeCellAddress(0, 0)]?.value).toBe('r2'); // A1 should now contain r2 (from row with value '1')
    expect(sheet.cells[makeCellAddress(0, 1)]?.value).toBe('r1'); // A2 should now contain r1 (from row with value '2')
    expect(sheet.cells[makeCellAddress(0, 2)]?.value).toBe('r3'); // A3 should still contain r3 (blank row)
    
    // Verify the values in column C are in the correct order
    expect(sheet.cells[makeCellAddress(2, 0)]?.value).toBe('1');  // C1 should be '1'
    expect(sheet.cells[makeCellAddress(2, 1)]?.value).toBe('2');  // C2 should be '2'
    expect(sheet.cells[makeCellAddress(2, 2)]?.value).toBeUndefined(); // C3 should be blank/undefined
  });

  test('sortByColumn with descending order (Z→A) places blanks at bottom', () => {
    const store = useStore.getState();
    
    // Sort column C (index 2) in descending order
    store.sortByColumn(2, 'desc');
    
    // Get the current state after sorting
    const state = useStore.getState();
    const sheet = state.workbook.sheets[0];
    
    // Expected order: '2', '1', blank
    // Check markers in column A to verify row movement
    expect(sheet.cells[makeCellAddress(0, 0)]?.value).toBe('r1'); // A1 should still contain r1 (from row with value '2')
    expect(sheet.cells[makeCellAddress(0, 1)]?.value).toBe('r2'); // A2 should still contain r2 (from row with value '1')
    expect(sheet.cells[makeCellAddress(0, 2)]?.value).toBe('r3'); // A3 should still contain r3 (blank row)
    
    // Verify the values in column C are in the correct order
    expect(sheet.cells[makeCellAddress(2, 0)]?.value).toBe('2');  // C1 should be '2'
    expect(sheet.cells[makeCellAddress(2, 1)]?.value).toBe('1');  // C2 should be '1'
    expect(sheet.cells[makeCellAddress(2, 2)]?.value).toBeUndefined(); // C3 should be blank/undefined
  });

  test('sortByColumn maintains stable order for blank cells', () => {
    const store = useStore.getState();
    
    // Add another blank row
    store.setCellValue(makeCellAddress(0, 3), 'r4'); // A4 marker
    // C4 is intentionally left blank
    
    // Sort column C (index 2) in ascending order
    store.sortByColumn(2, 'asc');
    
    // Get the current state after sorting
    const state = useStore.getState();
    const sheet = state.workbook.sheets[0];
    
    // Expected order: '1', '2', blank, blank (with r3 before r4 due to stable sort)
    // Check that blank rows maintain their relative order
    expect(sheet.cells[makeCellAddress(0, 2)]?.value).toBe('r3'); // A3 should contain r3 (first blank row)
    expect(sheet.cells[makeCellAddress(0, 3)]?.value).toBe('r4'); // A4 should contain r4 (second blank row)
  });

  test('sortByColumn handles numeric values correctly', () => {
    const store = useStore.getState();
    
    // Set up numeric test data (replacing strings with actual numbers)
    store.setCellValue(makeCellAddress(2, 0), '10');  // C1
    store.setCellValue(makeCellAddress(2, 1), '2');   // C2
    
    // Sort column C (index 2) in ascending order
    store.sortByColumn(2, 'asc');
    
    // Get the current state after sorting
    const state = useStore.getState();
    const sheet = state.workbook.sheets[0];
    
    // Expected numeric order: 2, 10 (not lexicographic '10', '2')
    expect(sheet.cells[makeCellAddress(0, 0)]?.value).toBe('r2'); // A1 should contain r2 (from row with value '2')
    expect(sheet.cells[makeCellAddress(0, 1)]?.value).toBe('r1'); // A2 should contain r1 (from row with value '10')
    
    expect(sheet.cells[makeCellAddress(2, 0)]?.value).toBe('2');  // C1 should be '2'
    expect(sheet.cells[makeCellAddress(2, 1)]?.value).toBe('10'); // C2 should be '10'
  });
});
