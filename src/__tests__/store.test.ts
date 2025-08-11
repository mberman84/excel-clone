import { describe, test, expect, beforeEach, vi, afterEach } from 'vitest';
import { useStore, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../store';
import { Sheet } from '../types';

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

describe('store', () => {
  beforeEach(() => {
    resetStore();
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('initial state', () => {
    test('has expected initial values', () => {
      const state = useStore.getState();
      
      // Check workbook structure
      expect(state.workbook.sheets.length).toBe(1);
      expect(state.workbook.activeIndex).toBe(0);
      
      // Check selection
      expect(state.selection).toEqual({ row: 1, col: 1, endRow: 1, endCol: 1 });
      
      // Check editing state
      expect(state.editing).toEqual({ addr: null, draft: '' });
      
      // Check history
      expect(state.past).toEqual([]);
      expect(state.future).toEqual([]);
    });

    test('initial sheet has expected structure', () => {
      const { workbook } = useStore.getState();
      const sheet = workbook.sheets[0];
      
      expect(sheet.id).toBeDefined();
      expect(sheet.name).toBe('Sheet 1');
      expect(sheet.cells).toEqual({});
      expect(sheet.colWidths).toEqual([]);
      expect(sheet.rowHeights).toEqual([]);
    });
  });

  describe('cell selection and editing', () => {
    test('selectCell updates selection state', () => {
      const store = useStore.getState();
      
      store.selectCell(3, 4);
      
      expect(store.selection).toEqual({ row: 3, col: 4, endRow: 3, endCol: 4 });
    });

    test('startEdit sets editing state', () => {
      const store = useStore.getState();
      const addr = 'B2';
      
      store.startEdit(addr);
      
      expect(store.editing).toEqual({ addr, draft: '' });
    });

    test('startEdit initializes draft with existing cell value', () => {
      const store = useStore.getState();
      const addr = 'B2';
      const value = 'test value';
      
      // Set a cell value first
      store.setCellValue(addr, value);
      
      // Then start editing
      store.startEdit(addr);
      
      expect(store.editing).toEqual({ addr, draft: value });
    });

    test('setDraft updates draft value', () => {
      const store = useStore.getState();
      const addr = 'B2';
      const draft = 'new draft value';
      
      store.startEdit(addr);
      store.setDraft(draft);
      
      expect(store.editing).toEqual({ addr, draft });
    });

    test('commitEdit saves draft to cell and clears editing state', () => {
      const store = useStore.getState();
      const addr = 'B2';
      const value = 'test value';
      
      store.startEdit(addr);
      store.setDraft(value);
      store.commitEdit();
      
      // Check editing state is cleared
      expect(store.editing).toEqual({ addr: null, draft: '' });
      
      // Check cell value was saved
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value);
      
      // Check past was updated
      expect(store.past.length).toBe(1);
    });

    test('commitEdit removes empty cells', () => {
      const store = useStore.getState();
      const addr = 'B2';
      
      // First add a cell
      store.setCellValue(addr, 'test');
      expect(store.workbook.sheets[0].cells[addr]).toBeDefined();
      
      // Then edit it to empty and commit
      store.startEdit(addr);
      store.setDraft('');
      store.commitEdit();
      
      // Cell should be removed
      expect(store.workbook.sheets[0].cells[addr]).toBeUndefined();
    });

    test('cancelEdit clears editing state without saving', () => {
      const store = useStore.getState();
      const addr = 'B2';
      const originalValue = 'original value';
      const draftValue = 'draft value';
      
      // Set original value
      store.setCellValue(addr, originalValue);
      
      // Start editing with new draft
      store.startEdit(addr);
      store.setDraft(draftValue);
      
      // Cancel edit
      store.cancelEdit();
      
      // Check editing state is cleared
      expect(store.editing).toEqual({ addr: null, draft: '' });
      
      // Check cell value remains unchanged
      expect(store.workbook.sheets[0].cells[addr].value).toBe(originalValue);
    });

    test('setCellValue directly updates cell without editing state', () => {
      const store = useStore.getState();
      const addr = 'C3';
      const value = 'direct value';
      
      store.setCellValue(addr, value);
      
      // Check cell value was saved
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value);
      
      // Check editing state remains empty
      expect(store.editing).toEqual({ addr: null, draft: '' });
      
      // Check past was updated
      expect(store.past.length).toBe(1);
    });

    test('setCellValue removes empty cells', () => {
      const store = useStore.getState();
      const addr = 'C3';
      
      // First add a cell
      store.setCellValue(addr, 'test');
      expect(store.workbook.sheets[0].cells[addr]).toBeDefined();
      
      // Then set it to empty
      store.setCellValue(addr, '');
      
      // Cell should be removed
      expect(store.workbook.sheets[0].cells[addr]).toBeUndefined();
    });
  });

  describe('undo/redo', () => {
    test('undo reverts cell changes', () => {
      const store = useStore.getState();
      const addr = 'D4';
      const value = 'test value';
      
      // Make a change
      store.setCellValue(addr, value);
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value);
      
      // Undo the change
      store.undo();
      
      // Cell should be gone
      expect(store.workbook.sheets[0].cells[addr]).toBeUndefined();
      
      // Past should be empty, future should have one state
      expect(store.past.length).toBe(0);
      expect(store.future.length).toBe(1);
    });

    test('redo reapplies undone changes', () => {
      const store = useStore.getState();
      const addr = 'D4';
      const value = 'test value';
      
      // Make a change
      store.setCellValue(addr, value);
      
      // Undo the change
      store.undo();
      
      // Redo the change
      store.redo();
      
      // Cell should have the value again
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value);
      
      // Past should have one state, future should be empty
      expect(store.past.length).toBe(1);
      expect(store.future.length).toBe(0);
    });

    test('making a new change after undo clears the future', () => {
      const store = useStore.getState();
      const addr1 = 'D4';
      const addr2 = 'E5';
      const value1 = 'value 1';
      const value2 = 'value 2';
      
      // Make a change
      store.setCellValue(addr1, value1);
      
      // Undo the change
      store.undo();
      
      // Make a different change
      store.setCellValue(addr2, value2);
      
      // Future should be empty
      expect(store.future.length).toBe(0);
      
      // Past should have one state
      expect(store.past.length).toBe(1);
      
      // Only the second cell should exist
      expect(store.workbook.sheets[0].cells[addr1]).toBeUndefined();
      expect(store.workbook.sheets[0].cells[addr2].value).toBe(value2);
    });
  });

  describe('cell formatting', () => {
    test('toggleFormat toggles boolean format properties', () => {
      const store = useStore.getState();
      const addr = 'E5';
      
      // Toggle bold on
      store.toggleFormat(addr, 'bold');
      expect(store.workbook.sheets[0].cells[addr].format?.bold).toBe(true);
      
      // Toggle bold off
      store.toggleFormat(addr, 'bold');
      expect(store.workbook.sheets[0].cells[addr].format?.bold).toBe(false);
      
      // Toggle italic on
      store.toggleFormat(addr, 'italic');
      expect(store.workbook.sheets[0].cells[addr].format?.italic).toBe(true);
      
      // Toggle underline on
      store.toggleFormat(addr, 'underline');
      expect(store.workbook.sheets[0].cells[addr].format?.underline).toBe(true);
      
      // Check both are still on
      expect(store.workbook.sheets[0].cells[addr].format?.italic).toBe(true);
      expect(store.workbook.sheets[0].cells[addr].format?.underline).toBe(true);
    });

    test('setTextColor sets text color', () => {
      const store = useStore.getState();
      const addr = 'F6';
      const color = '#ff0000';
      
      store.setTextColor(addr, color);
      
      expect(store.workbook.sheets[0].cells[addr].format?.textColor).toBe(color);
    });

    test('setFillColor sets fill color', () => {
      const store = useStore.getState();
      const addr = 'G7';
      const color = '#00ff00';
      
      store.setFillColor(addr, color);
      
      expect(store.workbook.sheets[0].cells[addr].format?.fillColor).toBe(color);
    });

    test('formatting works on empty cells', () => {
      const store = useStore.getState();
      const addr = 'H8';
      
      // Format an empty cell
      store.toggleFormat(addr, 'bold');
      
      // Cell should exist with empty value and format
      expect(store.workbook.sheets[0].cells[addr].value).toBe('');
      expect(store.workbook.sheets[0].cells[addr].format?.bold).toBe(true);
    });

    test('formatting preserves existing cell value', () => {
      const store = useStore.getState();
      const addr = 'I9';
      const value = 'test value';
      
      // Set a value
      store.setCellValue(addr, value);
      
      // Add formatting
      store.toggleFormat(addr, 'bold');
      store.setTextColor(addr, '#ff0000');
      
      // Value should be preserved
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value);
      
      // Format should be applied
      expect(store.workbook.sheets[0].cells[addr].format?.bold).toBe(true);
      expect(store.workbook.sheets[0].cells[addr].format?.textColor).toBe('#ff0000');
    });
  });

  describe('column and row sizing', () => {
    test('setColWidth sets column width', () => {
      const store = useStore.getState();
      const colIndex = 2;
      const width = 200;
      
      store.setColWidth(colIndex, width);
      
      expect(store.workbook.sheets[0].colWidths[colIndex]).toBe(width);
    });

    test('setRowHeight sets row height', () => {
      const store = useStore.getState();
      const rowIndex = 3;
      const height = 50;
      
      store.setRowHeight(rowIndex, height);
      
      expect(store.workbook.sheets[0].rowHeights[rowIndex]).toBe(height);
    });

    test('setColWidth clamps to minimum width', () => {
      const store = useStore.getState();
      const colIndex = 4;
      const tooSmall = 10; // Below minimum of 40
      
      store.setColWidth(colIndex, tooSmall);
      
      // Should be clamped to minimum
      expect(store.workbook.sheets[0].colWidths[colIndex]).toBe(40);
    });

    test('setRowHeight clamps to minimum height', () => {
      const store = useStore.getState();
      const rowIndex = 5;
      const tooSmall = 5; // Below minimum of 18
      
      store.setRowHeight(rowIndex, tooSmall);
      
      // Should be clamped to minimum
      expect(store.workbook.sheets[0].rowHeights[rowIndex]).toBe(18);
    });

    test('negative indices are ignored', () => {
      const store = useStore.getState();
      
      store.setColWidth(-1, 100);
      store.setRowHeight(-1, 100);
      
      // Should not have entries for negative indices
      expect(store.workbook.sheets[0].colWidths[-1]).toBeUndefined();
      expect(store.workbook.sheets[0].rowHeights[-1]).toBeUndefined();
    });
  });

  describe('sheet operations', () => {
    test('addSheet adds a new sheet', () => {
      const store = useStore.getState();
      
      store.addSheet();
      
      expect(store.workbook.sheets.length).toBe(2);
      expect(store.workbook.activeIndex).toBe(1); // New sheet becomes active
      expect(store.workbook.sheets[1].name).toBe('Sheet 2');
      expect(store.workbook.sheets[1].cells).toEqual({});
    });

    test('renameSheet changes sheet name', () => {
      const store = useStore.getState();
      const newName = 'Renamed Sheet';
      
      store.renameSheet(0, newName);
      
      expect(store.workbook.sheets[0].name).toBe(newName);
    });

    test('deleteSheet removes a sheet', () => {
      const store = useStore.getState();
      
      // Add a second sheet
      store.addSheet();
      expect(store.workbook.sheets.length).toBe(2);
      
      // Delete the second sheet
      store.deleteSheet(1);
      
      expect(store.workbook.sheets.length).toBe(1);
      expect(store.workbook.activeIndex).toBe(0);
    });

    test('cannot delete the last sheet', () => {
      const store = useStore.getState();
      
      // Try to delete the only sheet
      store.deleteSheet(0);
      
      // Should still have one sheet
      expect(store.workbook.sheets.length).toBe(1);
    });

    test('setActiveSheet changes active sheet', () => {
      const store = useStore.getState();
      
      // Add a second sheet
      store.addSheet();
      expect(store.workbook.activeIndex).toBe(1);
      
      // Switch back to first sheet
      store.setActiveSheet(0);
      
      expect(store.workbook.activeIndex).toBe(0);
    });

    test('sheet data is independent between sheets', () => {
      const store = useStore.getState();
      const addr = 'A1';
      const value1 = 'Sheet 1 Value';
      const value2 = 'Sheet 2 Value';
      
      // Set value in first sheet
      store.setCellValue(addr, value1);
      
      // Add a second sheet
      store.addSheet();
      
      // Set different value in second sheet
      store.setCellValue(addr, value2);
      
      // Switch back to first sheet
      store.setActiveSheet(0);
      
      // Check values are independent
      expect(store.workbook.sheets[0].cells[addr].value).toBe(value1);
      
      // Switch to second sheet
      store.setActiveSheet(1);
      expect(store.workbook.sheets[1].cells[addr].value).toBe(value2);
    });

    test('sheet formatting is independent between sheets', () => {
      const store = useStore.getState();
      const addr = 'B2';
      
      // Format in first sheet
      store.toggleFormat(addr, 'bold');
      
      // Add a second sheet
      store.addSheet();
      
      // Format differently in second sheet
      store.toggleFormat(addr, 'italic');
      
      // Switch back to first sheet
      store.setActiveSheet(0);
      
      // Check formats are independent
      expect(store.workbook.sheets[0].cells[addr].format?.bold).toBe(true);
      expect(store.workbook.sheets[0].cells[addr].format?.italic).toBeUndefined();
      
      // Switch to second sheet
      store.setActiveSheet(1);
      expect(store.workbook.sheets[1].cells[addr].format?.bold).toBeUndefined();
      expect(store.workbook.sheets[1].cells[addr].format?.italic).toBe(true);
    });

    test('sheet sizing is independent between sheets', () => {
      const store = useStore.getState();
      const colIndex = 1;
      const rowIndex = 2;
      
      // Set sizes in first sheet
      store.setColWidth(colIndex, 200);
      store.setRowHeight(rowIndex, 50);
      
      // Add a second sheet
      store.addSheet();
      
      // Set different sizes in second sheet
      store.setColWidth(colIndex, 300);
      store.setRowHeight(rowIndex, 100);
      
      // Check sizes are independent
      expect(store.workbook.sheets[0].colWidths[colIndex]).toBe(200);
      expect(store.workbook.sheets[0].rowHeights[rowIndex]).toBe(50);
      
      expect(store.workbook.sheets[1].colWidths[colIndex]).toBe(300);
      expect(store.workbook.sheets[1].rowHeights[rowIndex]).toBe(100);
    });

    test('clearSheet preserves sheet id and name', () => {
      const store = useStore.getState();
      const addr = 'C3';
      
      // Add some data
      store.setCellValue(addr, 'test');
      store.setColWidth(1, 200);
      
      // Rename the sheet
      const customName = 'Custom Name';
      store.renameSheet(0, customName);
      
      // Remember the id
      const sheetId = store.workbook.sheets[0].id;
      
      // Clear the sheet
      store.clearSheet();
      
      // Check data is cleared but id and name preserved
      expect(store.workbook.sheets[0].id).toBe(sheetId);
      expect(store.workbook.sheets[0].name).toBe(customName);
      expect(store.workbook.sheets[0].cells).toEqual({});
      expect(store.workbook.sheets[0].colWidths).toEqual([]);
      expect(store.workbook.sheets[0].rowHeights).toEqual([]);
    });
  });

  describe('data utilities', () => {
    test('getUsedRange returns correct range', () => {
      const store = useStore.getState();
      
      // Add cells at various positions
      store.setCellValue('A1', 'top-left');
      store.setCellValue('C5', 'bottom-right');
      
      const range = store.getUsedRange();
      
      // Should include both cells plus some padding
      expect(range.maxRow).toBeGreaterThanOrEqual(5);
      expect(range.maxCol).toBeGreaterThanOrEqual(3);
    });

    test('getUsedRange has minimum size', () => {
      const store = useStore.getState();
      
      // Empty sheet
      const range = store.getUsedRange();
      
      // Should have minimum size
      expect(range.maxRow).toBeGreaterThanOrEqual(20);
      expect(range.maxCol).toBeGreaterThanOrEqual(10);
    });

    test('toAOA converts sheet to array of arrays', () => {
      const store = useStore.getState();
      
      // Add some cells
      store.setCellValue('A1', 'A1 value');
      store.setCellValue('B2', 'B2 value');
      
      const aoa = store.toAOA();
      
      // Should include both cells
      expect(aoa[0][0]).toBe('A1 value');
      expect(aoa[1][1]).toBe('B2 value');
      
      // Empty cells should be empty strings
      expect(aoa[0][1]).toBe('');
      expect(aoa[1][0]).toBe('');
    });

    test('toAOAAll converts all sheets to array of arrays', () => {
      const store = useStore.getState();
      
      // Add data to first sheet
      store.setCellValue('A1', 'Sheet 1 A1');
      
      // Add a second sheet with data
      store.addSheet();
      store.setCellValue('B2', 'Sheet 2 B2');
      
      const allData = store.toAOAAll();
      
      // Should have data for both sheets
      expect(allData.length).toBe(2);
      expect(allData[0].name).toBe('Sheet 1');
      expect(allData[1].name).toBe('Sheet 2');
      
      // Check data
      expect(allData[0].data[0][0]).toBe('Sheet 1 A1');
      expect(allData[1].data[1][1]).toBe('Sheet 2 B2');
    });

    test('fromAOA imports data to active sheet', () => {
      const store = useStore.getState();
      
      // Create some data
      const data = [
        ['A1 value', 'B1 value'],
        ['A2 value', 'B2 value']
      ];
      
      // Import the data
      store.fromAOA(data);
      
      // Check cells were created
      expect(store.workbook.sheets[0].cells['A1'].value).toBe('A1 value');
      expect(store.workbook.sheets[0].cells['B1'].value).toBe('B1 value');
      expect(store.workbook.sheets[0].cells['A2'].value).toBe('A2 value');
      expect(store.workbook.sheets[0].cells['B2'].value).toBe('B2 value');
    });

    test('fromAOA skips empty cells', () => {
      const store = useStore.getState();
      
      // Create data with empty cells
      const data = [
        ['A1 value', ''],
        ['', 'B2 value']
      ];
      
      // Import the data
      store.fromAOA(data);
      
      // Check only non-empty cells were created
      expect(store.workbook.sheets[0].cells['A1'].value).toBe('A1 value');
      expect(store.workbook.sheets[0].cells['B1']).toBeUndefined();
      expect(store.workbook.sheets[0].cells['A2']).toBeUndefined();
      expect(store.workbook.sheets[0].cells['B2'].value).toBe('B2 value');
    });

    test('fromAOA replaces existing sheet data', () => {
      const store = useStore.getState();
      
      // Add some initial data
      store.setCellValue('A1', 'original');
      store.setCellValue('C3', 'original');
      
      // Create new data
      const data = [
        ['new A1', 'new B1']
      ];
      
      // Import the data
      store.fromAOA(data);
      
      // Check cells were replaced
      expect(store.workbook.sheets[0].cells['A1'].value).toBe('new A1');
      expect(store.workbook.sheets[0].cells['B1'].value).toBe('new B1');
      expect(store.workbook.sheets[0].cells['C3']).toBeUndefined();
    });
  });

  describe('persistence', () => {
    test('cell changes are saved to localStorage', () => {
      const store = useStore.getState();
      
      // Add a cell
      store.setCellValue('A1', 'test value');
      
      // Check localStorage was called
      expect(localStorageMock.setItem).toHaveBeenCalled();
      
      // Get the saved data
      const key = 'excel-clone/workbook/v1';
      const savedData = JSON.parse(localStorageMock.store[key]);
      
      // Check the data includes our cell
      expect(savedData.sheets[0].cells.A1.value).toBe('test value');
    });

    test('formatting changes are saved to localStorage', () => {
      const store = useStore.getState();
      
      // Add formatting
      store.toggleFormat('A1', 'bold');
      
      // Check localStorage was called
      expect(localStorageMock.setItem).toHaveBeenCalled();
      
      // Get the saved data
      const key = 'excel-clone/workbook/v1';
      const savedData = JSON.parse(localStorageMock.store[key]);
      
      // Check the data includes our formatting
      expect(savedData.sheets[0].cells.A1.format.bold).toBe(true);
    });

    test('sizing changes are saved to localStorage', () => {
      const store = useStore.getState();
      
      // Change sizes
      store.setColWidth(1, 200);
      store.setRowHeight(2, 50);
      
      // Check localStorage was called
      expect(localStorageMock.setItem).toHaveBeenCalled();
      
      // Get the saved data
      const key = 'excel-clone/workbook/v1';
      const savedData = JSON.parse(localStorageMock.store[key]);
      
      // Check the data includes our sizes
      expect(savedData.sheets[0].colWidths[1]).toBe(200);
      expect(savedData.sheets[0].rowHeights[2]).toBe(50);
    });

    test('sheet operations are saved to localStorage', () => {
      const store = useStore.getState();
      
      // Add a sheet
      store.addSheet();
      
      // Check localStorage was called
      expect(localStorageMock.setItem).toHaveBeenCalled();
      
      // Get the saved data
      const key = 'excel-clone/workbook/v1';
      const savedData = JSON.parse(localStorageMock.store[key]);
      
      // Check the data includes both sheets
      expect(savedData.sheets.length).toBe(2);
      expect(savedData.activeIndex).toBe(1);
    });

    test('store loads from localStorage on initialization', () => {
      // Set up mock data
      const mockSheet: Sheet = {
        id: 'test-id',
        name: 'Test Sheet',
        cells: { 'A1': { value: 'loaded value' } },
        colWidths: [undefined, 200],
        rowHeights: [undefined, 50]
      };
      
      const mockWorkbook = {
        sheets: [mockSheet],
        activeIndex: 0
      };
      
      // Save to localStorage
      const key = 'excel-clone/workbook/v1';
      localStorageMock.store[key] = JSON.stringify(mockWorkbook);
      
      // Reset the store module to force reinitialization
      vi.resetModules();
      const { useStore: freshStore } = require('../store');
      
      // Get the state
      const state = freshStore.getState();
      
      // Check data was loaded
      expect(state.workbook.sheets[0].id).toBe('test-id');
      expect(state.workbook.sheets[0].name).toBe('Test Sheet');
      expect(state.workbook.sheets[0].cells.A1.value).toBe('loaded value');
      expect(state.workbook.sheets[0].colWidths[1]).toBe(200);
      expect(state.workbook.sheets[0].rowHeights[1]).toBe(50);
    });
  });
});
