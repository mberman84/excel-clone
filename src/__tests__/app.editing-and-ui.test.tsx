import React from 'react';
import { describe, test, expect, beforeEach, afterEach, vi } from 'vitest';
import { render, screen, within, fireEvent, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import App from '../App';
import { useStore } from '../store';
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

// Mock ResizeObserver
class ResizeObserverMock {
  observe = vi.fn();
  unobserve = vi.fn();
  disconnect = vi.fn();
}

window.ResizeObserver = ResizeObserverMock;

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

// Helper to seed the store with test data
const seedStore = (cells: Record<string, string>) => {
  const store = useStore.getState();
  Object.entries(cells).forEach(([addr, value]) => {
    store.setCellValue(addr, value);
  });
};

// Helper to find a cell by coordinates
const findCellByCoords = (container: HTMLElement, row: number, col: number) => {
  // Skip header row and column
  const rowIndex = row + 1;
  const colIndex = col + 1;
  
  // Find all cells
  const cells = container.querySelectorAll('.cell:not(.cell--header)');
  
  // Calculate the index in the grid (accounting for header row and column)
  const index = (rowIndex - 1) * 100 + (colIndex - 1); // 100 is COLS from SheetGrid
  
  return cells[index] as HTMLElement;
};

// Helper to find a cell by its displayed text content
const findCellByText = (container: HTMLElement, text: string) => {
  const cells = container.querySelectorAll('.cell:not(.cell--header)');
  return Array.from(cells).find(cell => cell.textContent === text) as HTMLElement;
};

// Helper to find a header cell by its label
const findHeaderByLabel = (container: HTMLElement, label: string) => {
  const headers = container.querySelectorAll('.cell--header');
  return Array.from(headers).find(header => header.textContent === label) as HTMLElement;
};

// Helper to find the version badge
const findVersionBadge = () => {
  return screen.getByText(/v2025\.08\.\d+-\d+/);
};

describe('Excel Clone App - UI and Editing Tests', () => {
  beforeEach(() => {
    resetStore();
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('Cell editing', () => {
    test('double-click shows input with focus and blinking caret immediately', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Seed a cell with some data
      seedStore({ 'A1': 'Test Value' });
      
      // Find the cell
      const cell = findCellByText(container, 'Test Value');
      expect(cell).toBeTruthy();
      
      // Double-click the cell
      await user.dblClick(cell!);
      
      // Check that the cell is in editing mode
      const editingCell = container.querySelector('.cell--editing');
      expect(editingCell).toBeTruthy();
      
      // Check that the input is focused
      const input = editingCell!.querySelector('input');
      expect(input).toBeTruthy();
      expect(document.activeElement).toBe(input);
      
      // Check that the input has the cell value
      expect(input!.value).toBe('Test Value');
    });

    test('editing shows visual feedback without requiring typing', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Seed a cell with some data
      seedStore({ 'B2': 'Another Value' });
      
      // Find the cell
      const cell = findCellByCoords(container, 1, 1); // B2
      
      // Double-click the cell
      await user.dblClick(cell);
      
      // Check that the cell has the editing class immediately
      expect(cell.classList.contains('cell--editing')).toBe(true);
      
      // Check that the overlay pseudo-element is applied (can't directly test, but can check class)
      const computedStyle = window.getComputedStyle(cell);
      expect(cell.classList.contains('cell--editing')).toBe(true);
      
      // Check that the input is visible and focused
      const input = cell.querySelector('input');
      expect(input).toBeTruthy();
      expect(document.activeElement).toBe(input);
    });
    
    test('formula editing shows ghost overlay with highlighted references', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Seed cells with data
      seedStore({
        'A1': '10',
        'A2': '20',
        'B1': '30'
      });
      
      // Find C1 cell to enter a formula
      const cell = findCellByCoords(container, 0, 2); // C1
      
      // Double-click and enter a formula
      await user.dblClick(cell);
      await user.keyboard('=A1+B1');
      
      // Check that the formula ghost overlay is displayed
      const ghostOverlay = container.querySelector('.formula-ghost');
      expect(ghostOverlay).toBeTruthy();
      
      // Check that the references are highlighted in the ghost
      const refs = ghostOverlay!.querySelectorAll('.formula-ref');
      expect(refs.length).toBe(2); // A1 and B1
      
      // Check that the referenced cells have highlight classes
      const a1Cell = findCellByCoords(container, 0, 0);
      const b1Cell = findCellByCoords(container, 0, 1);
      
      expect(a1Cell.className).toMatch(/cell--ref-\d/);
      expect(b1Cell.className).toMatch(/cell--ref-\d/);
    });
    
    test('commit formula with Enter evaluates and displays result', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Seed cells with data
      seedStore({
        'A1': '10',
        'A2': '20'
      });
      
      // Find B1 cell to enter a formula
      const cell = findCellByCoords(container, 0, 1); // B1
      
      // Double-click and enter a formula
      await user.dblClick(cell);
      await user.keyboard('=A1+A2');
      await user.keyboard('{Enter}');
      
      // Check that the formula was evaluated
      expect(cell.textContent).toBe('30');
    });
  });

  describe('Selection', () => {
    test('drag selection updates selection range', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find cells for drag selection
      const startCell = findCellByCoords(container, 1, 1); // B2
      const endCell = findCellByCoords(container, 3, 3); // D4
      
      // Start selection at B2
      await user.click(startCell);
      
      // Get initial selection state
      const initialState = useStore.getState().selection;
      expect(initialState.row).toBe(2);
      expect(initialState.col).toBe(2);
      expect(initialState.endRow).toBe(2);
      expect(initialState.endCol).toBe(2);
      
      // Simulate drag to D4
      fireEvent.mouseDown(startCell);
      fireEvent.mouseEnter(endCell, { buttons: 1 }); // buttons: 1 simulates left mouse button held down
      fireEvent.mouseUp(endCell);
      
      // Check updated selection state
      const updatedState = useStore.getState().selection;
      expect(updatedState.row).toBe(2);
      expect(updatedState.col).toBe(2);
      expect(updatedState.endRow).toBe(4);
      expect(updatedState.endCol).toBe(4);
      
      // Check that cells in the range have the in-range class
      const cellC3 = findCellByCoords(container, 2, 2); // C3
      expect(cellC3.classList.contains('cell--in-range')).toBe(true);
    });
    
    test('clicking column header selects entire column', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find column B header
      const colHeader = findHeaderByLabel(container, 'B');
      expect(colHeader).toBeTruthy();
      
      // Click the column header
      await user.click(colHeader!);
      
      // Check selection state
      const selection = useStore.getState().selection;
      expect(selection.row).toBe(1);
      expect(selection.col).toBe(2); // Column B
      expect(selection.endRow).toBe(1000); // ROWS constant from SheetGrid
      expect(selection.endCol).toBe(2);
      
      // Check that cells in column B have the in-range class
      const cellB2 = findCellByCoords(container, 1, 1);
      const cellB10 = findCellByCoords(container, 9, 1);
      
      expect(cellB2.classList.contains('cell--in-range') || 
             cellB2.classList.contains('cell--selected')).toBe(true);
      expect(cellB10.classList.contains('cell--in-range') || 
             cellB10.classList.contains('cell--selected')).toBe(true);
    });
    
    test('clicking row header selects entire row', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find row 3 header
      const rowHeader = findHeaderByLabel(container, '3');
      expect(rowHeader).toBeTruthy();
      
      // Click the row header
      await user.click(rowHeader!);
      
      // Check selection state
      const selection = useStore.getState().selection;
      expect(selection.row).toBe(3); // Row 3
      expect(selection.col).toBe(1);
      expect(selection.endRow).toBe(3);
      expect(selection.endCol).toBe(100); // COLS constant from SheetGrid
      
      // Check that cells in row 3 have the in-range class
      const cellA3 = findCellByCoords(container, 2, 0);
      const cellE3 = findCellByCoords(container, 2, 4);
      
      expect(cellA3.classList.contains('cell--in-range') || 
             cellA3.classList.contains('cell--selected')).toBe(true);
      expect(cellE3.classList.contains('cell--in-range') || 
             cellE3.classList.contains('cell--selected')).toBe(true);
    });
  });

  describe('Column and Row Resizing', () => {
    test('column resizing via drag handle updates column width', async () => {
      const { container } = render(<App />);
      
      // Find column B header
      const colHeader = findHeaderByLabel(container, 'B');
      expect(colHeader).toBeTruthy();
      
      // Find the resize handle
      const resizeHandle = colHeader!.querySelector('.col-resize-handle');
      expect(resizeHandle).toBeTruthy();
      
      // Get initial column width
      const initialWidth = useStore.getState().workbook.sheets[0].colWidths[1] || 120; // Default width
      
      // Simulate drag resize
      fireEvent.mouseDown(resizeHandle!, { clientX: 0 });
      fireEvent.mouseMove(document, { clientX: 50 }); // Move 50px to the right
      fireEvent.mouseUp(document);
      
      // Check that column width was updated
      const newWidth = useStore.getState().workbook.sheets[0].colWidths[1];
      expect(newWidth).toBeGreaterThan(initialWidth);
    });
    
    test('double-clicking column resize handle auto-fits content', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Seed a cell with wide content
      seedStore({ 'B2': 'This is a wide content cell to test auto-fit' });
      
      // Find column B header
      const colHeader = findHeaderByLabel(container, 'B');
      
      // Find the resize handle
      const resizeHandle = colHeader!.querySelector('.col-resize-handle');
      
      // Double-click the resize handle
      await user.dblClick(resizeHandle!);
      
      // Check that column width was updated to fit content
      const newWidth = useStore.getState().workbook.sheets[0].colWidths[1];
      expect(newWidth).toBeGreaterThan(120); // Default width
    });
  });

  describe('Sheet Tabs', () => {
    test('adding a new tab creates a new sheet', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find the add tab button
      const addTabButton = container.querySelector('.tab.add-tab');
      expect(addTabButton).toBeTruthy();
      
      // Click to add a new tab
      await user.click(addTabButton!);
      
      // Check that a new sheet was added
      const sheets = useStore.getState().workbook.sheets;
      expect(sheets.length).toBe(2);
      expect(sheets[1].name).toBe('Sheet 2');
      
      // Check that the new tab is displayed and active
      const tabs = container.querySelectorAll('.tab:not(.add-tab)');
      expect(tabs.length).toBe(2);
      expect(tabs[1].classList.contains('active')).toBe(true);
    });
    
    test('renaming a tab updates sheet name', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find the first tab
      const tab = container.querySelector('.tab:not(.add-tab)');
      expect(tab).toBeTruthy();
      
      // Double-click to start renaming
      await user.dblClick(tab!);
      
      // Find the rename input
      const renameInput = tab!.querySelector('input');
      expect(renameInput).toBeTruthy();
      
      // Clear and type new name
      await user.clear(renameInput!);
      await user.type(renameInput!, 'Renamed Sheet');
      await user.keyboard('{Enter}');
      
      // Check that sheet name was updated
      expect(useStore.getState().workbook.sheets[0].name).toBe('Renamed Sheet');
      
      // Check that tab displays new name
      expect(tab!.textContent).toContain('Renamed Sheet');
    });
    
    test('data is independent between tabs', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Add data to first sheet
      seedStore({ 'A1': 'Sheet 1 Data' });
      
      // Add a new sheet
      const addTabButton = container.querySelector('.tab.add-tab');
      await user.click(addTabButton!);
      
      // Check that A1 is empty in the second sheet
      const cellA1 = findCellByCoords(container, 0, 0);
      expect(cellA1.textContent).toBe('');
      
      // Add different data to second sheet
      await user.dblClick(cellA1);
      await user.keyboard('Sheet 2 Data');
      await user.keyboard('{Enter}');
      
      // Switch back to first sheet
      const firstTab = container.querySelectorAll('.tab:not(.add-tab)')[0];
      await user.click(firstTab);
      
      // Check that A1 has the first sheet data
      const cellA1FirstSheet = findCellByCoords(container, 0, 0);
      expect(cellA1FirstSheet.textContent).toBe('Sheet 1 Data');
      
      // Switch to second sheet again
      const secondTab = container.querySelectorAll('.tab:not(.add-tab)')[1];
      await user.click(secondTab);
      
      // Check that A1 has the second sheet data
      const cellA1SecondSheet = findCellByCoords(container, 0, 0);
      expect(cellA1SecondSheet.textContent).toBe('Sheet 2 Data');
    });
  });

  describe('Undo/Redo', () => {
    test('undo button reverts cell changes', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find A1 cell and enter data
      const cellA1 = findCellByCoords(container, 0, 0);
      await user.dblClick(cellA1);
      await user.keyboard('Test Undo');
      await user.keyboard('{Enter}');
      
      // Verify data was entered
      expect(cellA1.textContent).toBe('Test Undo');
      
      // Find and click undo button
      const undoButton = screen.getByText('Undo');
      await user.click(undoButton);
      
      // Check that cell content was reverted
      expect(cellA1.textContent).toBe('');
    });
    
    test('redo button reapplies undone changes', async () => {
      const user = userEvent.setup();
      const { container } = render(<App />);
      
      // Find A1 cell and enter data
      const cellA1 = findCellByCoords(container, 0, 0);
      await user.dblClick(cellA1);
      await user.keyboard('Test Redo');
      await user.keyboard('{Enter}');
      
      // Verify data was entered
      expect(cellA1.textContent).toBe('Test Redo');
      
      // Undo the change
      const undoButton = screen.getByText('Undo');
      await user.click(undoButton);
      
      // Verify undo worked
      expect(cellA1.textContent).toBe('');
      
      // Redo the change
      const redoButton = screen.getByText('Redo');
      await user.click(redoButton);
      
      // Check that cell content was restored
      expect(cellA1.textContent).toBe('Test Redo');
    });
  });

  describe('Dark Mode', () => {
    test('dark mode toggle adds theme-dark class and persists in localStorage', async () => {
      const user = userEvent.setup();
      render(<App />);
      
      // Initially in light mode
      expect(document.documentElement.classList.contains('theme-dark')).toBe(false);
      
      // Find and click dark mode toggle
      const darkModeButton = screen.getByText('Dark mode');
      await user.click(darkModeButton);
      
      // Check that theme-dark class was added
      expect(document.documentElement.classList.contains('theme-dark')).toBe(true);
      
      // Check that preference was saved to localStorage
      expect(localStorageMock.setItem).toHaveBeenCalledWith('excel-clone/theme', 'dark');
      
      // Toggle back to light mode
      const lightModeButton = screen.getByText('Light mode');
      await user.click(lightModeButton);
      
      // Check that theme-dark class was removed
      expect(document.documentElement.classList.contains('theme-dark')).toBe(false);
      
      // Check that preference was updated in localStorage
      expect(localStorageMock.setItem).toHaveBeenCalledWith('excel-clone/theme', 'light');
    });
    
    test('dark mode preference is loaded from localStorage', async () => {
      // Set dark mode preference in localStorage
      localStorageMock.store['excel-clone/theme'] = 'dark';
      
      // Render app
      render(<App />);
      
      // Check that theme-dark class was applied
      expect(document.documentElement.classList.contains('theme-dark')).toBe(true);
      
      // Check that button shows Light mode text
      expect(screen.getByText('Light mode')).toBeTruthy();
    });
  });

  describe('Version Badge', () => {
    test('version badge shows correct format', () => {
      render(<App />);
      
      // Find version badge
      const versionBadge = findVersionBadge();
      expect(versionBadge).toBeTruthy();
      
      // Check that version follows expected format
      expect(versionBadge.textContent).toMatch(/v2025\.08\.\d+-\d+/);
    });
  });
});
