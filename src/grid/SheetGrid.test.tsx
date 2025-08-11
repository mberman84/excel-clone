// @vitest-environment jsdom
import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import SheetGrid from './SheetGrid';
import { useStore } from '../store';

// Mock the store to avoid dependency issues
// Expose a stable spy so the test can assert against it
const mockSortByColumn = vi.fn();

vi.mock('../store', () => ({
  useStore: vi.fn().mockImplementation((selector) => {
    const state = {
      workbook: {
        sheets: [{ id: 'sheet1', cells: {}, colWidths: {}, rowHeights: {} }],
        activeIndex: 0
      },
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' },
      selectCell: vi.fn(),
      startEdit: vi.fn(),
      setDraft: vi.fn(),
      commitEdit: vi.fn(),
      cancelEdit: vi.fn(),
      setColWidth: vi.fn(),
      setRowHeight: vi.fn(),
      getUsedRange: vi.fn().mockReturnValue({ maxRow: 10, maxCol: 10 }),
      setSelectionEnd: vi.fn(),
      selectRange: vi.fn(),
      sortByColumn: mockSortByColumn
    };
    return selector(state);
  }),
  // constants consumed directly from the module
  DEFAULT_ROW_HEIGHT: 28,
  DEFAULT_COL_WIDTH: 120
}));

// Augment mock with sizing constants expected by SheetGrid
vi.mocked(useStore); // ensure type inference

describe('SheetGrid', () => {
  test('opens column sort menu on header arrow click', async () => {
    // Render the component
    render(<SheetGrid />);
    
    // Find the dropdown trigger buttons
    const triggers = await screen.findAllByRole('button', { name: '▼' });
    expect(triggers.length).toBeGreaterThan(0);
    
    // Click the first trigger
    fireEvent.click(triggers[0]);
    
    // Verify the menu appears with both options
    expect(await screen.findByText('Sort A→Z')).toBeInTheDocument();
    expect(screen.getByText('Sort Z→A')).toBeInTheDocument();
    
    // Click the Sort A→Z option
    fireEvent.click(screen.getByText('Sort A→Z'));
    
    // Verify the menu closes
    await waitFor(() => {
      expect(screen.queryByText('Sort A→Z')).not.toBeInTheDocument();
    });
    
    // Verify sortByColumn was called with the right parameters
    expect(mockSortByColumn).toHaveBeenCalledWith(0, 'asc');
  });
});
