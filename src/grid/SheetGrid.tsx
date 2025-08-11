import React, { useCallback, useRef, useEffect, useState, useMemo } from 'react'
import {
  VariableSizeGrid as Grid,
  GridChildComponentProps,
} from 'react-window'
import classNames from 'classnames'
import { useStore } from '../store'
import { columnIndexToLabel, makeCellAddress, parseCellAddress, expandRange } from '../utils/cellAddresses'
import { evaluateDisplay, isFormula } from '../formula'
import { Sheet } from '../types'

const ROWS = 1000
const COLS = 100
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
} from '../store'

/**
 * Fallback sizes (Excel-ish defaults)
 *   • rows  18-20 px  (we keep 28 from earlier design)
 *   • cols ~64-70 px  (we keep 120 from earlier)
 */
const BASE_ROW_HEIGHT = DEFAULT_ROW_HEIGHT
const BASE_COL_WIDTH = DEFAULT_COL_WIDTH

// Canvas for text measurement
let measureCanvas: HTMLCanvasElement | null = null;
function getMeasureCanvas() {
  if (!measureCanvas) {
    measureCanvas = document.createElement('canvas');
  }
  return measureCanvas;
}

function measureTextWidth(text: string): number {
  const canvas = getMeasureCanvas();
  const ctx = canvas.getContext('2d');
  if (!ctx) return 0;
  ctx.font = '14px system-ui, -apple-system, sans-serif';
  return ctx.measureText(text).width + 18; // Add padding
}

/**
 * Parse cell references from a formula string and generate highlighting
 */
function parseFormulaReferences(draft: string): { 
  refMap: Map<string, number>, 
  ghostHTML: string 
} {
  // Basic HTML escaper to prevent injection / malformed markup
  const esc = (s: string) =>
    s.replace(/[&<>"']/g, c => (
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' } as Record<string, string>)[c]
    ));

  // Default empty result
  const refMap = new Map<string, number>();
  
  // If not a formula, return empty result
  if (!draft.startsWith('=')) {
    return { refMap, ghostHTML: esc(draft) };
  }
  
  // Find all cell references (A1, $A$1) and ranges (A1:B3)
  const refRegex = /(\$?[A-Za-z]+\$?[0-9]+(?::\$?[A-Za-z]+\$?[0-9]+)?)/g;
  const matches = [...draft.matchAll(refRegex)];
  
  // Process each match
  let lastIndex = 0;
  let colorIndex = 1;
  let parts: string[] = [];
  
  for (const match of matches) {
    const ref = match[0];
    const startIndex = match.index!;
    
    // Add text before this reference
    if (startIndex > lastIndex) {
      parts.push(esc(draft.substring(lastIndex, startIndex)));
    }
    
    // Check if it's a range (contains :)
    if (ref.includes(':')) {
      try {
        // Expand the range and add each cell to the map
        const [start, end] = ref.split(':');
        const addresses = expandRange(start, end);
        
        // Assign the same color to all cells in the range
        const refClass = colorIndex;
        for (const addr of addresses) {
          refMap.set(addr, refClass);
        }
        
        // Increment color index (1-4, then cycle)
        colorIndex = colorIndex % 4 + 1;
        
        // Add the highlighted range to the HTML
        parts.push(`<span class="formula-ref ref-${refClass}">${esc(ref)}</span>`);
      } catch (e) {
        // If range expansion fails, just add the text
        parts.push(esc(ref));
      }
    } else {
      try {
        // Parse to verify it's a valid reference
        parseCellAddress(ref);
        
        // Assign color and add to map
        const refClass = colorIndex;
        refMap.set(ref, refClass);
        
        // Increment color index (1-4, then cycle)
        colorIndex = colorIndex % 4 + 1;
        
        // Add the highlighted cell reference to the HTML
        parts.push(`<span class="formula-ref ref-${refClass}">${esc(ref)}</span>`);
      } catch (e) {
        // If parsing fails, just add the text
        parts.push(esc(ref));
      }
    }
    
    lastIndex = startIndex + ref.length;
  }
  
  // Add any remaining text
  if (lastIndex < draft.length) {
    parts.push(esc(draft.substring(lastIndex)));
  }
  
  return {
    refMap,
    ghostHTML: parts.join('')
  };
}

// Type for tracking drag state
type DragState = {
  type: 'cells' | 'row' | 'col';
  anchorRow: number;
  anchorCol: number;
} | null;

// data object fed into each virtualised cell so it always receives
// the latest state without relying on stale closures
type GridData = {
  sheet: Sheet
  selection: ReturnType<typeof useStore>['selection']
  editing: ReturnType<typeof useStore>['editing']
  selectCell: (r: number, c: number) => void
  setSelectionEnd: (r: number, c: number) => void
  selectRange: (sr: number, sc: number, er: number, ec: number) => void
  startEdit: (addr: string) => void
  setDraft: (v: string) => void
  commitEdit: () => void
  cancelEdit: () => void
  setColWidth: (c: number, px: number) => void
  setRowHeight: (r: number, px: number) => void
  refMap: Map<string, number>
  ghostHTML: string
  dragState: DragState
  setDragState: (state: DragState) => void
  sortByColumn: (col: number, direction: 'asc' | 'desc') => void
}

// Dedicated component for column headers
const HeaderCell = ({ 
  columnIndex, 
  style, 
  classes, 
  selection, 
  selectRange, 
  setDragState, 
  dragState, 
  columnIndexToLabel, 
  startColumnResize, 
  autoFitColumnWidth,
  sortByColumn
}: { 
  columnIndex: number
  style: React.CSSProperties
  classes: string
  selection: ReturnType<typeof useStore>['selection']
  selectRange: (sr: number, sc: number, er: number, ec: number) => void
  setDragState: (state: DragState) => void
  dragState: DragState
  columnIndexToLabel: (index: number) => string
  startColumnResize: (e: React.MouseEvent) => void
  autoFitColumnWidth: (columnIndex: number) => void
  sortByColumn: (col: number, direction: 'asc' | 'desc') => void
}) => {
  const [menuOpen, setMenuOpen] = useState(false);
  const headerRef = useRef<HTMLDivElement>(null);
  
  // Close menu when clicking outside
  useEffect(() => {
    if (!menuOpen) return;
    
    const handleClickOutside = (e: MouseEvent) => {
      if (headerRef.current && !headerRef.current.contains(e.target as Node)) {
        setMenuOpen(false);
      }
    };
    
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [menuOpen]);
  
  return (
    <div 
      ref={headerRef}
      style={style} 
      className={classes}
      onMouseDown={(e) => {
        // Ignore if clicking on resize handle or menu
        if ((e.target as HTMLElement).closest('.col-resize-handle') || 
            (e.target as HTMLElement).closest('.col-menu-trigger') ||
            (e.target as HTMLElement).closest('.col-menu')) return;
        if (e.button !== 0) return; // Left click only
        
        if (e.shiftKey) {
          // Extend selection from current position
          selectRange(1, selection.col, ROWS, columnIndex);
          setDragState({
            type: 'col',
            anchorRow: 1,
            anchorCol: selection.col
          });
        } else {
          // Select entire column
          selectRange(1, columnIndex, ROWS, columnIndex);
          setDragState({
            type: 'col',
            anchorRow: 1,
            anchorCol: columnIndex
          });
        }
      }}
      onMouseEnter={(e) => {
        // Update selection when dragging
        if (dragState?.type === 'col' && (e.buttons & 1) !== 0) {
          selectRange(
            1, 
            Math.min(dragState.anchorCol, columnIndex),
            ROWS,
            Math.max(dragState.anchorCol, columnIndex)
          );
        }
      }}
    >
      {columnIndexToLabel(columnIndex - 1)}
      
      {/* Sort dropdown trigger */}
      <button 
        className="col-menu-trigger"
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setMenuOpen(!menuOpen);
        }}
      >
        ▼
      </button>
      
      {/* Sort dropdown menu */}
      {menuOpen && (
        <div className="col-menu">
          <div 
            className="col-menu-item"
            onClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              sortByColumn(columnIndex - 1, 'asc');
              setMenuOpen(false);
            }}
          >
            Sort A→Z
          </div>
          <div 
            className="col-menu-item"
            onClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              sortByColumn(columnIndex - 1, 'desc');
              setMenuOpen(false);
            }}
          >
            Sort Z→A
          </div>
        </div>
      )}
      
      <div 
        className="col-resize-handle"
        onMouseDown={startColumnResize}
        onDoubleClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          autoFitColumnWidth(columnIndex);
        }}
      />
    </div>
  );
};

export default function SheetGrid() {
  const { 
    workbook, selection, editing, selectCell, startEdit, setDraft, commitEdit, cancelEdit,
    setColWidth, setRowHeight, getUsedRange, setSelectionEnd, selectRange, sortByColumn
  } = useStore(s => ({
    workbook: s.workbook,
    selection: s.selection,
    editing: s.editing,
    selectCell: s.selectCell,
    startEdit: s.startEdit,
    setDraft: s.setDraft,
    commitEdit: s.commitEdit,
    cancelEdit: s.cancelEdit,
    setColWidth: s.setColWidth,
    setRowHeight: s.setRowHeight,
    getUsedRange: s.getUsedRange,
    setSelectionEnd: s.setSelectionEnd,
    selectRange: s.selectRange,
    sortByColumn: s.sortByColumn,
  }))

  // Derive the active sheet reactively from the workbook
  const sheet = useMemo(() => workbook.sheets[workbook.activeIndex], [workbook])

  // Parse formula references when editing
  const { refMap, ghostHTML } = useMemo(() => {
    return parseFormulaReferences(editing.draft);
  }, [editing.draft]);

  // Track drag state
  const [dragState, setDragState] = useState<DragState>(null);

  // Clear drag state on document mouseup
  useEffect(() => {
    const handleMouseUp = () => {
      setDragState(null);
    };
    
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  const gridRef = useRef<any>(null)

  const onKeyDown = useCallback((e: React.KeyboardEvent<HTMLDivElement>) => {
    const { row, col } = selection
    const isEditing = !!editing.addr
    if (isEditing) return // input handles its own keys
    if (row <= 0 || col <= 0) return
    let nr = row, nc = col
    if (e.key === 'ArrowUp') { nr = Math.max(1, row - 1) }
    else if (e.key === 'ArrowDown') { nr = Math.min(ROWS, row + 1) }
    else if (e.key === 'ArrowLeft') { nc = Math.max(1, col - 1) }
    else if (e.key === 'ArrowRight') { nc = Math.min(COLS, col + 1) }
    else if (e.key === 'Tab') { e.preventDefault(); nc = Math.min(COLS, col + (e.shiftKey ? -1 : 1)); if (nc < 1) { nc = COLS; nr = Math.max(1, row - 1) } else if (nc > COLS) { nc = 1; nr = Math.min(ROWS, row + 1) } }
    else if (e.key === 'Enter') { e.preventDefault(); const addr = makeCellAddress(col - 1, row - 1); startEdit(addr); return }
    else if (e.key === 'F2') { const addr = makeCellAddress(col - 1, row - 1); startEdit(addr); return }
    else { return }
    selectCell(nr, nc)
    gridRef.current?.scrollToItem({ rowIndex: nr, columnIndex: nc, align: 'smart' } as any)
  }, [selection, editing])

  // Auto-fit column width based on content
  const autoFitColumnWidth = useCallback((columnIndex: number) => {
    if (columnIndex <= 0) return;
    
    const colIdx = columnIndex - 1;
    const { maxRow } = getUsedRange();
    const rowsToCheck = Math.min(200, maxRow);
    let maxWidth = 40; // Minimum width

    for (let r = 0; r < rowsToCheck; r++) {
      const addr = makeCellAddress(colIdx, r);
      const cell = sheet.cells[addr];
      if (!cell) continue;
      
      const display = isFormula(cell.value) 
        ? evaluateDisplay(addr, sheet) 
        : cell.value;
      
      if (display) {
        const width = measureTextWidth(String(display));
        maxWidth = Math.max(maxWidth, Math.min(600, width)); // Clamp to reasonable max
      }
    }

    setColWidth(colIdx, maxWidth);
    if (gridRef.current) {
      gridRef.current.resetAfterColumnIndex(columnIndex, true);
    }
  }, [sheet, setColWidth, getUsedRange]);

  // Auto-fit row height based on content
  const autoFitRowHeight = useCallback((rowIndex: number) => {
    if (rowIndex <= 0) return;
    
    const rowIdx = rowIndex - 1;
    let maxLines = 1;

    for (let c = 0; c < COLS; c++) {
      const addr = makeCellAddress(c, rowIdx);
      const cell = sheet.cells[addr];
      if (!cell) continue;
      
      const display = isFormula(cell.value) 
        ? evaluateDisplay(addr, sheet) 
        : cell.value;
      
      if (display) {
        const lines = String(display).split('\n').length;
        maxLines = Math.max(maxLines, lines);
      }
    }

    setRowHeight(rowIdx, maxLines * BASE_ROW_HEIGHT);
    if (gridRef.current) {
      gridRef.current.resetAfterRowIndex(rowIndex, true);
    }
  }, [sheet, setRowHeight]);

  // item renderer receives fresh data via props.data
  const Cell = ({ columnIndex, rowIndex, style, data }: GridChildComponentProps<GridData>) => {
    const {
      sheet,
      selection,
      editing,
      selectCell,
      setSelectionEnd,
      selectRange,
      startEdit,
      setDraft,
      commitEdit,
      cancelEdit,
      setColWidth,
      setRowHeight,
      refMap,
      ghostHTML,
      dragState,
      setDragState,
      sortByColumn
    } = data

    const isHeaderRow = rowIndex === 0
    const isHeaderCol = columnIndex === 0
    const classes = classNames('cell', {
      'cell--header': isHeaderRow || isHeaderCol,
    })

    // Column resize functionality
    const startColumnResize = useCallback((e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      const startX = e.clientX;
      const colIdx = columnIndex - 1;
      const startWidth = sheet.colWidths?.[colIdx] ?? BASE_COL_WIDTH;
      
      const onMouseMove = (moveEvent: MouseEvent) => {
        const delta = moveEvent.clientX - startX;
        const newWidth = Math.max(40, startWidth + delta);
        setColWidth(colIdx, newWidth);
        if (gridRef.current) {
          gridRef.current.resetAfterColumnIndex(columnIndex, true);
        }
      };
      
      const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
      };
      
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    }, [columnIndex, sheet.colWidths]);

    // Row resize functionality
    const startRowResize = useCallback((e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      
      const startY = e.clientY;
      const rowIdx = rowIndex - 1;
      const startHeight = sheet.rowHeights?.[rowIdx] ?? BASE_ROW_HEIGHT;
      
      const onMouseMove = (moveEvent: MouseEvent) => {
        const delta = moveEvent.clientY - startY;
        const newHeight = Math.max(18, startHeight + delta);
        setRowHeight(rowIdx, newHeight);
        if (gridRef.current) {
          gridRef.current.resetAfterRowIndex(rowIndex, true);
        }
      };
      
      const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
      };
      
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    }, [rowIndex, sheet.rowHeights]);

    if (isHeaderRow && isHeaderCol) {
      return <div style={style} className={classes}></div>
    }
    
    if (isHeaderRow) {
      return (
        <HeaderCell 
          columnIndex={columnIndex}
          style={style}
          classes={classes}
          selection={selection}
          selectRange={selectRange}
          setDragState={setDragState}
          dragState={dragState}
          columnIndexToLabel={columnIndexToLabel}
          startColumnResize={startColumnResize}
          autoFitColumnWidth={autoFitColumnWidth}
          sortByColumn={sortByColumn}
        />
      );
    }
    
    if (isHeaderCol) {
      return (
        <div 
          style={style} 
          className={classes}
          onMouseDown={(e) => {
            // Ignore if clicking on resize handle
            if ((e.target as HTMLElement).closest('.row-resize-handle')) return;
            if (e.button !== 0) return; // Left click only
            
            if (e.shiftKey) {
              // Extend selection from current position
              selectRange(selection.row, 1, rowIndex, COLS);
              setDragState({
                type: 'row',
                anchorRow: selection.row,
                anchorCol: 1
              });
            } else {
              // Select entire row
              selectRange(rowIndex, 1, rowIndex, COLS);
              setDragState({
                type: 'row',
                anchorRow: rowIndex,
                anchorCol: 1
              });
            }
          }}
          onMouseEnter={(e) => {
            // Update selection when dragging
            if (dragState?.type === 'row' && (e.buttons & 1) !== 0) {
              selectRange(
                Math.min(dragState.anchorRow, rowIndex),
                1,
                Math.max(dragState.anchorRow, rowIndex),
                COLS
              );
            }
          }}
        >
          {rowIndex}
          <div 
            className="row-resize-handle"
            onMouseDown={startRowResize}
            onDoubleClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              autoFitRowHeight(rowIndex);
            }}
          />
        </div>
      )
    }
    
    const addr = makeCellAddress(columnIndex - 1, rowIndex - 1)
    const isAnchorCell = selection.row === rowIndex && selection.col === columnIndex
    const isInSelectionRange = (
      rowIndex >= Math.min(selection.row, selection.endRow) &&
      rowIndex <= Math.max(selection.row, selection.endRow) &&
      columnIndex >= Math.min(selection.col, selection.endCol) &&
      columnIndex <= Math.max(selection.col, selection.endCol)
    )
    const isEditingHere = editing.addr === addr
    const cell = sheet.cells[addr]
    const display = cell ? (isFormula(cell.value) ? evaluateDisplay(addr, sheet) : cell.value) : ''

    // Check if this cell is referenced in the current formula
    const refClass = editing.addr && refMap.has(addr) && addr !== editing.addr 
      ? `cell--ref-${refMap.get(addr)}` 
      : '';

    const fmt = cell?.format
    const styleMerged: React.CSSProperties = { 
      ...style, 
      fontWeight: fmt?.bold ? '700' : undefined, 
      fontStyle: fmt?.italic ? 'italic' : undefined, 
      textDecoration: fmt?.underline ? 'underline' : undefined, 
      color: fmt?.textColor, 
      backgroundColor: fmt?.fillColor 
    }

    // keep a ref so we can force-focus when entering editing mode
    const inputRef = useRef<HTMLInputElement>(null)

    // auto-focus & place caret at end whenever editing begins
    useEffect(() => {
      if (isEditingHere && inputRef.current) {
        const el = inputRef.current
        el.focus()
        const len = el.value.length
        // place caret at end
        try { el.setSelectionRange(len, len) } catch {}
      }
    }, [isEditingHere])

    if (isEditingHere) {
      const isFormulaEditing = editing.draft.startsWith('=')
      return (
        <div
          style={styleMerged}
          className={classNames(classes, {
            'cell--selected': isAnchorCell,
            'cell--in-range': isInSelectionRange && !isAnchorCell,
            'cell--editing': isEditingHere,
            'cell--formula-editing': isFormulaEditing,
          })}
        >
          <input
            className="cell-input"
            autoFocus
            ref={inputRef}
            value={editing.draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={() => commitEdit()}
            onKeyDown={(e) => {
              if (e.key === 'Enter') { e.preventDefault(); commitEdit() }
              else if (e.key === 'Escape') { e.preventDefault(); cancelEdit() }
            }}
          />
          {isFormulaEditing && (
            <div
              className="formula-ghost"
              dangerouslySetInnerHTML={{ __html: ghostHTML }}
            />
          )}
        </div>
      )
    }

    return (
      <div
        style={styleMerged}
        className={classNames(classes, {
          'cell--selected': isAnchorCell,
          'cell--in-range': isInSelectionRange && !isAnchorCell,
          'cell--editing': isEditingHere,
          [refClass]: !!refClass,
        })}
        onMouseDown={(e) => {
          if (e.button !== 0) return // Left click only
          /* If this is the second click of a double-click, start editing immediately */
          if (e.detail === 2) {
            e.preventDefault();
            e.stopPropagation();
            selectCell(rowIndex, columnIndex);
            startEdit(addr);
            return;
          }
          
          if (e.shiftKey) {
            // Extend selection from current position
            selectRange(selection.row, selection.col, rowIndex, columnIndex);
            setDragState({
              type: 'cells',
              anchorRow: selection.row,
              anchorCol: selection.col
            });
          } else {
            // Start new selection
            selectCell(rowIndex, columnIndex);
            setDragState({
              type: 'cells',
              anchorRow: rowIndex,
              anchorCol: columnIndex
            });
          }
        }}
        /* separate explicit double-click handler (outside of onMouseDown) */
        onDoubleClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          selectCell(rowIndex, columnIndex);
          startEdit(addr);
        }}
        /* if this was the second click of a double-click, start editing immediately */
        /* (handled above) */
        onMouseEnter={(e) => {
          // Update selection when dragging cells
          if (dragState?.type === 'cells' && (e.buttons & 1) !== 0) {
            setSelectionEnd(rowIndex, columnIndex);
          }
        }}
      >
        {display as any}
      </div>
    )
  }

  return (
    <div className="sheet" tabIndex={0} onKeyDown={onKeyDown}>
      {/*
        Compose the live item data object on every render so each virtualised
        cell re-evaluates with the latest sheet/selection/editing values.
      */}
      <Grid
        ref={gridRef}
        /* ensure react-window's internal cache resets when switching sheets */
        key={sheet.id}
        columnCount={COLS + 1}
        columnWidth={(index: number) =>
          index === 0
            ? 48 /* row-header column */
            : sheet.colWidths?.[index - 1] ?? BASE_COL_WIDTH
        }
        height={window.innerHeight - 130}
        rowCount={ROWS + 1}
        rowHeight={(index: number) =>
          index === 0
            ? 28 /* header row */
            : sheet.rowHeights?.[index - 1] ?? BASE_ROW_HEIGHT
        }
        width={window.innerWidth - 20}
        itemData={{
          sheet,
          selection,
          editing,
          selectCell,
          setSelectionEnd,
          selectRange,
          startEdit,
          setDraft,
          commitEdit,
          cancelEdit,
          setColWidth,
          setRowHeight,
          refMap,
          ghostHTML,
          dragState,
          setDragState,
          sortByColumn,
        } as GridData}
      >
        {Cell}
      </Grid>
    </div>
  )
}
