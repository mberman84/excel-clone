import React, { useCallback, useRef, useEffect, useState } from 'react'
import {
  VariableSizeGrid as Grid,
  GridChildComponentProps,
} from 'react-window'
import classNames from 'classnames'
import { useStore } from '../store'
import { columnIndexToLabel, makeCellAddress, parseCellAddress } from '../utils/cellAddresses'
import { evaluateDisplay, isFormula } from '../formula'

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

// data object fed into each virtualised cell so it always receives
// the latest state without relying on stale closures
type GridData = {
  sheet: ReturnType<typeof useStore>['sheet']
  selection: ReturnType<typeof useStore>['selection']
  editing: ReturnType<typeof useStore>['editing']
  selectCell: (r: number, c: number) => void
  startEdit: (addr: string) => void
  setDraft: (v: string) => void
  commitEdit: () => void
  cancelEdit: () => void
  setColWidth: (c: number, px: number) => void
  setRowHeight: (r: number, px: number) => void
}

export default function SheetGrid() {
  const { 
    sheet, selection, editing, selectCell, startEdit, setDraft, commitEdit, cancelEdit,
    setColWidth, setRowHeight, getUsedRange
  } = useStore(s => ({
    sheet: s.sheet,
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
  }))

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
      startEdit,
      setDraft,
      commitEdit,
      cancelEdit,
      setColWidth,
      setRowHeight,
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
        <div style={style} className={classes}>
          {columnIndexToLabel(columnIndex - 1)}
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
      )
    }
    
    if (isHeaderCol) {
      return (
        <div style={style} className={classes}>
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
    const isSelected = selection.row === rowIndex && selection.col === columnIndex
    const isEditingHere = editing.addr === addr
    const cell = sheet.cells[addr]
    const display = cell ? (isFormula(cell.value) ? evaluateDisplay(addr, sheet) : cell.value) : ''

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
      return (
        <div
          style={styleMerged}
          className={classNames(classes, {
            'cell--selected': isSelected,
            'cell--editing': isEditingHere,
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
        </div>
      )
    }

    return (
      <div
        style={styleMerged}
        className={classNames(classes, {
          'cell--selected': isSelected,
          'cell--editing': isEditingHere,
        })}
        onDoubleClick={(e) => {
          e.preventDefault()
          /* explicit double-click editing per requirements */
          selectCell(rowIndex, columnIndex)
          startEdit(addr)
        }}
        onMouseDown={(e) => {
          if (e.button !== 0) return
          /* always select on mouse down */
          selectCell(rowIndex, columnIndex)
          /* if this was the second click of a double-click, start editing immediately */
          if (e.detail === 2) {
            e.preventDefault()
            startEdit(addr)
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
        /* ensure react-window’s internal cache resets when switching sheets */
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
          startEdit,
          setDraft,
          commitEdit,
          cancelEdit,
          setColWidth,
          setRowHeight,
        } as GridData}
      >
        {Cell}
      </Grid>
    </div>
  )
}
