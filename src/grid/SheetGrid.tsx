import React, { useCallback, useRef, useEffect } from 'react'
import { FixedSizeGrid as Grid, GridChildComponentProps } from 'react-window'
import classNames from 'classnames'
import { useStore } from '../store'
import { columnIndexToLabel, makeCellAddress } from '../utils/cellAddresses'
import { evaluateDisplay, isFormula } from '../formula'

const ROWS = 1000
const COLS = 100
const ROW_HEIGHT = 28
const COL_WIDTH = 120

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
}

export default function SheetGrid() {
  const { sheet, selection, editing, selectCell, startEdit, setDraft, commitEdit, cancelEdit } = useStore(s => ({
    sheet: s.sheet,
    selection: s.selection,
    editing: s.editing,
    selectCell: s.selectCell,
    startEdit: s.startEdit,
    setDraft: s.setDraft,
    commitEdit: s.commitEdit,
    cancelEdit: s.cancelEdit,
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
    } = data

    const isHeaderRow = rowIndex === 0
    const isHeaderCol = columnIndex === 0
    const classes = classNames('cell', {
      'cell--header': isHeaderRow || isHeaderCol,
    })

    if (isHeaderRow && isHeaderCol) {
      return <div style={style} className={classes}></div>
    }
    if (isHeaderRow) {
      return <div style={style} className={classes}>{columnIndexToLabel(columnIndex - 1)}</div>
    }
    if (isHeaderCol) {
      return <div style={style} className={classes}>{rowIndex}</div>
    }
    const addr = makeCellAddress(columnIndex - 1, rowIndex - 1)
    const isSelected = selection.row === rowIndex && selection.col === columnIndex
    const isEditingHere = editing.addr === addr
    const cell = sheet.cells[addr]
    const display = cell ? (isFormula(cell.value) ? evaluateDisplay(addr, sheet) : cell.value) : ''

    const fmt = cell?.format
    const styleMerged: React.CSSProperties = { ...style, fontWeight: fmt?.bold ? '700' : undefined, fontStyle: fmt?.italic ? 'italic' : undefined, textDecoration: fmt?.underline ? 'underline' : undefined, color: fmt?.textColor, backgroundColor: fmt?.fillColor }

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
        columnCount={COLS + 1}
        columnWidth={COL_WIDTH}
        height={window.innerHeight - 130}
        rowCount={ROWS + 1}
        rowHeight={ROW_HEIGHT}
        width={Math.min(window.innerWidth - 20, (COLS + 1) * COL_WIDTH)}
        itemData={{
          sheet,
          selection,
          editing,
          selectCell,
          startEdit,
          setDraft,
          commitEdit,
          cancelEdit,
        } as GridData}
      >
        {Cell}
      </Grid>
    </div>
  )
}
