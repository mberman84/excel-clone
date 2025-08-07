import React, { useEffect, useMemo, useRef, useState } from 'react'
import SheetGrid from './grid/SheetGrid'
import { useStore } from './store'
import { makeCellAddress } from './utils/cellAddresses'
import { exportToCSV, exportToXLSX, importFromFile } from './io/xlsx'

// ---------------------------------------------------------------------------
// Build / release identifier
// ---------------------------------------------------------------------------
const VERSION = 'v2025.08.07-6'

export default function App() {
  const { sheet, selection, editing, selectCell, startEdit, setDraft, commitEdit, cancelEdit, toggleFormat, setTextColor, setFillColor, undo, redo, toAOA, fromAOA } = useStore(s => ({
    sheet: s.sheet,
    selection: s.selection,
    editing: s.editing,
    selectCell: s.selectCell,
    startEdit: s.startEdit,
    setDraft: s.setDraft,
    commitEdit: s.commitEdit,
    cancelEdit: s.cancelEdit,
    toggleFormat: s.toggleFormat,
    setTextColor: s.setTextColor,
    setFillColor: s.setFillColor,
    undo: s.undo,
    redo: s.redo,
    toAOA: s.toAOA,
    fromAOA: s.fromAOA,
  }))

  const selectedAddr = useMemo(() => {
    if (selection.row <= 0 || selection.col <= 0) return null
    return makeCellAddress(selection.col - 1, selection.row - 1)
  }, [selection])

  useEffect(() => {
    if (selectedAddr && !editing.addr) {
      const v = sheet.cells[selectedAddr]?.value ?? ''
      setDraft(v)
    }
  }, [selectedAddr, editing.addr])

  // Global key to start editing by typing
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (editing.addr || !selectedAddr) return
      if (selection.row <= 0 || selection.col <= 0) return
      if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
        startEdit(selectedAddr)
        setTimeout(() => useStore.getState().setDraft(e.key), 0)
      }
      if (e.key === 'Backspace' || e.key === 'Delete') {
        startEdit(selectedAddr)
        setTimeout(() => useStore.getState().setDraft(''), 0)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [editing.addr, selectedAddr, selection])

  const onImport = async (file: File) => {
    const aoa = await importFromFile(file)
    fromAOA(aoa)
  }

  const onExportCSV = () => {
    exportToCSV('sheet.csv', toAOA())
  }
  const onExportXLSX = () => {
    exportToXLSX('sheet.xlsx', toAOA())
  }

  return (
    <div className="app">
      <div className="toolbar">
        <button onClick={() => selectedAddr && toggleFormat(selectedAddr, 'bold')}><b>B</b></button>
        <button onClick={() => selectedAddr && toggleFormat(selectedAddr, 'italic')}><i>I</i></button>
        <button onClick={() => selectedAddr && toggleFormat(selectedAddr, 'underline')}><u>U</u></button>
        <label className="color-picker">Text <input type="color" onChange={(e) => selectedAddr && setTextColor(selectedAddr, e.target.value)} /></label>
        <label className="color-picker">Fill <input type="color" onChange={(e) => selectedAddr && setFillColor(selectedAddr, e.target.value)} /></label>
        <span className="spacer" />
        <button onClick={undo}>Undo</button>
        <button onClick={redo}>Redo</button>
        <span className="spacer" />
        <button onClick={() => exportToCSV('sheet.csv', toAOA())}>Export CSV</button>
        <button onClick={() => exportToXLSX('sheet.xlsx', toAOA())}>Export XLSX</button>
        <label className="import-btn">Import<input type="file" accept=".csv,.xlsx,.xls" onChange={(e) => e.target.files && e.target.files[0] && onImport(e.target.files[0])} /></label>
        {/* build version shown for easy cache-busting verification */}
        <span className="version-badge">{VERSION}</span>
      </div>

      <div className="formula-bar">
        <div className="addr">{selectedAddr ?? ''}</div>
        <input
          value={editing.addr ? editing.draft : (selectedAddr ? (sheet.cells[selectedAddr]?.value ?? '') : '')}
          onChange={(e) => setDraft(e.target.value)}
          onFocus={() => selectedAddr && startEdit(selectedAddr)}
          onBlur={() => commitEdit()}
          onKeyDown={(e) => { if (e.key === 'Enter') commitEdit(); if (e.key === 'Escape') cancelEdit() }}
          placeholder="Type a value or =formula"
        />
      </div>

      <SheetGrid />
    </div>
  )
}
