import React, { useEffect, useMemo, useRef, useState } from 'react'
import SheetGrid from './grid/SheetGrid'
import { useStore } from './store'
import { makeCellAddress } from './utils/cellAddresses'
import { exportToCSV, exportWorkbookToXLSX, importFromFile } from './io/xlsx'
import classNames from 'classnames'

// ---------------------------------------------------------------------------
// Build / release identifier
// ---------------------------------------------------------------------------
const VERSION = 'v2025.08.08-2'

// ---------------------------------------------------------------------------
// Theme helpers
// ---------------------------------------------------------------------------
const THEME_KEY = 'excel-clone/theme'
const getInitialTheme = (): 'light' | 'dark' => {
  try {
    const stored = localStorage.getItem(THEME_KEY)
    if (stored === 'light' || stored === 'dark') return stored
  } catch {}
  // fallback to prefers-color-scheme
  return window.matchMedia &&
    window.matchMedia('(prefers-color-scheme: dark)').matches
    ? 'dark'
    : 'light'
}

export default function App() {
  /* --------------------------------------------------------------------- */
  /* Theme state                                                           */
  /* --------------------------------------------------------------------- */
  const [theme, setTheme] = useState<'light' | 'dark'>(getInitialTheme)

  // Apply / persist on change
  useEffect(() => {
    const el = document.documentElement
    if (theme === 'dark') el.classList.add('theme-dark')
    else el.classList.remove('theme-dark')
    try { localStorage.setItem(THEME_KEY, theme) } catch {}
  }, [theme])

  const {
    workbook,
    selection,
    editing,
    selectCell,
    startEdit,
    setDraft,
    commitEdit,
    cancelEdit,
    toggleFormat,
    setTextColor,
    setFillColor,
    undo,
    redo,
    toAOA,
    toAOAAll,
    fromAOA,
    addSheet,
    renameSheet,
    deleteSheet,
    setActiveSheet,
  } = useStore(s => ({
    workbook: s.workbook,
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
    toAOAAll: s.toAOAAll,
    fromAOA: s.fromAOA,

    // tab actions
    addSheet: s.addSheet,
    renameSheet: s.renameSheet,
    deleteSheet: s.deleteSheet,
    setActiveSheet: s.setActiveSheet,
  }))

  // Derive the active sheet reactively from the workbook
  const sheet = useMemo(() => workbook.sheets[workbook.activeIndex], [workbook])

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
    exportWorkbookToXLSX('workbook.xlsx', toAOAAll())
  }

  // -----------------------------------------------------------------------
  // Tab-bar local rename state
  // -----------------------------------------------------------------------
  const [renamingIndex, setRenamingIndex] = useState<number | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const renameInputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (renamingIndex != null) {
      renameInputRef.current?.focus()
    }
  }, [renamingIndex])

  return (
    <div className="app">
      <div className="toolbar">
        {/* format buttons */}
        <div className="segmented">
          <button className="btn" onClick={() => selectedAddr && toggleFormat(selectedAddr, 'bold')}>
            <b>B</b>
          </button>
          <button className="btn" onClick={() => selectedAddr && toggleFormat(selectedAddr, 'italic')}>
            <i>I</i>
          </button>
          <button className="btn" onClick={() => selectedAddr && toggleFormat(selectedAddr, 'underline')}>
            <u>U</u>
          </button>
        </div>

        <label className="color-picker">Text <input type="color" onChange={(e) => selectedAddr && setTextColor(selectedAddr, e.target.value)} /></label>
        <label className="color-picker">Fill <input type="color" onChange={(e) => selectedAddr && setFillColor(selectedAddr, e.target.value)} /></label>
        <span className="spacer" />
        {/* undo / redo */}
        <div className="segmented">
          <button className="btn" onClick={undo}>Undo</button>
          <button className="btn" onClick={redo}>Redo</button>
        </div>
        <span className="spacer" />
        {/* file actions */}
        <div className="segmented">
          <button className="btn" onClick={onExportCSV}>Export CSV</button>
          <button className="btn" onClick={onExportXLSX}>Export XLSX</button>
          <label className="import-btn btn">
            Import
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              onChange={(e) => e.target.files && e.target.files[0] && onImport(e.target.files[0])}
            />
          </label>
        </div>
        {/* dark-mode toggle */}
        <button
          className="btn"
          onClick={() => setTheme(t => (t === 'light' ? 'dark' : 'light'))}
        >
          {theme === 'light' ? 'Dark mode' : 'Light mode'}
        </button>
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

      {/* ---------------------------------------------------------------- */}
      {/* Bottom Tab Bar                                                  */}
      {/* ---------------------------------------------------------------- */}
      <div className="tab-bar">
        {workbook.sheets.map((sh, idx) => (
          <div
            key={sh.id}
            className={classNames('tab', { active: idx === workbook.activeIndex })}
            onClick={() => setActiveSheet(idx)}
            onDoubleClick={() => {
              setRenamingIndex(idx)
              setRenameDraft(sh.name ?? `Sheet ${idx + 1}`)
            }}
          >
            {renamingIndex === idx ? (
              <input
                ref={renameInputRef}
                value={renameDraft}
                onChange={e => setRenameDraft(e.target.value)}
                onBlur={() => {
                  renameSheet(idx, renameDraft.trim() || `Sheet ${idx + 1}`)
                  setRenamingIndex(null)
                }}
                onKeyDown={e => {
                  if (e.key === 'Enter') {
                    renameSheet(idx, renameDraft.trim() || `Sheet ${idx + 1}`)
                    setRenamingIndex(null)
                  }
                  if (e.key === 'Escape') {
                    setRenamingIndex(null)
                  }
                }}
              />
            ) : (
              <>
                <span className="name">{sh.name}</span>
                {workbook.sheets.length > 1 && (
                  <button
                    className="delete-tab"
                    onClick={e => {
                      e.stopPropagation()
                      deleteSheet(idx)
                    }}
                  >
                    ×
                  </button>
                )}
              </>
            )}
          </div>
        ))}
        <button className="tab add-tab" onClick={addSheet}>＋</button>
      </div>
    </div>
  )
}
