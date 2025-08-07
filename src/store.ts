import { create } from 'zustand'
import { Sheet, Cell, CellFormat } from './types'
import { makeCellAddress, parseCellAddress } from './utils/cellAddresses'
import { evaluateDisplay, isFormula } from './formula'

const STORAGE_KEY = 'excel-clone/sheet/v1'

function emptySheet(): Sheet { return { cells: {} } }

type State = {
  sheet: Sheet
  selection: { row: number; col: number }
  editing: { addr: string | null; draft: string }
  past: Sheet[]
  future: Sheet[]
  selectCell: (row: number, col: number) => void
  startEdit: (addr: string) => void
  setDraft: (v: string) => void
  commitEdit: () => void
  cancelEdit: () => void
  setCellValue: (addr: string, value: string) => void
  toggleFormat: (addr: string, key: keyof CellFormat) => void
  setTextColor: (addr: string, color: string) => void
  setFillColor: (addr: string, color: string) => void
  clearSheet: () => void
  undo: () => void
  redo: () => void
  replaceSheet: (s: Sheet) => void
  getUsedRange: () => { maxRow: number; maxCol: number }
  toAOA: () => (string | number)[][]
  fromAOA: (data: (string | number)[][]) => void
}

function cloneSheet(s: Sheet): Sheet {
  return { cells: { ...s.cells } }
}

function saveLocal(s: Sheet) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(s))
}

function loadLocal(): Sheet | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return null
    const parsed = JSON.parse(raw)
    if (!parsed || typeof parsed !== 'object' || !parsed.cells) return null
    return { cells: parsed.cells }
  } catch { return null }
}

export const useStore = create<State>((set, get) => ({
  sheet: loadLocal() ?? emptySheet(),
  selection: { row: 1, col: 1 },
  editing: { addr: null, draft: '' },
  past: [],
  future: [],

  selectCell: (row, col) => {
    const { editing } = get()
    // If an edit is in progress, commit it before changing selection
    if (editing.addr) {
      get().commitEdit()
    }
    // Update selection without altering editing state directly
    set({ selection: { row, col } })
  },

  startEdit: (addr) => set({ editing: { addr, draft: get().sheet.cells[addr]?.value ?? '' } }),
  setDraft: (v) => set(state => ({ editing: { ...state.editing, draft: v } })),

  commitEdit: () => {
    const { editing, sheet, past } = get()
    if (!editing.addr) return
    const next: Sheet = cloneSheet(sheet)
    if (editing.draft === '' && !next.cells[editing.addr]?.format) {
      delete next.cells[editing.addr]
    } else {
      next.cells[editing.addr] = { ...(next.cells[editing.addr] ?? { value: '' }), value: editing.draft }
    }
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [], editing: { addr: null, draft: '' } })
  },

  cancelEdit: () => set({ editing: { addr: null, draft: '' } }),

  setCellValue: (addr, value) => {
    const { sheet, past } = get()
    const next: Sheet = cloneSheet(sheet)
    if (value === '' && !next.cells[addr]?.format) delete next.cells[addr]
    else next.cells[addr] = { ...(next.cells[addr] ?? { value: '' }), value }
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  toggleFormat: (addr, key) => {
    const { sheet, past } = get()
    const next: Sheet = cloneSheet(sheet)
    const prev = next.cells[addr] ?? { value: '' }
    const fmt: CellFormat = { ...(prev.format ?? {}) }
    ;(fmt as any)[key] = !(fmt as any)[key]
    next.cells[addr] = { ...prev, format: fmt }
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  setTextColor: (addr, color) => {
    const { sheet, past } = get()
    const next: Sheet = cloneSheet(sheet)
    const prev = next.cells[addr] ?? { value: '' }
    next.cells[addr] = { ...prev, format: { ...(prev.format ?? {}), textColor: color } }
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  setFillColor: (addr, color) => {
    const { sheet, past } = get()
    const next: Sheet = cloneSheet(sheet)
    const prev = next.cells[addr] ?? { value: '' }
    next.cells[addr] = { ...prev, format: { ...(prev.format ?? {}), fillColor: color } }
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  clearSheet: () => {
    const { sheet, past } = get()
    const next = emptySheet()
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  undo: () => {
    const { past, future, sheet } = get()
    if (!past.length) return
    const prev = past[past.length - 1]
    const newPast = past.slice(0, -1)
    const newFuture = [cloneSheet(sheet), ...future].slice(0, 50)
    saveLocal(prev)
    set({ sheet: prev, past: newPast, future: newFuture, editing: { addr: null, draft: '' } })
  },

  redo: () => {
    const { past, future, sheet } = get()
    if (!future.length) return
    const next = future[0]
    const newFuture = future.slice(1)
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: newFuture, editing: { addr: null, draft: '' } })
  },

  replaceSheet: (s) => {
    const { sheet, past } = get()
    const next = cloneSheet(s)
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },

  getUsedRange: () => {
    const { sheet } = get()
    let maxRow = 0, maxCol = 0
    for (const addr of Object.keys(sheet.cells)) {
      const v = sheet.cells[addr]?.value
      if (v == null || (typeof v === 'string' && v.trim() === '')) continue
      try {
        const { col, row } = parseCellAddress(addr)
        if (row + 1 > maxRow) maxRow = row + 1
        if (col + 1 > maxCol) maxCol = col + 1
      } catch {}
    }
    return { maxRow: Math.max(maxRow, 20), maxCol: Math.max(maxCol, 10) }
  },

  toAOA: () => {
    const { sheet, getUsedRange } = get()
    const { maxRow, maxCol } = getUsedRange()
    const arr: (string | number)[][] = []
    for (let r = 0; r < maxRow; r++) {
      const row: (string | number)[] = []
      for (let c = 0; c < maxCol; c++) {
        const addr = makeCellAddress(c, r)
        const cell = sheet.cells[addr]
        if (!cell) { row.push(''); continue }
        const disp = isFormula(cell.value) ? evaluateDisplay(addr, sheet) : cell.value
        row.push(disp as any)
      }
      arr.push(row)
    }
    return arr
  },

  fromAOA: (data) => {
    const next: Sheet = emptySheet()
    for (let r = 0; r < data.length; r++) {
      const row = data[r]
      for (let c = 0; c < row.length; c++) {
        const addr = makeCellAddress(c, r)
        const v = row[c]
        if (v === '' || v == null) continue
        next.cells[addr] = { value: String(v) }
      }
    }
    const { sheet, past } = get()
    const newPast = [...past, cloneSheet(sheet)].slice(-50)
    saveLocal(next)
    set({ sheet: next, past: newPast, future: [] })
  },
}))
