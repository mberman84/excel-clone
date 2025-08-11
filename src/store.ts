import { create } from 'zustand'
import { Sheet, Cell, CellFormat } from './types'
import { makeCellAddress, parseCellAddress } from './utils/cellAddresses'
import { evaluateDisplay, isFormula } from './formula'
import { v4 as uuidv4 } from 'uuid'

const STORAGE_KEY = 'excel-clone/workbook/v1'
const LEGACY_STORAGE_KEY = 'excel-clone/sheet/v1'
export const DEFAULT_COL_WIDTH = 120
export const DEFAULT_ROW_HEIGHT = 28

function emptySheet(index = 0): Sheet {
  return { 
    id: uuidv4(),
    name: `Sheet ${index + 1}`,
    cells: {}, 
    colWidths: [], 
    rowHeights: [] 
  }
}

type Workbook = {
  sheets: Sheet[]
  activeIndex: number
}

function emptyWorkbook(): Workbook {
  return {
    sheets: [emptySheet(0)],
    activeIndex: 0
  }
}

type State = {
  workbook: Workbook
  selection: { row: number; col: number; endRow: number; endCol: number }
  editing: { addr: string | null; draft: string }
  past: Workbook[]
  future: Workbook[]
  
  // Sheet navigation
  addSheet: () => void
  renameSheet: (index: number, name: string) => void
  deleteSheet: (index: number) => void
  setActiveSheet: (index: number) => void
  
  // Cell selection and editing
  selectCell: (row: number, col: number) => void
  setSelectionEnd: (row: number, col: number) => void
  selectRange: (sr: number, sc: number, er: number, ec: number) => void
  startEdit: (addr: string) => void
  setDraft: (v: string) => void
  commitEdit: () => void
  cancelEdit: () => void
  setCellValue: (addr: string, value: string) => void
  
  // Formatting
  toggleFormat: (addr: string, key: keyof CellFormat) => void
  setTextColor: (addr: string, color: string) => void
  setFillColor: (addr: string, color: string) => void
  
  // Sizing
  setColWidth: (col: number, px: number) => void
  setRowHeight: (row: number, px: number) => void
  
  // Workbook operations
  clearSheet: () => void
  undo: () => void
  redo: () => void
  replaceSheet: (s: Sheet) => void
  
  // Data conversion
  getUsedRange: () => { maxRow: number; maxCol: number }
  toAOA: () => (string | number)[][]
  toAOAAll: () => { name: string, data: (string | number)[][] }[]
  fromAOA: (data: (string | number)[][]) => void
  
  // Sorting
  sortByColumn: (col: number, direction: 'asc' | 'desc') => void
}

function cloneSheet(s: Sheet): Sheet {
  return {
    id: s.id || uuidv4(),
    name: s.name || 'Sheet',
    cells: { ...s.cells },
    colWidths: s.colWidths ? [...s.colWidths] : [],
    rowHeights: s.rowHeights ? [...s.rowHeights] : [],
  }
}

function cloneWorkbook(wb: Workbook): Workbook {
  return {
    sheets: wb.sheets.map(s => cloneSheet(s)),
    activeIndex: wb.activeIndex
  }
}

function saveLocal(wb: Workbook) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(wb))
}

function loadLocal(): Workbook | null {
  try {
    // Try to load from new workbook format
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw) {
      const parsed = JSON.parse(raw)
      if (parsed && typeof parsed === 'object' && Array.isArray(parsed.sheets)) {
        // Ensure all sheets have proper structure
        const sheets = parsed.sheets.map((s: any, i: number) => {
          if (!s || typeof s !== 'object' || !s.cells) {
            return emptySheet(i)
          }
          return {
            id: s.id || uuidv4(),
            name: s.name || `Sheet ${i + 1}`,
            cells: s.cells || {},
            colWidths: Array.isArray(s.colWidths) ? s.colWidths : [],
            rowHeights: Array.isArray(s.rowHeights) ? s.rowHeights : [],
          }
        })
        
        // Validate activeIndex
        const activeIndex = typeof parsed.activeIndex === 'number' && 
                            parsed.activeIndex >= 0 && 
                            parsed.activeIndex < sheets.length
                            ? parsed.activeIndex : 0
        
        return { sheets, activeIndex }
      }
    }
    
    // Try to migrate from legacy single-sheet format
    const legacyRaw = localStorage.getItem(LEGACY_STORAGE_KEY)
    if (legacyRaw) {
      const parsed = JSON.parse(legacyRaw)
      if (parsed && typeof parsed === 'object' && parsed.cells) {
        const sheet = {
          id: uuidv4(),
          name: 'Sheet 1',
          cells: parsed.cells,
          colWidths: Array.isArray(parsed.colWidths) ? parsed.colWidths : [],
          rowHeights: Array.isArray(parsed.rowHeights) ? parsed.rowHeights : [],
        }
        return {
          sheets: [sheet],
          activeIndex: 0
        }
      }
    }
    
    return null
  } catch { 
    return null 
  }
}

export const useStore = create<State>((set, get) => ({
  workbook: loadLocal() || emptyWorkbook(),
  selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
  editing: { addr: null, draft: '' },
  past: [],
  future: [],

  // Sheet navigation
  addSheet: () => {
    const { workbook, past } = get()
    const newWorkbook = cloneWorkbook(workbook)
    const newSheet = emptySheet(newWorkbook.sheets.length)
    newWorkbook.sheets.push(newSheet)
    newWorkbook.activeIndex = newWorkbook.sheets.length - 1
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [],
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' }
    })
  },
  
  renameSheet: (index, name) => {
    const { workbook, past } = get()
    if (index < 0 || index >= workbook.sheets.length) return
    
    const newWorkbook = cloneWorkbook(workbook)
    newWorkbook.sheets[index].name = name
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },
  
  deleteSheet: (index) => {
    const { workbook, past } = get()
    if (index < 0 || index >= workbook.sheets.length || workbook.sheets.length <= 1) return
    
    const newWorkbook = cloneWorkbook(workbook)
    newWorkbook.sheets.splice(index, 1)
    
    // Adjust active index if needed
    if (newWorkbook.activeIndex >= newWorkbook.sheets.length) {
      newWorkbook.activeIndex = newWorkbook.sheets.length - 1
    }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [],
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' }
    })
  },
  
  setActiveSheet: (index) => {
    const { workbook, editing } = get()
    if (index < 0 || index >= workbook.sheets.length || index === workbook.activeIndex) return
    
    // Commit any pending edits before switching sheets
    if (editing.addr) {
      get().commitEdit()
    }
    
    const newWorkbook = { ...workbook, activeIndex: index }
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook,
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' }
    })
  },

  // Cell selection and editing
  selectCell: (row, col) => {
    const { editing } = get()
    // If an edit is in progress, commit it before changing selection
    if (editing.addr) {
      get().commitEdit()
    }
    // Update selection without altering editing state directly
    set({ selection: { row, col, endRow: row, endCol: col } })
  },
  
  setSelectionEnd: (row, col) => {
    // Update only the end points of the selection
    set(state => ({ 
      selection: { 
        ...state.selection, 
        endRow: row, 
        endCol: col 
      } 
    }))
  },
  
  selectRange: (sr, sc, er, ec) => {
    const { editing } = get()
    // If an edit is in progress, commit it before changing selection
    if (editing.addr) {
      get().commitEdit()
    }
    
    // Clamp values to be >= 1
    const startRow = Math.max(1, sr)
    const startCol = Math.max(1, sc)
    const endRow = Math.max(1, er)
    const endCol = Math.max(1, ec)
    
    // Update selection
    set({ 
      selection: { 
        row: startRow, 
        col: startCol, 
        endRow: endRow, 
        endCol: endCol 
      } 
    })
  },

  startEdit: (addr) => {
    const { workbook } = get()
    const sheet = workbook.sheets[workbook.activeIndex]
    set({ editing: { addr, draft: sheet.cells[addr]?.value ?? '' } })
  },
  
  setDraft: (v) => set(state => ({ editing: { ...state.editing, draft: v } })),

  commitEdit: () => {
    const { editing, workbook, past } = get()
    if (!editing.addr) return
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    if (editing.draft === '' && !activeSheet.cells[editing.addr]?.format) {
      delete activeSheet.cells[editing.addr]
    } else {
      activeSheet.cells[editing.addr] = { 
        ...(activeSheet.cells[editing.addr] ?? { value: '' }), 
        value: editing.draft 
      }
    }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [], 
      editing: { addr: null, draft: '' } 
    })
  },

  cancelEdit: () => set({ editing: { addr: null, draft: '' } }),

  setCellValue: (addr, value) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    if (value === '' && !activeSheet.cells[addr]?.format) {
      delete activeSheet.cells[addr]
    } else {
      activeSheet.cells[addr] = { 
        ...(activeSheet.cells[addr] ?? { value: '' }), 
        value 
      }
    }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  toggleFormat: (addr, key) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    const prev = activeSheet.cells[addr] ?? { value: '' }
    const fmt: CellFormat = { ...(prev.format ?? {}) }
    ;(fmt as any)[key] = !(fmt as any)[key]
    
    activeSheet.cells[addr] = { ...prev, format: fmt }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  setTextColor: (addr, color) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    const prev = activeSheet.cells[addr] ?? { value: '' }
    activeSheet.cells[addr] = { 
      ...prev, 
      format: { ...(prev.format ?? {}), textColor: color } 
    }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  setFillColor: (addr, color) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    const prev = activeSheet.cells[addr] ?? { value: '' }
    activeSheet.cells[addr] = { 
      ...prev, 
      format: { ...(prev.format ?? {}), fillColor: color } 
    }
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  setColWidth: (col, px) => {
    if (col < 0) return
    const min = 40
    const width = Math.max(min, px)
    
    const { workbook, past } = get()
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    if (!activeSheet.colWidths) activeSheet.colWidths = []
    activeSheet.colWidths[col] = width
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  setRowHeight: (row, px) => {
    if (row < 0) return
    const min = 18
    const height = Math.max(min, px)
    
    const { workbook, past } = get()
    const newWorkbook = cloneWorkbook(workbook)
    const activeSheet = newWorkbook.sheets[newWorkbook.activeIndex]
    
    if (!activeSheet.rowHeights) activeSheet.rowHeights = []
    activeSheet.rowHeights[row] = height
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ workbook: newWorkbook, past: newPast, future: [] })
  },

  clearSheet: () => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeIndex = newWorkbook.activeIndex
    
    // Replace the active sheet with an empty one, preserving its id and name
    const oldSheet = newWorkbook.sheets[activeIndex]
    const newSheet = emptySheet(activeIndex)
    newSheet.id = oldSheet.id
    newSheet.name = oldSheet.name
    
    newWorkbook.sheets[activeIndex] = newSheet
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [],
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' }
    })
  },

  undo: () => {
    const { past, future, workbook } = get()
    if (!past.length) return
    
    const prev = past[past.length - 1]
    const newPast = past.slice(0, -1)
    const newFuture = [cloneWorkbook(workbook), ...future].slice(0, 50)
    
    saveLocal(prev)
    
    set({ 
      workbook: prev, 
      past: newPast, 
      future: newFuture, 
      editing: { addr: null, draft: '' },
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 }
    })
  },

  redo: () => {
    const { past, future, workbook } = get()
    if (!future.length) return
    
    const next = future[0]
    const newFuture = future.slice(1)
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    
    saveLocal(next)
    
    set({ 
      workbook: next, 
      past: newPast, 
      future: newFuture, 
      editing: { addr: null, draft: '' },
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 }
    })
  },

  replaceSheet: (s) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeIndex = newWorkbook.activeIndex
    
    // Preserve the sheet's id and name
    const oldSheet = newWorkbook.sheets[activeIndex]
    const newSheet = cloneSheet(s)
    newSheet.id = oldSheet.id
    newSheet.name = oldSheet.name
    
    newWorkbook.sheets[activeIndex] = newSheet
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [],
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 }
    })
  },

  getUsedRange: () => {
    const { workbook } = get()
    const sheet = workbook.sheets[workbook.activeIndex]
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
    const { workbook, getUsedRange } = get()
    const sheet = workbook.sheets[workbook.activeIndex]
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
  
  toAOAAll: () => {
    const { workbook } = get()
    
    return workbook.sheets.map(sheet => {
      // For each sheet, create a toAOA-like function
      let maxRow = 0, maxCol = 0
      
      // Find the used range for this sheet
      for (const addr of Object.keys(sheet.cells)) {
        const v = sheet.cells[addr]?.value
        if (v == null || (typeof v === 'string' && v.trim() === '')) continue
        
        try {
          const { col, row } = parseCellAddress(addr)
          if (row + 1 > maxRow) maxRow = row + 1
          if (col + 1 > maxCol) maxCol = col + 1
        } catch {}
      }
      
      maxRow = Math.max(maxRow, 20)
      maxCol = Math.max(maxCol, 10)
      
      // Convert to array of arrays
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
      
      return {
        name: sheet.name || 'Sheet',
        data: arr
      }
    })
  },

  fromAOA: (data) => {
    const { workbook, past } = get()
    
    const newWorkbook = cloneWorkbook(workbook)
    const activeIndex = newWorkbook.activeIndex
    
    // Create a new sheet with the imported data
    const oldSheet = newWorkbook.sheets[activeIndex]
    const newSheet = emptySheet(activeIndex)
    newSheet.id = oldSheet.id
    newSheet.name = oldSheet.name
    
    // Fill with data
    for (let r = 0; r < data.length; r++) {
      const row = data[r]
      for (let c = 0; c < row.length; c++) {
        const addr = makeCellAddress(c, r)
        const v = row[c]
        if (v === '' || v == null) continue
        newSheet.cells[addr] = { value: String(v) }
      }
    }
    
    newWorkbook.sheets[activeIndex] = newSheet
    
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({ 
      workbook: newWorkbook, 
      past: newPast, 
      future: [],
      selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      editing: { addr: null, draft: '' }
    })
  },
  
  sortByColumn: (col, direction) => {
    // No-op if col is negative
    if (col < 0) return
    
    const { workbook, past, getUsedRange, editing } = get()
    
    // Commit any active edit before sorting
    if (editing.addr) {
      get().commitEdit()
    }
    
    const newWorkbook = cloneWorkbook(workbook)
    const sheet = newWorkbook.sheets[newWorkbook.activeIndex]
    const { maxRow } = getUsedRange()
    
    // Build an array of rows with their sort keys
    const rows = []
    for (let r = 0; r < maxRow; r++) {
      const addr = makeCellAddress(col, r)
      const cell = sheet.cells[addr]
      
      // Get the display value (evaluated if formula)
      let key = ''
      if (cell) {
        key = isFormula(cell.value) ? evaluateDisplay(addr, sheet) : cell.value
      }
      
      // Determine if the key is numeric
      const num = Number(key)
      const isNum = !Number.isNaN(num) && key !== ''
      
      rows.push({
        oldRow: r,
        key: isNum ? num : key,
        keyStr: String(key),
        isNum
      })
    }
    
    // Sort the rows
    rows.sort((a, b) => {
      // Handle empty values - they should sort last for ascending, first for descending
      const aEmpty = a.keyStr === ''
      const bEmpty = b.keyStr === ''
      
      if (aEmpty && !bEmpty) return direction === 'asc' ? 1 : -1
      if (!aEmpty && bEmpty) return direction === 'asc' ? -1 : 1
      if (aEmpty && bEmpty) return a.oldRow - b.oldRow // stable sort
      
      // If both are numeric, sort numerically
      if (a.isNum && b.isNum) {
        return direction === 'asc' 
          ? (a.key as number) - (b.key as number) 
          : (b.key as number) - (a.key as number)
      }
      
      // Otherwise sort as strings
      const cmp = a.keyStr.localeCompare(b.keyStr)
      return direction === 'asc' ? cmp : -cmp
    })
    
    // Create a mapping from old row to new row
    const oldToNew: Record<number, number> = {}
    rows.forEach((row, newIdx) => {
      oldToNew[row.oldRow] = newIdx
    })
    
    // Remap cells
    const newCells: Record<string, Cell> = {}
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      try {
        const { col: c, row: r } = parseCellAddress(addr)
        
        if (r < maxRow) {
          // This row is within the sort range, remap it
          const newR = oldToNew[r]
          const newAddr = makeCellAddress(c, newR)
          newCells[newAddr] = cell
        } else {
          // This row is outside the sort range, keep it as is
          newCells[addr] = cell
        }
      } catch {
        // If parsing fails, keep the cell as is
        newCells[addr] = cell
      }
    }
    
    // Remap row heights
    const newRowHeights: number[] = []
    if (sheet.rowHeights) {
      // Copy heights for rows that are being sorted
      for (let r = 0; r < maxRow; r++) {
        if (r in oldToNew && sheet.rowHeights[r] !== undefined) {
          const newR = oldToNew[r]
          newRowHeights[newR] = sheet.rowHeights[r]
        }
      }
      
      // Preserve heights for rows beyond the sort range
      for (let r = maxRow; r < sheet.rowHeights.length; r++) {
        if (sheet.rowHeights[r] !== undefined) {
          newRowHeights[r] = sheet.rowHeights[r]
        }
      }
    }
    
    // Update the sheet with new cells and row heights
    sheet.cells = newCells
    sheet.rowHeights = newRowHeights
    
    // Update history and save
    const newPast = [...past, cloneWorkbook(workbook)].slice(-50)
    saveLocal(newWorkbook)
    
    set({
      workbook: newWorkbook,
      past: newPast,
      future: []
    })
  }
}))
