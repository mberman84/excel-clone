export type CellFormat = {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  textColor?: string // hex, e.g. #000000
  fillColor?: string // hex
}

export type Cell = {
  value: string // raw string; formulas start with '='
  format?: CellFormat
}

export type Sheet = {
  cells: Record<string, Cell>
  /**
   * Optional unique identifier for this sheet.  Helpful when persisting an
   * array of tabs so React key stability is guaranteed even after re-order.
   */
  id?: string
  /**
   * Human-readable tab name (defaults to “Sheet 1”, “Sheet 2”, …).
   */
  name?: string
  /**
   * Optional per-column widths (px). Index 0 => column A, etc.
   * Undefined index => default width from grid component.
   */
  colWidths?: number[]
  /**
   * Optional per-row heights (px). Index 0 => row 1, etc.
   * Undefined index => default height from grid component.
   */
  rowHeights?: number[]
}
