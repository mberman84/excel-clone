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
