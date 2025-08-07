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
}
