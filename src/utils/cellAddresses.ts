export function columnIndexToLabel(index: number): string {
  let label = ''
  let i = index
  do {
    const rem = i % 26
    label = String.fromCharCode(65 + rem) + label
    i = Math.floor(i / 26) - 1
  } while (i >= 0)
  return label
}

export function columnLabelToIndex(label: string): number {
  let res = 0
  const up = label.toUpperCase()
  for (let i = 0; i < up.length; i++) {
    const code = up.charCodeAt(i) - 64
    if (code < 1 || code > 26) throw new Error('Invalid column label')
    res = res * 26 + code
  }
  return res - 1
}

export function parseCellAddress(addr: string): { col: number; row: number } {
  const m = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(addr)
  if (!m) throw new Error(`Invalid cell address: ${addr}`)
  const col = columnLabelToIndex(m[1])
  const row = parseInt(m[2], 10) - 1
  return { col, row }
}

export function makeCellAddress(col: number, row: number): string {
  return `${columnIndexToLabel(col)}${row + 1}`
}

export function expandRange(a: string, b: string): string[] {
  const A = parseCellAddress(a)
  const B = parseCellAddress(b)
  const c1 = Math.min(A.col, B.col)
  const c2 = Math.max(A.col, B.col)
  const r1 = Math.min(A.row, B.row)
  const r2 = Math.max(A.row, B.row)
  const out: string[] = []
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      out.push(makeCellAddress(c, r))
    }
  }
  return out
}
