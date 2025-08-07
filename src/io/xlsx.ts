import * as XLSX from 'xlsx'

export function exportToXLSX(filename: string, aoa: (string | number)[][]) {
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet(aoa)
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : filename + '.xlsx')
}

export function exportToCSV(filename: string, aoa: (string | number)[][]) {
  const ws = XLSX.utils.aoa_to_sheet(aoa)
  const csv = XLSX.utils.sheet_to_csv(ws)
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename.endsWith('.csv') ? filename : filename + '.csv'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

export async function importFromFile(file: File): Promise<(string | number)[][]> {
  const ext = file.name.toLowerCase().split('.').pop() || ''
  if (ext === 'csv') {
    const text = await file.text()
    const wb = XLSX.read(text, { type: 'string' })
    const sheetName = wb.SheetNames[0]
    const ws = wb.Sheets[sheetName]
    return XLSX.utils.sheet_to_json(ws, { header: 1 }) as any
  } else {
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })
    const sheetName = wb.SheetNames[0]
    const ws = wb.Sheets[sheetName]
    return XLSX.utils.sheet_to_json(ws, { header: 1 }) as any
  }
}
