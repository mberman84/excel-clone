import { Sheet } from './types'
import { expandRange } from './utils/cellAddresses'

type Token =
  | { t: 'num'; v: number }
  | { t: 'cell'; a: string }
  | { t: 'op'; v: '+' | '-' | '*' | '/' | '^' | ':' | 'neg' }
  | { t: 'lp' }
  | { t: 'rp' }
  | { t: 'comma' }
  | { t: 'func'; name: string }
  | { t: 'call'; name: string; argc: number }

const isWhite = (c: string) => c === ' ' || c === '\t' || c === '\n'
const isDigit = (c: string) => c >= '0' && c <= '9'
const isLetter = (c: string) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')

export function isFormula(raw: string): boolean {
  return raw.trim().startsWith('=')
}

function tokenize(expr: string): Token[] {
  const s = expr.trim().replace(/^=/, '')
  const tokens: Token[] = []
  const n = s.length
  let i = 0
  while (i < n) {
    const ch = s[i]
    if (isWhite(ch)) { i++; continue }
    if (ch === ',') { tokens.push({ t: 'comma' }); i++; continue }
    if (ch === '(') { tokens.push({ t: 'lp' }); i++; continue }
    if (ch === ')') { tokens.push({ t: 'rp' }); i++; continue }
    if (ch === ':' ) { tokens.push({ t: 'op', v: ':' }); i++; continue }
    if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^') {
      tokens.push({ t: 'op', v: ch })
      i++
      continue
    }
    if (isDigit(ch) || (ch === '.' && i + 1 < n && isDigit(s[i + 1]))) {
      let j = i
      while (j < n && (isDigit(s[j]) || s[j] === '.')) j++
      const num = parseFloat(s.slice(i, j))
      tokens.push({ t: 'num', v: num })
      i = j
      continue
    }
    if (isLetter(ch) || ch === '$') {
      // identifier or cell ref (letters+digits). allow $ in addr
      let j = i
      while (j < n && (isLetter(s[j]) || s[j] === '$')) j++
      const head = s.slice(i, j)
      let k = j
      while (k < n && isDigit(s[k])) k++
      if (k > j) {
        // cell ref like A1 or $A$1
        const addr = (head + s.slice(j, k)).replaceAll('$', '').toUpperCase()
        tokens.push({ t: 'cell', a: addr })
        i = k
        continue
      } else {
        // function? lookahead if next non-space is '('
        let m = k
        while (m < n && isWhite(s[m])) m++
        const name = head.toUpperCase()
        if (m < n && s[m] === '(') {
          tokens.push({ t: 'func', name })
          i = j
          continue
        } else {
          // bare identifier not supported => treat as error by pushing func and expecting ()
          tokens.push({ t: 'func', name })
          i = j
          continue
        }
      }
    }
    throw new Error(`Unexpected character '${ch}' in formula`)
  }
  return tokens
}

function toRPN(tokens: Token[]): Token[] {
  const out: Token[] = []
  const ops: Token[] = []
  const calls: { name: string; argc: number }[] = []
  let prev: Token | undefined
  const prec: Record<string, { p: number; r: 'L' | 'R' }> = {
    ':': { p: 5, r: 'L' },
    '^': { p: 4, r: 'R' },
    '*': { p: 3, r: 'L' },
    '/': { p: 3, r: 'L' },
    '+': { p: 2, r: 'L' },
    '-': { p: 2, r: 'L' },
    'neg': { p: 6, r: 'R' },
  }

  for (let i = 0; i < tokens.length; i++) {
    let t = tokens[i]
    // rewrite unary '-'
    if (t.t === 'op' && t.v === '-') {
      const prevType = prev?.t
      if (!prev || prevType === 'op' || prevType === 'lp' || prevType === 'comma' || prevType === 'func') {
        t = { t: 'op', v: 'neg' }
      }
    }

    if (t.t === 'num' || t.t === 'cell') {
      out.push(t)
    } else if (t.t === 'func') {
      ops.push(t)
    } else if (t.t === 'comma') {
      while (ops.length && ops[ops.length - 1].t !== 'lp') {
        out.push(ops.pop()!)
      }
      if (!ops.length) throw new Error('Misplaced comma')
      if (!calls.length) throw new Error('Comma outside function')
      calls[calls.length - 1].argc++
    } else if (t.t === 'lp') {
      // if previous was a func, start a call frame
      const top = ops[ops.length - 1]
      if (top && top.t === 'func') {
        calls.push({ name: (top as any).name, argc: 0 })
      }
      ops.push(t)
    } else if (t.t === 'rp') {
      while (ops.length && ops[ops.length - 1].t !== 'lp') {
        out.push(ops.pop()!)
      }
      if (!ops.length) throw new Error('Mismatched parentheses')
      ops.pop() // pop lp
      // if top of ops is a func, emit call
      if (ops.length && ops[ops.length - 1].t === 'func') {
        const fn = ops.pop() as Extract<Token, { t: 'func' }>
        const frame = calls.pop()
        const argc = frame ? frame.argc + (frame.argc > 0 || prev?.t !== 'lp' ? 1 : 0) : 0
        out.push({ t: 'call', name: fn.name, argc })
      }
    } else if (t.t === 'op') {
      while (
        ops.length &&
        ops[ops.length - 1].t === 'op' &&
        ((prec[(ops[ops.length - 1] as any).v].p > prec[t.v].p) ||
          (prec[(ops[ops.length - 1] as any).v].p === prec[t.v].p && prec[t.v].r === 'L'))
      ) {
        out.push(ops.pop()!)
      }
      ops.push(t)
    }
    prev = t
  }
  while (ops.length) {
    const op = ops.pop()!
    if (op.t === 'lp' || op.t === 'rp') throw new Error('Mismatched parentheses')
    if (op.t === 'func') throw new Error('Function call missing parentheses')
    out.push(op)
  }
  return out
}

function toNumber(x: any): number {
  if (typeof x === 'number') return x
  const n = typeof x === 'string' ? parseFloat(x) : NaN
  return isFinite(n) ? n : 0
}

function evalRPN(rpn: Token[], resolveCell: (addr: string) => any): any {
  const st: any[] = []
  for (const t of rpn) {
    if (t.t === 'num') st.push(t.v)
    else if (t.t === 'cell') st.push({ __cell__: t.a })
    else if (t.t === 'op') {
      if (t.v === 'neg') {
        const a = st.pop()
        st.push(-toNumber(resolveVal(a)))
        continue
      }
      if (t.v === ':') {
        const b = st.pop()
        const a = st.pop()
        const A = addrOf(a)
        const B = addrOf(b)
        if (!A || !B) throw new Error('Range operator requires cell refs')
        st.push({ __range__: [A, B] })
        continue
      }
      const b = st.pop()
      const a = st.pop()
      const av = toNumber(resolveVal(a))
      const bv = toNumber(resolveVal(b))
      switch (t.v) {
        case '+': st.push(av + bv); break
        case '-': st.push(av - bv); break
        case '*': st.push(av * bv); break
        case '/': st.push(bv === 0 ? '#DIV/0!' : av / bv); break
        case '^': st.push(Math.pow(av, bv)); break
      }
    } else if (t.t === 'call') {
      const args: any[] = []
      for (let i = 0; i < t.argc; i++) args.unshift(st.pop())
      st.push(callFunc(t.name, args, resolveCell))
    }
  }
  return st.pop()

  function resolveVal(x: any): any {
    if (!x) return 0
    if (typeof x === 'object' && x.__cell__) return resolveCell(x.__cell__)
    if (typeof x === 'object' && x.__range__) {
      // resolve to array of values
      const [a, b] = x.__range__ as [string, string]
      const cells = expandRange(a, b)
      return cells.map(c => resolveCell(c))
    }
    return x
  }
  function addrOf(x: any): string | null {
    if (typeof x === 'object' && x.__cell__) return x.__cell__
    return null
  }
}

function callFunc(name: string, args: any[], resolveCell: (addr: string) => any): any {
  const N = name.toUpperCase()
  const flat: any[] = []
  for (const a of args) {
    if (Array.isArray(a)) {
      for (const v of a) flat.push(v)
    } else if (typeof a === 'object' && (a as any).__range__) {
      const [ra, rb] = (a as any).__range__ as [string, string]
      for (const addr of expandRange(ra, rb)) flat.push(resolveCell(addr))
    } else if (typeof a === 'object' && (a as any).__cell__) {
      flat.push(resolveCell((a as any).__cell__))
    } else {
      flat.push(a)
    }
  }
  const nums = flat.map(v => (typeof v === 'number' ? v : (typeof v === 'string' ? parseFloat(v) : NaN))).filter(n => isFinite(n))
  switch (N) {
    case 'SUM': return nums.reduce((a, b) => a + b, 0)
    case 'AVERAGE': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0
    case 'MIN': return nums.length ? Math.min(...nums) : 0
    case 'MAX': return nums.length ? Math.max(...nums) : 0
    case 'COUNT': return nums.length
    default: throw new Error(`Unknown function ${name}`)
  }
}

export function evaluateCell(addr: string, sheet: Sheet, cache: Map<string, any> = new Map(), visiting: Set<string> = new Set()): { value: number | string; error?: string } {
  if (cache.has(addr)) return { value: cache.get(addr) }
  if (visiting.has(addr)) return { value: '#CYCLE!', error: '#CYCLE!' }
  visiting.add(addr)
  const cell = sheet.cells[addr]
  if (!cell || cell.value.trim() === '') {
    cache.set(addr, '')
    visiting.delete(addr)
    return { value: '' }
  }
  const raw = cell.value
  if (!isFormula(raw)) {
    cache.set(addr, raw)
    visiting.delete(addr)
    return { value: raw }
  }
  try {
    const tokens = tokenize(raw)
    const rpn = toRPN(tokens)
    const result = evalRPN(rpn, (a) => {
      const r = evaluateCell(a, sheet, cache, visiting)
      return typeof r.value === 'string' && r.value.startsWith('#') ? 0 : r.value
    })
    if (result === '#DIV/0!') {
      cache.set(addr, result)
      visiting.delete(addr)
      return { value: result, error: result }
    }
    cache.set(addr, result)
    visiting.delete(addr)
    return { value: result }
  } catch (e: any) {
    visiting.delete(addr)
    const msg = '#ERR'
    cache.set(addr, msg)
    return { value: msg, error: msg }
  }
}

export function evaluateDisplay(addr: string, sheet: Sheet): string | number {
  const cell = sheet.cells[addr]
  if (!cell) return ''
  if (!isFormula(cell.value)) return cell.value
  return evaluateCell(addr, sheet).value
}
