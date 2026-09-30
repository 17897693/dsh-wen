// formula.js — xlsx 公式**轻量重算**（第十二轮需求 3）。
//
// 目标不是"实现 Excel"，而是把**没有缓存值 / 缓存值过时**的表算出一个可用结果，
// 并且**绝不猜**：不认识的函数、不认识的语法，一律在单元格上标 `unsupported` 并保留原值。
//
// 支持（子集，写死在这份清单里，README/SKILL 同步）：
//   函数（29 个）：SUM AVERAGE MIN MAX COUNT COUNTA COUNTIF COUNTIFS SUMIF AVERAGEIF
//                  IF VLOOKUP ROUND ROUNDUP ROUNDDOWN ABS INT MOD
//                  LEFT RIGHT MID LEN TRIM UPPER LOWER CONCAT CONCATENATE TEXT VALUE
//   运算：+ - * / ^ & % （一元负号）、比较 = <> < > <= >=
//   引用：A1 $A$1 A1:B3 Sheet2!A1 Sheet2!A1:B3、TRUE/FALSE、字符串、百分比
//   错误值：#REF! #DIV/0! #VALUE! #NAME? #NUM! #N/A（带单元格地址返回）
// 所有函数名/工作表名大小写不敏感；引用解析失败 → #REF!；除零 → #DIV/0!。
// 循环引用 → 该格标 `#CIRC!`（自定义名，避免与 Excel 语义混淆）。
// 不认识的函数/引用**一律显式标 `unsupported`**（`details` 与 `unsupportedTokens` 都带标签），绝不猜值：
//   · 整列引用 `A:A` → 标签带原文：`A:A（整列引用）`；
//   · 外部工作簿 `[Book1]…` → 词法把 `[` 判为坏字符，标签 `不支持的引用/字符 "["`；
//   · `_xlfn.` 前缀（如 `_xlfn.SUM`）按函数名比对，不在上表 29 个里 → 标签 `_XLFN.SUM`
// 大表护栏 `DSH_OFFICE_RECALC_MAX_CELLS`（opt-in，见 recalcMaxCells）：未设 / `0` / 非法值 = 无上限；
// 整本单元格总数超限 → **整本原样返回**（`evaluated: 0`、`formulaCells: 0`）+ `report.skipped`，
// 调用方外显为 `stats.recalc.skipped`；绝不静默截断、不半算、不动任何公式与缓存值。

const ERROR_KINDS = ['#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NUM!', '#N/A', '#CIRC!']
export const SUPPORTED_FUNCTIONS = [
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'SUMIF', 'AVERAGEIF',
  'IF', 'VLOOKUP', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'INT', 'MOD',
  // 第二轮需求 3b：文本函数族
  'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER', 'CONCAT', 'CONCATENATE', 'TEXT', 'VALUE',
]
const isError = v => typeof v === 'string' && ERROR_KINDS.includes(v)

// ---------------------------------------------------------------------------
// 词法
// ---------------------------------------------------------------------------

function lex(src) {
  const out = []
  let i = 0
  const s = String(src)
  while (i < s.length) {
    const c = s[i]
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue }
    if (c === '"') {
      let j = i + 1
      let text = ''
      while (j < s.length) {
        if (s[j] === '"' && s[j + 1] === '"') { text += '"'; j += 2; continue }
        if (s[j] === '"') break
        text += s[j++]
      }
      out.push({ k: 'str', v: text })
      i = j + 1
      continue
    }
    const num = /^\d+(\.\d+)?([eE][-+]?\d+)?/.exec(s.slice(i))
    if (num && !/[A-Za-z_$]/.test(s[i - 1] ?? '')) { out.push({ k: 'num', v: Number(num[0]) }); i += num[0].length; continue }
    // 带引号的工作表名（`'My Sheet'!A1`）
    const q = /^'((?:[^']|'')+)'/.exec(s.slice(i))
    if (q) { out.push({ k: 'id', v: q[1].replace(/''/g, "'") }); i += q[0].length; continue }
    const id = /^(\$?[A-Za-z]{1,3}\$?\d{1,7}|[A-Za-z_][A-Za-z0-9_.]*)/.exec(s.slice(i))
    if (id) {
      out.push({ k: 'id', v: id[0] })
      i += id[0].length
      continue
    }
    if ('+-*/^&=<>(),:%!'.includes(c)) {
      const two = s.slice(i, i + 2)
      if (two === '<=' || two === '>=' || two === '<>') { out.push({ k: 'op', v: two }); i += 2; continue }
      out.push({ k: c === '!' && out.length && out[out.length - 1].k === 'id' && /^[A-Za-z_][A-Za-z0-9_.]*$/.test(out[out.length - 1].v) ? 'bang' : 'op', v: c })
      i++
      continue
    }
    out.push({ k: 'bad', v: c })
    i++
  }
  return out
}

// ---------------------------------------------------------------------------
// 语法（递归下降）
// ---------------------------------------------------------------------------

function parse(tokens) {
  let p = 0
  const peek = () => tokens[p]
  const eat = k => { const t = tokens[p]; if (t && t.k === k) { p++; return t } return null }
  const node = (type, extra) => ({ type, ...extra })

  function parseExpr() { return parseCompare() }
  function parseCompare() {
    let left = parseConcat()
    while (peek() && peek().k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
      const op = tokens[p++].v
      left = node('bin', { op, left, right: parseConcat() })
    }
    return left
  }
  function parseConcat() {
    let left = parseAdd()
    while (peek() && peek().k === 'op' && peek().v === '&') { p++; left = node('bin', { op: '&', left, right: parseAdd() }) }
    return left
  }
  function parseAdd() {
    let left = parseMul()
    while (peek() && peek().k === 'op' && (peek().v === '+' || peek().v === '-')) {
      const op = tokens[p++].v
      left = node('bin', { op, left, right: parseMul() })
    }
    return left
  }
  function parseMul() {
    let left = parsePow()
    while (peek() && peek().k === 'op' && (peek().v === '*' || peek().v === '/')) {
      const op = tokens[p++].v
      left = node('bin', { op, left, right: parsePow() })
    }
    return left
  }
  function parsePow() {
    const left = parseUnary()
    if (peek() && peek().k === 'op' && peek().v === '^') { p++; return node('bin', { op: '^', left, right: parsePow() }) }
    return left
  }
  function parseUnary() {
    if (peek() && peek().k === 'op' && (peek().v === '-' || peek().v === '+')) {
      const op = tokens[p++].v
      return node('unary', { op, value: parseUnary() })
    }
    return parsePostfix()
  }
  function parsePostfix() {
    let v = parsePrimary()
    while (peek() && peek().k === 'op' && peek().v === '%') { p++; v = node('percent', { value: v }) }
    return v
  }
  function parsePrimary() {
    const t = peek()
    if (!t) return node('err', { v: '#VALUE!' })
    if (t.k === 'num') { p++; return node('lit', { v: t.v }) }
    if (t.k === 'str') { p++; return node('lit', { v: t.v }) }
    if (t.k === 'op' && t.v === '(') { p++; const e = parseExpr(); if (peek() && peek().k === 'op' && peek().v === ')') p++; return e }
    if (t.k === 'id') {
      p++
      // Sheet!ref / Sheet!A1:B3
      let sheet = null
      let name = t.v
      // 整列引用 `A:A`（需求 3c）：显式标 unsupported，绝不猜成某个单元格
      if (sheet === null && /^\$?[A-Za-z]{1,3}$/.test(name) && peek() && peek().k === 'op' && peek().v === ':') {
        const save = p
        p++
        const r2 = peek()
        if (r2 && r2.k === 'id' && /^\$?[A-Za-z]{1,3}$/.test(r2.v)) {
          p++
          return node('badref', { v: `${name}:${r2.v}（整列引用）` })
        }
        p = save
      }
      if (peek() && peek().k === 'bang') {
        p++
        sheet = String(name).toLowerCase()   // 工作表名大小写不敏感（构造时统一小写键）
        const r = peek()
        if (r && r.k === 'id') { p++; name = r.v } else name = ''
      }
      const isRef = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/.test(name)
      if (peek() && peek().k === 'op' && peek().v === '(') {
        p++
        const args = []
        if (!(peek() && peek().k === 'op' && peek().v === ')')) {
          for (;;) {
            args.push(parseExpr())
            if (peek() && peek().k === 'op' && peek().v === ',') { p++; continue }
            break
          }
        }
        if (peek() && peek().k === 'op' && peek().v === ')') p++
        return node('call', { name: String(name).toUpperCase(), args, sheet })
      }
      if (sheet !== null) {
        // Sheet!A1 或 Sheet!A1:B3（冒号在词法里是 op）
        let to = null
        if (peek() && peek().k === 'op' && peek().v === ':') {
          p++
          const r2 = peek()
          if (r2 && r2.k === 'id') { p++; to = r2.v }
        }
        return node('ref', { sheet, from: name, to })
      }
      if (isRef) {
        let to = null
        if (peek() && peek().k === 'op' && peek().v === ':') {
          p++
          const r2 = peek()
          if (r2 && r2.k === 'id' && /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/.test(r2.v)) { p++; to = r2.v }
        }
        return node('ref', { sheet: null, from: name, to })
      }
      const up = String(name).toUpperCase()
      if (up === 'TRUE') return node('lit', { v: true })
      if (up === 'FALSE') return node('lit', { v: false })
      return node('name', { v: name })
    }
    if (t.k === 'bad') { p++; return node('badref', { v: `不支持的引用/字符 "${t.v}"` }) }
    p++
    return node('err', { v: '#VALUE!' })
  }
  const ast = parseExpr()
  return { ast, rest: tokens.slice(p) }
}

// ---------------------------------------------------------------------------
// 求值
// ---------------------------------------------------------------------------

function colToIndex(letters) {
  let c = 0
  for (const ch of letters.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64)
  return c - 1
}
function splitRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref)
  if (!m) return null
  return { col: colToIndex(m[1]), row: Number(m[2]) - 1 }
}
const numOf = v => {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (typeof v === 'string') {
    const t = v.trim()
    if (t && /^-?\d+(\.\d+)?([eE][-+]?\d+)?%?$/.test(t)) return t.endsWith('%') ? Number(t.slice(0, -1)) / 100 : Number(t)
  }
  return null
}
const cmp = (a, b) => {
  const na = numOf(a)
  const nb = numOf(b)
  if (na !== null && nb !== null) return na === nb ? 0 : na < nb ? -1 : 1
  const sa = String(a ?? '')
  const sb = String(b ?? '')
  return sa === sb ? 0 : sa < sb ? -1 : 1
}

/** Excel 的"值 → 文本"口径（布尔大写，数字用通用格式）。 */
const textOf = v => {
  if (v === null || v === undefined) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return String(v)
  return String(v)
}

/**
 * 判据匹配器（COUNTIF/SUMIF/COUNTIFS 共用）。
 * 支持 `= <> < > <= >=` 前缀 + Excel 通配符 `*` `?`（不区分大小写）。
 */
function matchCriteria(crit) {
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(String(crit).trim())
  const op = m[1] || '='
  const target = m[2]
  const wild = /[*?]/.test(target)
  const re = wild
    ? new RegExp('^' + target.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
    : null
  return v => {
    if (wild && (op === '=' || op === '<>')) {
      const hit = re.test(textOf(v))
      return op === '=' ? hit : !hit
    }
    // 数值比较只命中**数值单元格**（Excel 口径：文本/布尔不满足 ">75" 这类条件）
    if (op !== '=' && op !== '<>') {
      if (typeof v !== 'number') return false
      const n = numOf(v)
      if (n === null) return false
      const t = numOf(target)
      if (t === null) return false
      return op === '<' ? n < t : op === '>' ? n > t : op === '<=' ? n <= t : n >= t
    }
    const c = cmp(v, target)
    return op === '=' ? c === 0 : c !== 0
  }
}

/**
 * `TEXT(value, format)` 的数字子集（零依赖）：
 * 支持 `0`/`#`/`,`/`.`/`%` 组成的数字格式；**日期等自定义格式显式报 unsupported**（绝不猜）。
 */
function formatTextValue(v, fmt) {
  const f = String(fmt ?? '')
  if (!f) return textOf(v)
  if (!/^[#0,.]+%?$/.test(f)) return { __unsupported: `TEXT(格式 "${f}")` }
  const n = numOf(v)
  if (n === null) return textOf(v)
  const pct = f.includes('%')
  const body = f.replace(/%/g, '')
  const thousands = body.includes(',')
  const dot = body.indexOf('.')
  const decimals = dot >= 0 ? (body.slice(dot + 1).match(/[0#]/g) || []).length : 0
  let s = (pct ? n * 100 : n).toFixed(decimals)
  if (thousands) {
    const [ip, fp] = s.split('.')
    s = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fp ? '.' + fp : '')
  }
  return s + (pct ? '%' : '')
}

/** 单表求值上下文。 */
function makeEvaluator(sheets) {
  const cache = new Map()
  const inProgress = new Set()
  const unsupported = new Set()
  const errors = []

  const keyOf = (sheet, r, c) => `${sheet}\u0000${r}\u0000${c}`

  function cellRaw(sheet, row, col) {
    const sh = sheets.get(sheet)
    if (!sh) return { err: '#REF!' }
    const rowArr = sh.rows[row]
    if (!rowArr) return { empty: true }
    const cell = rowArr[col]
    if (cell === undefined || cell === null) return { empty: true }
    if (typeof cell === 'object') return cell
    return { v: cell }
  }

  /** 一个单元格的**值**（有公式就算公式，没有就用缓存值）。 */
  function cellValue(sheet, row, col) {
    const k = keyOf(sheet, row, col)
    if (cache.has(k)) return cache.get(k)
    const raw = cellRaw(sheet, row, col)
    if (raw.empty) return ''
    if (raw.err) return raw.err
    if (!raw.f) return raw.v
    if (inProgress.has(k)) return '#CIRC!'
    inProgress.add(k)
    let value
    try { value = evalFormula(String(raw.f).replace(/^=/, ''), sheet) } catch { value = '#VALUE!' }
    inProgress.delete(k)
    if (isError(value)) errors.push({ cell: refOf(row, col), formula: `=${raw.f}`, error: value })
    cache.set(k, value)
    return value
  }
  const refOf = (row, col) => {
    let s = ''
    let n = col + 1
    while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26) }
    return `${s}${row + 1}`
  }

  /** 引用展开成"扁平值数组"（范围按行优先；单格就是 1 个元素）。 */
  function expand(sheet, from, to) {
    const a = splitRef(from)
    if (!a) return { err: '#REF!' }
    if (!to) {
      const v = cellValue(sheet, a.row, a.col)
      return { values: [v], single: true }
    }
    const b = splitRef(to)
    if (!b) return { err: '#REF!' }
    const vals = []
    for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
      for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) vals.push(cellValue(sheet, r, c))
    }
    return { values: vals, single: false, from: a, to: b }
  }

  function evalNode(n, sheet) {
    switch (n.type) {
      case 'lit': return n.v
      case 'err': return n.v
      case 'name':
        unsupported.add(n.v)
        return { __unsupported: n.v }
      case 'badref':
        // 整列引用 `A:A` / 外部工作簿引用 `[Book1]…` 等：显式 unsupported（带原文），绝不猜值
        unsupported.add(n.v)
        return { __unsupported: n.v }
      case 'ref': {
        const e = expand(n.sheet || sheet, n.from, n.to)
        if (e.err) return e.err
        return e.single ? e.values[0] : { __array: e.values }
      }
      case 'percent': {
        const v = evalNode(n.value, sheet)
        const num = numOf(isObj(v) ? null : v)
        return num === null ? '#VALUE!' : num / 100
      }
      case 'unary': {
        const v = evalNode(n.value, sheet)
        if (isObj(v)) return v
        const num = numOf(v)
        if (num === null) return '#VALUE!'
        return n.op === '-' ? -num : num
      }
      case 'bin': {
        const l = evalNode(n.left, sheet)
        const r = evalNode(n.right, sheet)
        if (isObj(l)) return l.__array ? l : l
        if (isObj(r)) return r.__array ? r : r
        if (isError(l)) return l
        if (isError(r)) return r
        if (n.op === '&') return `${fmtVal(l)}${fmtVal(r)}`
        if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
          const c = cmp(l, r)
          return ({ '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 })[n.op]
        }
        const a = numOf(l)
        const b = numOf(r)
        if (a === null || b === null) return '#VALUE!'
        if (n.op === '+') return a + b
        if (n.op === '-') return a - b
        if (n.op === '*') return a * b
        if (n.op === '/') return b === 0 ? '#DIV/0!' : a / b
        if (n.op === '^') return a ** b
        return '#VALUE!'
      }
      case 'call': return evalCall(n, sheet)
      default: return '#VALUE!'
    }
  }
  const isObj = v => v && typeof v === 'object' && ('__array' in v || '__unsupported' in v)
  const fmtVal = v => (v === true ? 'TRUE' : v === false ? 'FALSE' : String(v ?? ''))
  /** 参数扁平化：范围展开成多个值；未支持的函数/名字直接把标记透传上去。 */
  function flatArgs(args, sheet) {
    const out = []
    const origins = []
    for (const a of args) {
      const v = evalNode(a, sheet)
      if (isObj(v)) {
        if (v.__unsupported) return { unsupported: v.__unsupported }
        for (const x of v.__array) { out.push(x); origins.push('range') }
      } else { out.push(v); origins.push('scalar') }
    }
    return { values: out, origins }
  }
  /**
   * 数字提取（需求 3b：与真 Excel 对齐）。
   * **区域来源**里的文本与布尔一律忽略（`SUM(A1:A3)` 不把文本 "3" / TRUE 算进去）；
   * **直接写在参数里**的文本/布尔参与换算（`SUM("3",TRUE)` = 4）。空值两者都忽略。
   */
  const numericOf = (vals, origins, { rangeSkipsTextBool = true } = {}) => {
    const out = []
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i]
      if (isError(v)) continue
      if (rangeSkipsTextBool && origins && origins[i] === 'range' && (typeof v === 'string' || typeof v === 'boolean')) continue
      const n = numOf(v)
      if (n !== null) out.push(n)
    }
    return out
  }

  function evalCall(n, sheet) {
    const name = n.name
    if (!SUPPORTED_FUNCTIONS.includes(name)) { unsupported.add(name); return { __unsupported: name } }
    if (name === 'IF') {
      if (n.args.length < 2) return '#VALUE!'
      const c = evalNode(n.args[0], sheet)
      if (isObj(c)) return c
      if (isError(c)) return c
      const truthy = typeof c === 'boolean' ? c : numOf(c) !== null ? numOf(c) !== 0 : Boolean(String(c ?? '').length)
      if (truthy) return evalNode(n.args[1], sheet)
      return n.args.length > 2 ? evalNode(n.args[2], sheet) : false
    }
    if (name === 'VLOOKUP') {
      // ⚠ 第 2 个参数**本身就该是区域**，不能按"值"去求（旧写法把区域判成不支持，
      // 结果 VLOOKUP 永远走 unsupported 分支）。这里只求 lookup 值/列号/是否精确，
      // 区域用引用节点单独取几何信息。
      if (n.args.length < 3) return '#VALUE!'
      const target = evalNode(n.args[0], sheet)
      if (isObj(target)) return target
      if (isError(target)) return target
      const colIdxV = evalNode(n.args[2], sheet)
      if (isObj(colIdxV)) return { __unsupported: name }
      const colIdx = numOf(colIdxV)
      let exact = true
      if (n.args.length > 3) {
        const e2 = evalNode(n.args[3], sheet)
        if (isObj(e2)) return { __unsupported: name }
        exact = e2 !== false
      }
      const e = expandRefNode(n.args[1], sheet)
      if (!e) { unsupported.add(`${name}(区域不是直接引用)`); return { __unsupported: `${name}(区域不是直接引用)` } }
      if (colIdx === null || colIdx < 1) return '#VALUE!'
      const width = Math.abs(e.to.col - e.from.col) + 1
      const height = Math.abs(e.to.row - e.from.row) + 1
      const r0 = Math.min(e.from.row, e.to.row)
      const c0 = Math.min(e.from.col, e.to.col)
      if (colIdx > width) return '#REF!'
      let best = null
      for (let r = 0; r < height; r++) {
        const v = cellValue(sheet, r0 + r, c0)
        const c = cmp(v, target)
        if (c === 0) { best = r; break }
        if (!exact && c < 0) best = r
      }
      if (best === null) return '#N/A'
      return cellValue(sheet, r0 + best, c0 + Math.round(colIdx) - 1)
    }
    const fa = flatArgs(n.args, sheet)
    if (fa.unsupported) return { __unsupported: fa.unsupported }
    const vals = fa.values
    switch (name) {
      case 'SUM': return numericOf(vals, fa.origins).reduce((a, b) => a + b, 0)
      case 'AVERAGE': {
        const nums = numericOf(vals, fa.origins)
        return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : '#DIV/0!'
      }
      case 'MIN': {
        const nums = numericOf(vals, fa.origins)
        return nums.length ? Math.min(...nums) : 0
      }
      case 'MAX': {
        const nums = numericOf(vals, fa.origins)
        return nums.length ? Math.max(...nums) : 0
      }
      case 'COUNT': return numericOf(vals, fa.origins).length
      case 'COUNTA': return vals.filter(v => v !== '' && v !== null && v !== undefined).length
      case 'COUNTIF': {
        // ⚠ 判据（第 2 个参数）**不能**混进被统计的值里：升到 switch 之前的 `flatArgs`
        // 会把所有参数摊平，判据字符串自己也就成了"总是匹配自己"的一项（实测 +1）。
        if (n.args.length < 2) return '#VALUE!'
        const only = flatArgs([n.args[0]], sheet)
        if (only.unsupported) return { __unsupported: only.unsupported }
        const crit = evalNode(n.args[1], sheet)
        if (isObj(crit)) return { __unsupported: name }
        const test = matchCriteria(crit)
        return only.values.filter(test).length
      }
      case 'SUMIF':
      case 'AVERAGEIF': {
        // 判据与区域**分开**求值（第 2 参数是判据、第 3 是求和区），同样不摊平
        if (n.args.length < 2) return '#VALUE!'
        const range = flatArgs([n.args[0]], sheet)
        if (range.unsupported) return { __unsupported: range.unsupported }
        const sumRange = n.args.length > 2 ? flatArgs([n.args[2]], sheet) : range
        if (sumRange.unsupported) return { __unsupported: sumRange.unsupported }
        const crit = evalNode(n.args[1], sheet)
        if (isObj(crit)) return { __unsupported: name }
        const test = matchCriteria(crit)
        const picked = []
        for (let i = 0; i < range.values.length; i++) if (test(range.values[i])) picked.push(sumRange.values[i])
        // 求和区里的文本/布尔同样按 Excel 忽略（picked 全部来自区域）
        const nums = numericOf(picked, picked.map(() => 'range'))
        if (name === 'SUMIF') return nums.reduce((a, b) => a + b, 0)
        return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : '#DIV/0!'
      }
      case 'COUNTIFS': {
        if (n.args.length < 2 || n.args.length % 2 !== 0) return '#VALUE!'
        const pairs = []
        for (let i = 0; i < n.args.length; i += 2) {
          const r = flatArgs([n.args[i]], sheet)
          if (r.unsupported) return { __unsupported: r.unsupported }
          const c = evalNode(n.args[i + 1], sheet)
          if (isObj(c)) return { __unsupported: name }
          pairs.push({ values: r.values, test: matchCriteria(c) })
        }
        const len = Math.min(...pairs.map(p => p.values.length))
        let count = 0
        for (let i = 0; i < len; i++) if (pairs.every(p => p.test(p.values[i]))) count++
        return count
      }
      case 'LEFT':
      case 'RIGHT': {
        const s = textOf(vals[0])
        const k = vals.length > 1 ? numOf(vals[1]) : 1
        if (k === null || k < 0) return '#VALUE!'
        const n2 = Math.floor(k)
        return name === 'LEFT' ? s.slice(0, n2) : s.slice(Math.max(0, s.length - n2))
      }
      case 'MID': {
        const s = textOf(vals[0])
        const start = numOf(vals[1])
        const k = numOf(vals[2])
        if (start === null || k === null || start < 1 || k < 0) return '#VALUE!'
        return s.substr(Math.floor(start) - 1, Math.floor(k))
      }
      case 'LEN': return textOf(vals[0]).length
      case 'TRIM': return textOf(vals[0]).replace(/^\s+|\s+$/g, '').replace(/\s+/g, ' ')
      case 'UPPER': return textOf(vals[0]).toUpperCase()
      case 'LOWER': return textOf(vals[0]).toLowerCase()
      case 'CONCAT':
      case 'CONCATENATE': return vals.map(v => (isObj(v) ? '' : textOf(v))).join('')
      case 'TEXT': {
        // 数字格式子集；日期等自定义格式显式 unsupported（绝不猜）
        const r = formatTextValue(vals[0], vals[1])
        return r && typeof r === 'object' ? { __unsupported: r.__unsupported } : r
      }
      case 'VALUE': {
        const raw = textOf(vals[0]).trim()
        const neg = /^\((.*)\)$/.exec(raw)
        const body = (neg ? `-${neg[1]}` : raw).replace(/,/g, '')
        const pct = body.endsWith('%')
        const base = pct ? body.slice(0, -1) : body
        const num = numOf(base)
        if (num === null) return '#VALUE!'
        return pct ? num / 100 : num
      }
      case 'ROUND':
      case 'ROUNDUP':
      case 'ROUNDDOWN': {
        const x = numOf(vals[0])
        const d = vals.length > 1 ? numOf(vals[1]) : 0
        if (x === null || d === null) return '#VALUE!'
        const f = 10 ** d
        if (name === 'ROUND') return Math.round((x + Number.EPSILON * Math.sign(x)) * f) / f
        if (name === 'ROUNDUP') return Math.sign(x) * Math.ceil(Math.abs(x) * f) / f
        return Math.sign(x) * Math.floor(Math.abs(x) * f) / f
      }
      case 'ABS': {
        const x = numOf(vals[0])
        return x === null ? '#VALUE!' : Math.abs(x)
      }
      case 'INT': {
        const x = numOf(vals[0])
        return x === null ? '#VALUE!' : Math.floor(x)
      }
      case 'MOD': {
        const a = numOf(vals[0])
        const b = numOf(vals[1])
        if (a === null || b === null) return '#VALUE!'
        return b === 0 ? '#DIV/0!' : a - b * Math.floor(a / b)
      }
      default: unsupported.add(name); return { __unsupported: name }
    }
  }
  /** VLOOKUP 需要拿到范围的几何信息，这里再解析一次引用节点。 */
  function expandRefNode(argNode, sheet) {
    if (!argNode || argNode.type !== 'ref') return null
    const a = splitRef(argNode.from)
    const b = argNode.to ? splitRef(argNode.to) : a
    if (!a || !b) return null
    return { from: a, to: b }
  }

  function evalFormula(src, sheet) { return evalNode(parse(lex(src)).ast, sheet) }

  return { cellValue, evalFormula, unsupported, errors }
}

/**
 * 重算的单元格上限护栏（opt-in，需求 3d）：`DSH_OFFICE_RECALC_MAX_CELLS`。
 * 未设 / 0 / 非法值 = 无上限（与旧行为逐字一致）。超限时**显式跳过**并记账，绝不静默截断。
 */
export function recalcMaxCells(raw = process.env.DSH_OFFICE_RECALC_MAX_CELLS) {
  if (raw === undefined || raw === null || raw === '') return 0
  const n = Number(raw)
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0
}

/**
 * 重算整本工作簿。
 * @param {object} model 统一内容模型里的 workbook
 * @returns {{model:object, report:object}}
 */
export function recalcWorkbook(model) {
  // 大表护栏：超过上限就把整本原样返回 + report.skipped（调用方转成 stats/notice）
  const maxCells = recalcMaxCells()
  if (maxCells > 0) {
    const totalCells = (model.sheets || []).reduce((a, sh) => a + (sh.rows || []).reduce((x, row) => x + row.length, 0), 0)
    if (totalCells > maxCells) {
      return {
        model,
        report: {
          formulaCells: 0,
          evaluated: 0,
          unsupported: 0,
          errors: 0,
          unsupportedTokens: [],
          supportedFunctions: SUPPORTED_FUNCTIONS,
          details: [],
          skipped: `表格单元格总数 ${totalCells} 超过 DSH_OFFICE_RECALC_MAX_CELLS=${maxCells}，已跳过重算（原公式与缓存值原样保留，未做任何截断）`,
        },
      }
    }
  }
  const sheets = new Map()
  for (const sh of model.sheets || []) sheets.set(String(sh.name).toLowerCase(), sh)
  const ev = makeEvaluator(sheets)
  const details = []
  const next = {
    ...model,
    sheets: (model.sheets || []).map(sh => ({
      ...sh,
      rows: (sh.rows || []).map((row, r) => row.map((cell, c) => {
        if (!cell || typeof cell !== 'object' || !cell.f) return cell
        const ref = (() => {
          let s = ''
          let n = c + 1
          while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26) }
          return `${s}${r + 1}`
        })()
        const value = ev.cellValue(String(sh.name).toLowerCase(), r, c)
        if (value && typeof value === 'object') {
          const label = value.__unsupported || '范围用在需要单值的语境'
          const detail = { sheet: sh.name, cell: ref, formula: `=${cell.f}`, value: cell.v, error: `unsupported: ${label}` }
          details.push(detail)
          return { ...cell, error: detail.error, recalc: 'unsupported' }
        }
        if (isError(value)) {
          // 错误值**照 Excel 显示**（表格里就是 `#N/A`/`#DIV/0!`），
          // 原缓存值不丢：进 details.cached，方便对照。
          const detail = { sheet: sh.name, cell: ref, formula: `=${cell.f}`, cached: cell.v, error: value }
          details.push(detail)
          return { ...cell, v: value, t: 'e', error: value, recalc: 'error' }
        }
        return { ...cell, v: value, t: typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : 's', recalc: 'ok' }
      })),
    })),
  }
  const total = (model.sheets || []).reduce((a, sh) => a + (sh.rows || []).reduce((x, row) => x + row.filter(c => c && typeof c === 'object' && c.f).length, 0), 0)
  const okCount = next.sheets.reduce((a, sh) => a + (sh.rows || []).reduce((x, row) => x + row.filter(c => c && typeof c === 'object' && c.recalc === 'ok').length, 0), 0)
  const unsupportedCount = next.sheets.reduce((a, sh) => a + (sh.rows || []).reduce((x, row) => x + row.filter(c => c && typeof c === 'object' && c.recalc === 'unsupported').length, 0), 0)
  const errorCount = next.sheets.reduce((a, sh) => a + (sh.rows || []).reduce((x, row) => x + row.filter(c => c && typeof c === 'object' && c.recalc === 'error').length, 0), 0)
  return {
    model: next,
    report: {
      formulaCells: total,
      evaluated: okCount,
      unsupported: unsupportedCount,
      errors: errorCount,
      unsupportedTokens: [...ev.unsupported].sort(),
      supportedFunctions: SUPPORTED_FUNCTIONS,
      details: details.slice(0, 100),
      note: unsupportedCount
        ? `${unsupportedCount} 个单元格含不支持的函数/语法，已保留原缓存值并标 error（**绝不猜值**）：${[...ev.unsupported].sort().join(', ')}`
        : undefined,
    },
  }
}
