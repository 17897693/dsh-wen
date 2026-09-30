// PDF text extraction (pure JS) and PDF generation (Latin via base-14,
// CJK via the predefined Adobe-GB1 STSong-Light mapping — no font embedding).
import { inflateRawSync, inflateSync, deflateSync } from 'node:zlib'
import { createHash } from 'node:crypto'
import { createDecryptor } from './pdfcrypt.js'
import { FontChain, winAnsiOk } from './pdffont.js'
// 插图链路（第十二轮需求 1c；第二轮需求 4a/4b 扩到调色板 PNG / BMP / GIF 首帧）：
// JPEG 原样 DCTDecode 内嵌、其余格式解成原始采样 FlateDecode。
import { readImageBytes, sniffImage, imageToRaw } from './image.js'

// PDF FlateDecode is zlib-wrapped (RFC 1950) in practice, but raw-deflate and
// slightly damaged streams exist too: try the zlib way first, then raw.
function inflatePdf(bytes) {
  try {
    return new Uint8Array(inflateSync(bytes, { finishFlush: 2 }))
  } catch {
    return new Uint8Array(inflateRawSync(bytes, { finishFlush: 2 }))
  }
}

// exposed for diagnostics and for callers that need raw page access
export { PdfFile }

// ============================ low-level scanning ============================

class PdfFile {
  constructor(buf) {
    this.buf = buf
    this.s = buf.toString('latin1')
    this.objs = new Map() // num -> { body, start }
    this.scanObjects()
    this.setupEncryption()
    this.expandObjStms()
  }
  /**
   * Locate /Encrypt and build the standard-security-handler decryptor. Files
   * that merely restrict copying/printing open with an empty user password and
   * are decrypted transparently; a real password is reported, never guessed.
   */
  setupEncryption() {
    const ref = /\/Encrypt\s+(\d{1,10})\s+(\d{1,5})\s+R/.exec(this.s)
    let enc = null
    let encNum = null
    if (ref && this.objs.has(Number(ref[1]))) {
      encNum = Number(ref[1])
      enc = parseValue(this.objs.get(encNum).body, { p: 0 }).v
    } else {
      const inline = /\/Encrypt\s*<<([\s\S]{0,2000}?)>>/.exec(this.s)
      if (inline) enc = parseValue(`<<${inline[1]}>>`, { p: 0 }).v
    }
    if (!enc || typeof enc !== 'object') return
    this.encryptMetadata = enc.EncryptMetadata !== false
    let id0 = Buffer.alloc(0)
    const idm = /\/ID\s*\[\s*<([0-9a-fA-F\s]*)>/.exec(this.s)
    if (idm) id0 = Buffer.from(idm[1].replace(/\s/g, ''), 'hex')
    this.enc = createDecryptor(enc, id0)
    this.encNum = encNum
  }
  /** True when the file carries an /Encrypt dictionary. */
  isEncrypted() {
    return this.enc !== undefined
  }
  decryptValueDeep(value, num, gen) {
    void value; void num; void gen
    return value
  }
  parsed(num) {
    const rec = this.objs.get(num)
    if (!rec) return undefined
    // Strings inside a regular object are encrypted; objects living in an
    // /ObjStm and the /Encrypt dictionary itself are already plaintext.
    const crypt = this.enc?.supported && rec.start >= 0 && num !== this.encNum
      ? { decryptString: bytes => this.enc.decrypt(bytes, num, 0, true) }
      : null
    return parseValue(rec.body, { p: 0 }, null, crypt).v
  }
  scanObjects() {
    const re = /(^|[\r\n\s])(\d{1,10})\s+(\d{1,5})\s+obj(?![\w])/g
    let m
    while ((m = re.exec(this.s))) {
      const num = Number(m[2])
      const start = m.index + m[0].length
      const end = this.s.indexOf('endobj', start)
      const body = this.s.slice(start, end < 0 ? this.s.length : end)
      const existing = this.objs.get(num)
      if (!existing || (body.trimStart().startsWith('<<') && !existing.body.trimStart().startsWith('<<'))) {
        if (!existing || body.trimStart().startsWith('<<')) this.objs.set(num, { body, start })
      }
    }
  }
  dict(num) {
    const v = this.parsed(num)
    return v && typeof v === 'object' && !isRef(v) ? v : null
  }
  value(num) {
    return this.parsed(num) ?? null
  }
  resolve(v, depth = 0) {
    if (depth > 8) return v
    if (isRef(v)) return this.resolve(this.value(v.__ref), depth + 1)
    return v
  }
  resolveDict(v, depth = 0) {
    const r = this.resolve(v, depth)
    if (isRef(r)) return this.resolveDict(r, depth + 1)
    return r && typeof r === 'object' ? r : null
  }
  streamBytes(num) {
    if (this._streams?.has(num)) return this._streams.get(num)
    const rec = this.objs.get(num)
    if (!rec) return undefined
    const sIdx = rec.body.indexOf('stream')
    if (sIdx < 0) return undefined
    const dict = this.dict(num) || {}
    let start = rec.start + sIdx + 6
    if (this.s[start] === '\r') start++
    if (this.s[start] === '\n') start++
    let len = Number(dict.Length)
    if (!Number.isFinite(len)) {
      const eIdx = this.s.indexOf('endstream', start)
      len = eIdx > start ? eIdx - start : 0
    }
    let bytes = this.buf.subarray(start, Math.min(start + len, this.buf.length))
    const filters = normalizeFilters(dict.Filter)
    const parms = dict.DecodeParms ? this.resolveDict(dict.DecodeParms) : null
    try {
      // Object streams, content streams etc. are encrypted (compress-then-encrypt);
      // cross-reference streams and unencrypted metadata are not.
      if (this.enc?.supported && num !== this.encNum) {
        const type = String(dict.Type ?? '')
        const skip = type === 'XRef' || (type === 'Metadata' && this.encryptMetadata === false)
        if (!skip && bytes.length) bytes = this.enc.decrypt(bytes, num, 0, true)
      }
      for (let f = filters.length - 1; f >= 0; f--) {
        const name = filters[f]
        const parm = Array.isArray(dict.DecodeParms) ? this.resolveDict(dict.DecodeParms[f]) : parms
        if (name === 'FlateDecode') bytes = inflatePdf(bytes)
        else if (name === 'LZWDecode') bytes = lzwDecode(bytes)
        else if (name === 'RunLengthDecode') bytes = runLengthDecode(bytes)
        else if (name === 'ASCIIHexDecode') bytes = asciiHexDecode(bytes)
        else if (name === 'ASCII85Decode') bytes = ascii85Decode(bytes)
        else throw new Error(`unsupported filter ${name}`)
        if (parm) bytes = applyPredictor(bytes, parm)
      }
    } catch { bytes = new Uint8Array(0) }
    ;(this._streams ??= new Map()).set(num, bytes)
    return bytes
  }
  expandObjStms() {
    const candidates = [...this.objs.keys()]
    for (const num of candidates) {
      const d = this.dict(num)
      if (!d || d.Type !== 'ObjStm') continue
      const bytes = this.streamBytes(num)
      if (!bytes || !bytes.length) continue
      const text = Buffer.from(bytes).toString('latin1')
      const count = Number(d.N || 0)
      const first = Number(d.First || 0)
      const nums = text.slice(0, first).trim().split(/\s+/).map(Number)
      for (let k = 0; k < count && k * 2 + 1 < nums.length; k++) {
        const objNum = nums[k * 2]
        const off = nums[k * 2 + 1]
        if (!Number.isFinite(objNum) || !Number.isFinite(off)) continue
        if (this.objs.has(objNum)) continue
        const end = k + 1 < count ? nums[(k + 1) * 2 + 1] : text.length - first
        // `nums` 是从 `text.slice(0, first)` 读出来的 → `off` / `end` 都是**区域内
        // 相对偏移**（`end` 的兜底 `text.length - first` 也印证这点），必须整体加回
        // `first` 才是 `text` 的下标。旧写法少了这一项 → 每个 ObjStm 内对象前错
        // `First` 字节：第一个对象恰好读到 ObjStm 的头部索引表本身（`/Pages` 于是
        // 解析不出 `/Kids`，页序退化成"对象号升序"），其余对象整体错位 →
        // 页 `/Resources` 错位 → `fontsOf` 返回空表 → 2 字节 CID 被拆成两个单字节
        // 字符（"整本中文乱码"的真因，见 CHANGELOG「本轮前提更正」）。
        // `first` 缺失/为 0 时 `first + off === off`，行为与旧版逐字一致。
        const from = first + off
        // 畸形文件兜底：`first + off` 越界就跳过该对象，绝不抛 —— 一个坏对象不该
        // 让整本文档解析失败。
        if (!Number.isFinite(from) || from < 0 || from >= text.length) continue
        const to = Math.min(text.length, Math.max(from, first + Math.max(off, end)))
        this.objs.set(objNum, { body: text.slice(from, to), start: -1 })
      }
    }
  }
  catalog() {
    const t = this.s.lastIndexOf('trailer')
    if (t >= 0) {
      const m = /\/Root\s+(\d{1,10})\s+(\d{1,5})\s+R/.exec(this.s.slice(t, t + 900))
      if (m && this.objs.has(Number(m[1]))) return this.dict(Number(m[1])) || {}
    }
    for (const num of this.objs.keys()) {
      const d = this.dict(num)
      if (d && d.Type === 'Catalog') return d
    }
    return {}
  }
  infoDict() {
    const t = this.s.lastIndexOf('trailer')
    if (t >= 0) {
      const m = /\/Info\s+(\d{1,10})\s+(\d{1,5})\s+R/.exec(this.s.slice(t, t + 900))
      if (m && this.objs.has(Number(m[1]))) return this.dict(Number(m[1])) || {}
    }
    return {}
  }
  pages() {
    const out = []
    const cat = this.catalog()
    const visit = (node, attrs, depth) => {
      if (depth > 64 || out.length > 5000) return
      const d = this.resolveDict(node)
      if (!d) return
      const merged = { ...attrs }
      for (const k of ['Resources', 'MediaBox', 'Rotate']) if (d[k] !== undefined) merged[k] = d[k]
      if (d.Type === 'Pages') {
        const kids = this.resolve(d.Kids)
        if (Array.isArray(kids)) for (const k of kids) visit(k, merged, depth + 1)
        return
      }
      out.push({ dict: d, attrs: merged })
    }
    if (cat.Pages) visit(cat.Pages, {}, 0)
    if (!out.length) {
      for (const num of this.objs.keys()) {
        const d = this.dict(num)
        if (d && (d.Type === 'Page' || (!d.Type && d.Contents))) out.push({ dict: d, attrs: { Resources: d.Resources, MediaBox: d.MediaBox } })
      }
    }
    return out
  }
  fontsOf(resources) {
    const out = new Map()
    const res = this.resolveDict(resources) || {}
    const fontDict = this.resolveDict(res.Font) || {}
    for (const [name, ref] of Object.entries(fontDict)) {
      const fd = this.resolveDict(ref)
      if (!fd) continue
      const font = { subtype: String(fd.Subtype || ''), base: String(fd.BaseFont || ''), toUnicode: null, encoding: null, encodingName: null }
      if (fd.ToUnicode) {
        const num = isRef(fd.ToUnicode) ? fd.ToUnicode.__ref : null
        if (num !== null) {
          const bytes = this.streamBytes(num)
          if (bytes && bytes.length) font.toUnicode = parseToUnicode(Buffer.from(bytes).toString('latin1'))
        }
      }
      const enc = this.resolveDict(fd.Encoding)
      if (enc && enc.Differences) font.encoding = buildDifferences(this.resolve(enc.Differences))
      else if (typeof fd.Encoding === 'string') font.encodingName = fd.Encoding
      if (font.subtype === 'Type0' && !font.toUnicode) {
        // try descendant /ToUnicode? ToUnicode lives on the Type0 font itself
        const desc = this.resolveDict(Array.isArray(fd.DescendantFonts) ? fd.DescendantFonts[0] : fd.DescendantFonts)
        void desc
      }
      out.set(name, font)
    }
    // Font arrays: /Font << /F1 [ ... ] >> is invalid; single fonts only.
    return out
  }
}

// ---------------- PDF value parsing ----------------

function normalizeFilters(f) {
  if (!f) return []
  if (typeof f === 'string') return [f]
  if (Array.isArray(f)) return f.map(x => String(x))
  return []
}

function isRef(v) { return v && typeof v === 'object' && typeof v.__ref === 'number' }

function unescapePDFString(raw) {
  const out = []
  for (let i = 0; i < raw.length; i++) {
    const c = raw[i]
    if (c !== '\\') { out.push(c); continue }
    const n = raw[++i]
    if (n === undefined) break
    switch (n) {
      case 'n': out.push('\n'); break
      case 'r': out.push('\r'); break
      case 't': out.push('\t'); break
      case 'b': out.push('\b'); break
      case 'f': out.push('\f'); break
      case '(': out.push('('); break
      case ')': out.push(')'); break
      case '\\': out.push('\\'); break
      case '\n': break
      case '\r': if (raw[i + 1] === '\n') i++; break
      default: {
        if (/[0-7]/.test(n)) {
          let oct = n
          while (oct.length < 3 && /[0-7]/.test(raw[i + 1])) oct += raw[++i]
          out.push(String.fromCharCode(parseInt(oct, 8) & 0xff))
        } else out.push(n)
      }
    }
  }
  return out.join('')
}

const DELIMS = new Set(' \t\r\n\f()<>[]{}/%'.split(''))

/** Decrypt one byte string in place of parsing (names are never encrypted). */
function decryptPdfString(str, crypt) {
  if (!crypt || !str) return str
  try {
    return crypt.decryptString(Buffer.from(str, 'latin1')).toString('latin1')
  } catch {
    return str
  }
}

function skipWs(s, i) {
  while (i.p < s.length && /[\s\r\n\f\t]/.test(s[i.p])) i.p++
}

function parseValue(s, i, objs = null, crypt = null) {
  skipWs(s, i)
  const c = s[i.p]
  if (c === undefined) return { v: null }
  if (c === '<') {
    if (s[i.p + 1] === '<') {
      i.p += 2
      const dict = {}
      for (;;) {
        skipWs(s, i)
        if (s[i.p] === '>' && s[i.p + 1] === '>') { i.p += 2; break }
        if (i.p >= s.length) break
        if (s[i.p] !== '/') { i.p++; continue }
        const key = parseName(s, i)
        skipWs(s, i)
        const save = i.p
        const val = parseValue(s, i, objs, crypt)
        if (i.p === save) { i.p++ }
        if (key) dict[key] = val.v
      }
      return { v: dict }
    }
    const end = s.indexOf('>', i.p + 1)
    const raw = s.slice(i.p + 1, end < 0 ? s.length : end)
    i.p = end < 0 ? s.length : end + 1
    const clean = raw.replace(/[^0-9a-fA-F]/g, '')
    let out = ''
    for (let k = 0; k + 1 < clean.length; k += 2) out += String.fromCharCode(parseInt(clean.slice(k, k + 2), 16))
    return { v: decryptPdfString(out, crypt) }
  }
  if (c === '(') {
    let depth = 1
    let p = i.p + 1
    let raw = ''
    while (p < s.length && depth > 0) {
      const ch = s[p]
      if (ch === '\\') { raw += ch + (s[p + 1] ?? ''); p += 2; continue }
      if (ch === '(') depth++
      else if (ch === ')') { depth--; if (!depth) break }
      raw += ch
      p++
    }
    i.p = p + 1
    return { v: decryptPdfString(unescapePDFString(raw), crypt) }
  }
  if (c === '[') {
    i.p++
    const arr = []
    for (;;) {
      skipWs(s, i)
      if (s[i.p] === ']') { i.p++; break }
      if (i.p >= s.length) break
      const before = i.p
      const val = parseValue(s, i, objs, crypt)
      if (i.p === before) { i.p++; continue }
      arr.push(val.v)
    }
    return { v: arr }
  }
  if (c === '/') return { v: parseName(s, i) }
  if (c === 't' && s.startsWith('true', i.p)) { i.p += 4; return { v: true } }
  if (c === 'f' && s.startsWith('false', i.p)) { i.p += 5; return { v: false } }
  if (c === 'n' && s.startsWith('null', i.p)) { i.p += 4; return { v: null } }
  const m = /^(-?\d+(?:\.\d+)?|\.\d+|-?\.\d+)/.exec(s.slice(i.p))
  if (m) {
    i.p += m[0].length
    const num = Number(m[0])
    const save = i.p
    skipWs(s, i)
    const gm = /^(\d{1,10})\s+R\b/.exec(s.slice(i.p))
    if (gm && Number.isInteger(num) && num >= 0) { i.p += gm[0].length; return { v: { __ref: num, gen: Number(gm[1]) } } }
    i.p = save
    return { v: num }
  }
  const w = /^[A-Za-z][A-Za-z0-9.#_-]*/.exec(s.slice(i.p))
  if (w) { i.p += w[0].length; return { v: w[0] } }
  i.p++
  return { v: c }
}

function parseName(s, i) {
  if (s[i.p] !== '/') return null
  let p = i.p + 1
  let name = ''
  while (p < s.length) {
    const ch = s[p]
    if (DELIMS.has(ch)) break
    if (ch === '#' && /[0-9a-fA-F]{2}/.test(s.slice(p + 1, p + 3))) { name += String.fromCharCode(parseInt(s.slice(p + 1, p + 3), 16)); p += 3; continue }
    name += ch
    p++
  }
  i.p = p
  return name
}

// ---------------- decoding helpers ----------------

function buildDifferences(arr) {
  const table = new Array(256).fill(null)
  let code = 0
  if (!Array.isArray(arr)) return table
  for (const el of arr) {
    if (typeof el === 'number') code = el
    else if (typeof el === 'string') { table[code & 0xff] = glyphToUnicode(el); code++ }
  }
  return table
}

const GLYPH_FIXED = {
  space: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&',
  quotesingle: "'", parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-',
  period: '.', slash: '/', colon: ':', semicolon: ';', less: '<', equal: '=', greater: '>', question: '?',
  at: '@', bracketleft: '[', backslash: '\\', bracketright: ']', asciicircum: '^', underscore: '_',
  grave: '`', braceleft: '{', bar: '|', braceright: '}', asciitilde: '~', bullet: '•', endash: '–',
  emdash: '—', quoteleft: '\u2018', quoteright: '\u2019', quotedblleft: '\u201c', quotedblright: '\u201d',
  fi: 'fi', fl: 'fl', ffl: 'ffl', ffi: 'ffi', ellipsis: '…', trademark: '™', Euro: '€', degree: '°',
  plusminus: '±', twosuperior: '²', threesuperior: '³', mu: 'µ', paragraph: '¶', registered: '®',
  copyright: '©', onehalf: '½', minus: '−', dagger: '†',
}

function glyphToUnicode(name) {
  if (GLYPH_FIXED[name]) return GLYPH_FIXED[name]
  if (/^g[a-f0-9]{2,6}$/i.test(name)) return String.fromCodePoint(parseInt(name.slice(1), 16))
  if (/^(cid|g\d)/i.test(name) || name === '.notdef') return ''
  if (/^uni[0-9a-f]{4}$/i.test(name)) return String.fromCharCode(parseInt(name.slice(3), 16))
  const wordNum = { zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9' }
  if (wordNum[name.toLowerCase()]) return wordNum[name.toLowerCase()]
  if (name.length === 1) return name
  const m = /^([A-Za-z])(er|lf|rt|o|s)$/.exec(name) // a-acute style ignored
  if (m) return m[1]
  return ''
}

function parseToUnicode(text) {
  const map = new Map()
  const sectionRe = /begin(bfchar|bfrange|cmchar|cm)?\b[\s\S]*?end\1?\b/g
  void sectionRe
  const bfcharRe = /beginbfchar([\s\S]*?)endbfchar/g
  const bfrangeRe = /beginbfrange([\s\S]*?)endbfrange/g
  const hexes = /<([0-9a-fA-F\s]*)>/g
  const codePointFromHex = hex => {
    const clean = hex.replace(/\s/g, '')
    if (!clean) return ''
    if (clean.length <= 4) {
      const cp = parseInt(clean, 16)
      return cp === 0 ? '' : String.fromCodePoint(cp)
    }
    let out = ''
    for (let i = 0; i + 1 < clean.length; i += 4) out += String.fromCharCode(parseInt(clean.slice(i, i + 4), 16))
    return out
  }
  let sm
  while ((sm = bfcharRe.exec(text))) {
    const items = [...sm[1].matchAll(hexes)]
    for (let k = 0; k + 1 < items.length; k += 2) {
      const src = parseInt(items[k][1].replace(/\s/g, ''), 16)
      map.set(src, codePointFromHex(items[k + 1][1]))
    }
  }
  while ((sm = bfrangeRe.exec(text))) {
    const body = sm[1]
    const lineRe = /<([0-9a-fA-F\s]+)>\s*<([0-9a-fA-F\s]+)>\s*(<([0-9a-fA-F\s]+)>|\[([^\]]*)\])/g
    let lm
    while ((lm = lineRe.exec(body))) {
      const lo = parseInt(lm[1].replace(/\s/g, ''), 16)
      const hi = parseInt(lm[2].replace(/\s/g, ''), 16)
      if (hi - lo > 65535) continue
      if (lm[4] !== undefined) {
        const first = codePointFromHex(lm[4])
        const base = first ? first.codePointAt(0) : 0
        for (let c = lo; c <= hi; c++) map.set(c, String.fromCodePoint(base + (c - lo)))
      } else if (lm[5]) {
        const arr = [...lm[5].matchAll(/<([0-9a-fA-F\s]+)>/g)].map(x => codePointFromHex(x[1]))
        for (let c = lo; c <= hi && c - lo < arr.length; c++) map.set(c, arr[c - lo])
      }
    }
  }
  return map
}

function applyPredictor(bytes, parm) {
  const pred = Number(parm.Predictor || 1)
  if (pred === 1 || !bytes.length) return bytes
  const colors = Math.max(1, Number(parm.Colors || 1))
  const bpc = Math.min(8, Math.max(1, Number(parm.BitsPerComponent || 8)))
  const columns = Math.max(1, Number(parm.Columns || 1))
  const rowLen = Math.ceil((colors * columns * bpc) / 8)
  const src = Buffer.from(bytes)
  if (pred === 2) {
    const out = Buffer.from(src)
    for (let r = 1; r * rowLen < out.length; r++) {
      for (let c = 0; c < rowLen; c++) out[r * rowLen + c] = (out[r * rowLen + c] + out[(r - 1) * rowLen + c]) & 0xff
    }
    return new Uint8Array(out)
  }
  const out = []
  let si = 0
  let prev = Buffer.alloc(rowLen)
  while (si < src.length) {
    const ft = src[si++]
    if (si + rowLen > src.length) break
    const cur = src.slice(si, si + rowLen)
    si += rowLen
    const res = Buffer.from(cur)
    if (ft === 1) for (let c = colors; c < rowLen; c++) res[c] = (res[c] + res[c - colors]) & 0xff
    else if (ft === 2) for (let c = 0; c < rowLen; c++) res[c] = (res[c] + prev[c]) & 0xff
    else if (ft === 3) for (let c = 0; c < rowLen; c++) res[c] = (res[c] + Math.floor(((c >= colors ? res[c - colors] : 0) + prev[c]) / 2)) & 0xff
    else if (ft === 4) {
      for (let c = 0; c < rowLen; c++) {
        const a = c >= colors ? res[c - colors] : 0
        const b = prev[c]
        const cc = c >= colors ? prev[c - colors] : 0
        const p = a + b - cc
        const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - cc)
        res[c] = (res[c] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : cc)) & 0xff
      }
    }
    out.push(res)
    prev = res
  }
  return new Uint8Array(Buffer.concat(out))
}

function lzwDecode(bytes) {
  const out = []
  let bitBuf = 0
  let bitCnt = 0
  let codeWidth = 9
  let dict = new Map()
  let next = 258
  const reset = () => { dict = new Map(); for (let i = 0; i < 256; i++) dict.set(i, [i]); codeWidth = 9; next = 258 }
  reset()
  let prevCode = null
  for (let i = 0; i <= bytes.length; i++) {
    if (i < bytes.length) { bitBuf = (bitBuf << 8) | bytes[i]; bitCnt += 8 }
    while (bitCnt >= codeWidth || (i === bytes.length && bitCnt > 0)) {
      if (bitCnt < codeWidth) break
      const code = (bitBuf >> (bitCnt - codeWidth)) & ((1 << codeWidth) - 1)
      bitCnt -= codeWidth
      if (code === 256) { reset(); prevCode = null; continue }
      if (code === 257) return Uint8Array.from(out)
      let seq
      if (dict.has(code)) seq = dict.get(code)
      else if (prevCode !== null && dict.has(prevCode)) seq = [...dict.get(prevCode), dict.get(prevCode)[0]]
      else return Uint8Array.from(out)
      out.push(...seq)
      if (prevCode !== null) {
        dict.set(next++, seq[0] !== undefined ? [...seq.slice(0, seq.length - 1), seq[0]] : seq)
        if (next + 1 >= (1 << codeWidth) && codeWidth < 12) codeWidth++
      }
      prevCode = code
    }
  }
  return Uint8Array.from(out)
}

function runLengthDecode(bytes) {
  const out = []
  for (let i = 0; i < bytes.length;) {
    const len = bytes[i++]
    if (len === 128) break
    if (len < 128) { for (let k = 0; k <= len && i < bytes.length; k++) out.push(bytes[i++]) }
    else { const b = bytes[i++]; for (let k = 0; k < 257 - len; k++) out.push(b) }
  }
  return Uint8Array.from(out)
}

function asciiHexDecode(bytes) {
  const s = Buffer.from(bytes).toString('latin1').replace(/[^0-9a-fA-F]/g, '')
  const out = new Uint8Array(Math.floor(s.length / 2))
  for (let i = 0; i < out.length; i++) out[i] = parseInt(s.slice(i * 2, i * 2 + 2), 16)
  return out
}

function ascii85Decode(bytes) {
  const s = Buffer.from(bytes).toString('latin1').replace(/\s/g, '').replace(/^<~/, '').split('~>')[0]
  const out = []
  for (let i = 0; i < s.length;) {
    if (s[i] === 'z') { out.push(0, 0, 0, 0); i++; continue }
    let acc = 0
    let n = 0
    for (; n < 5 && i + n < s.length; n++) acc = acc * 85 + (s.charCodeAt(i + n) - 33)
    for (let k = 0; k < n - 1; k++) out.push((acc >>> (24 - 8 * k)) & 0xff)
    i += n
  }
  return Uint8Array.from(out)
}

// ---------------- content stream tokenizing ----------------

function tokenizeContent(s) {
  const tokens = []
  let i = 0
  const n = s.length
  while (i < n) {
    const c = s[i]
    if (/\s/.test(c)) { i++; continue }
    if (c === '%') { while (i < n && s[i] !== '\n') i++; continue }
    if (c === '(') {
      let depth = 1
      let p = i + 1
      while (p < n && depth > 0) {
        if (s[p] === '\\') { p += 2; continue }
        if (s[p] === '(') depth++
        else if (s[p] === ')') depth--
        p++
      }
      tokens.push({ kind: 'string', raw: s.slice(i + 1, p - 1 >= i + 1 ? p - 1 : i + 1) })
      i = p
      continue
    }
    if (c === '<') {
      if (s[i + 1] === '<') {
        const end = s.indexOf('>>', i)
        i = end < 0 ? n : end + 2
        continue // dict operand — ignore
      }
      const end = s.indexOf('>', i)
      tokens.push({ kind: 'hex', raw: s.slice(i + 1, end < 0 ? n : end) })
      i = end < 0 ? n : end + 1
      continue
    }
    if (c === '[') { tokens.push({ kind: 'arr-open' }); i++; continue }
    if (c === ']') { tokens.push({ kind: 'arr-close' }); i++; continue }
    if (c === '{' || c === '}') { i++; continue }
    if (c === '/') {
      let p = i + 1
      while (p < n && !DELIMS.has(s[p])) {
        if (s[p] === '#' && /[0-9a-fA-F]{2}/.test(s.slice(p + 1, p + 3))) { p += 3; continue }
        p++
      }
      tokens.push({ kind: 'name', raw: s.slice(i + 1, p).replace(/#([0-9a-fA-F]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))) })
      i = p
      continue
    }
    let p = i
    while (p < n && !DELIMS.has(s[p]) && s[p] !== '{' && s[p] !== '}') p++
    if (p === i) { i++; continue }
    const raw = s.slice(i, p)
    if (/^-?\d+(?:\.\d+)?$/.test(raw)) tokens.push({ kind: 'num', num: Number(raw) })
    else tokens.push({ kind: 'op', raw })
    i = p
  }
  return tokens
}

const OPS_TEXT = new Set(['BT', 'ET', 'Tf', 'Td', 'TD', 'Tm', 'T*', 'TL', 'Tj', 'TJ', "'", '"'])
// 文本状态算子：只改状态、不进 `runs`，但在 PDF 里合法且普遍存在。
// 旧版把它们按"未知算子"处理（清操作数栈），状态于是永远读不到（`Tr` 尤其要命：
// `3 Tr` 是**不可见**文字，渲染模式不生效就会把水印/隐藏层当正文）。
const OPS_TEXT_STATE = new Set(['Tz', 'Tc', 'Tw', 'Ts', 'Tr'])

/**
 * PDF 矩阵乘法（`[a b c d e f]` ＝ `[[a b 0][c d 0][e f 1]]`，点按行向量 `p' = p·M`）。
 *
 * 语义**固定为**：`apply(matrixMul(a, b), p) === apply(a, apply(b, p))` —— 即"先 b 后 a"。
 * 命名成 `matrixMul` 并导出，是为了让"合成顺序"这条契约能被直接钉住（见 test.mjs 的
 * `矩阵：` 用例）；`extractPageText` 内部仍用局部别名 `mul`。
 *
 * 旧写法 e/f 两行有三个独立错误（`pdf.js:663-667`）：
 *   ① 读了越界的 `b[6]`（PDF 矩阵只有 0–5，读出来是 `undefined`）
 *   ② 该用 `a[4]/a[5]` 的地方写成了 `b[4]/b[5]`
 *   ③ 因为 `0 * undefined === NaN`，**连纯平移都炸**（`mul(ident, [1,0,0,1,5,7])` 的
 *      e/f 都是 NaN）→ 页内每个 run 的 y 都是 NaN → `Math.round(NaN / 2.5)` 是合法
 *      Map key → 整页所有 run 归进同一个桶 → **每页塌成一行**。
 */
export function matrixMul(a, b) {
  return [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
  ]
}

function extractPageText(pdf, page) {
  const fonts = pdf.fontsOf(page.attrs.Resources)
  const contentsVal = pdf.resolve(page.dict.Contents)
  const refs = []
  if (isRef(page.dict.Contents)) refs.push(page.dict.Contents.__ref)
  else if (Array.isArray(contentsVal)) for (const el of contentsVal) { if (isRef(el)) refs.push(el.__ref) }
  const runs = []
  for (const num of refs) {
    const bytes = pdf.streamBytes(num)
    if (!bytes || !bytes.length) continue
    const tokens = tokenizeContent(Buffer.from(bytes).toString('latin1'))
    let stack = []
    let arrMode = null
    let ctm = [1, 0, 0, 1, 0, 0]
    const gsStack = []              // q/Q 的图形状态栈（PDF 规范：图形状态含 CTM）
    let tlm = null                  // 文本行矩阵 Tlm —— **文本空间**，不含 CTM
    let leading = 0                 // TL
    let charSpacing = 0             // Tc（千分之一 em 之外的单位：直接是文本空间单位）
    let wordSpacing = 0             // Tw
    let hScale = 100                // Tz（百分数）
    let rise = 0                    // Ts
    let renderMode = 0              // Tr
    let curFont = null
    let fontSize = 10
    /**
     * 矩阵乘法别名 —— 语义见模块级 `matrixMul` 的注释（"先 b 后 a"）。
     */
    const mul = matrixMul
    const IDENT = [1, 0, 0, 1, 0, 0]
    /**
     * 文本渲染矩阵 `Trm = Tfs × Tlm × CTM`（规范 9.4.4）。
     * "先 Tfs、再 Tlm、再 CTM" 在"先 b 后 a"的 `mul` 下就是 `mul(ctm, mul(tlm, Tfs))`。
     * 旧版把 CTM 预先乘进 `tlm`（`mul(ctm, Tm)`），看似等价，但 `T*` / `'` / `"`
     * 之后再合成 `leading` 位移时就会在**设备空间**里平移（CTM 带旋转/缩放时错位）。
     */
    const trm = () => {
      const tfs = [fontSize * (hScale / 100), 0, 0, fontSize, 0, rise]
      return mul(ctm, mul(tlm || IDENT, tfs))
    }
    /** 设备空间的有效字号：CTM/Tm 的缩放必须算进去（否则宽度估算会差几十倍）。 */
    const effSize = trmMatrix => Math.hypot(trmMatrix[2], trmMatrix[3]) || Math.abs(fontSize) || 10
    /** `3 Tr` 是**不可见**文字（OCR 层）；`7 Tr` 只加裁剪路径，同样不产生可见正文。 */
    const invisibleText = () => renderMode === 3 || renderMode === 7
    const strOf = tok => tok.raw
    let skipUntilEI = false
    for (const t of tokens) {
      if (skipUntilEI) {
        if (t.kind === 'op' && t.raw === 'EI') skipUntilEI = false
        continue
      }
      if (t.kind === 'num' || t.kind === 'string' || t.kind === 'hex' || t.kind === 'name') {
        (arrMode ?? stack).push(t)
        continue
      }
      if (t.kind === 'arr-open') { arrMode = []; continue }
      if (t.kind === 'arr-close') {
        const items = arrMode || []
        arrMode = null
        stack.push({ kind: 'array', items })
        continue
      }
      if (t.kind !== 'op') continue
      const op = t.raw
      if (!OPS_TEXT.has(op) && !OPS_TEXT_STATE.has(op)) {
        if (op === 'BI') skipUntilEI = true
        if (op === 'q' || op === 'Q' || op === 'cm') {
          if (op === 'cm') {
            // 规范 8.4.4：`CTM_new = M_cm × CTM_old` —— 即"先 cm、后旧 CTM"。
            // 在"先 b 后 a"的 mul 下是 `mul(ctm, a)`；旧版写成 `mul(a, ctm)` 把顺序
            // 弄反了（带旋转/缩放时错位）。
            const a = stack.slice(-6).map(x => x.num)
            if (a.length === 6 && a.every(Number.isFinite)) ctm = mul(ctm, a)
          }
          // q/Q 只保存/恢复**图形状态**（含 CTM）；文本状态参数（Tf/TL/Tz/…）
          // 按规范不随 q/Q 回滚，所以这里只压 CTM。旧版 q/Q 只清操作数栈，
          // CTM 整页单调漂移永不复位 → 后续所有页元素坐标全错。
          if (op === 'q') gsStack.push(ctm.slice(0, 6))
          else if (op === 'Q') { const prev = gsStack.pop(); if (prev) ctm = prev }
          stack = []
          continue
        }
        stack = []
        continue
      }
      switch (op) {
        case 'BT': stack = []; tlm = IDENT.slice(); break
        case 'ET': stack = []; tlm = null; break
        case 'Tf': {
          const nameTok = stack[stack.length - 2]
          const sizeTok = stack[stack.length - 1]
          curFont = fonts.get(String(nameTok?.raw ?? '')) || null
          fontSize = Number(sizeTok?.num ?? 10) || 10
          stack = []
          break
        }
        case 'Td': case 'TD': {
          const x = Number(stack[stack.length - 2]?.num ?? 0)
          const y = Number(stack[stack.length - 1]?.num ?? 0)
          // 规范 9.4.2：`Tlm_new = T_translate × Tlm_old` —— 位移先做、再套上旧的 Tlm，
          // 这样 (x, y) 就是"沿当前文本行的方向"偏移（行被旋转时也对）。
          // 在 `mul(A,B)`＝"先 b 后 a"的语义下，`apply(T_translate, 再 Tlm_old)`
          // 必须写成 `mul(Tlm_old, T_translate)`。
          // 旧版写的是 `tlm = mul(ctm, Td)`：既**丢掉上一个 Tlm**（连续 Td 的多行文本
          // 会全部落在同一点），又把位移预先揉进了 CTM。
          tlm = mul(tlm || IDENT, [1, 0, 0, 1, x, y])
          if (op === 'TD') leading = -y
          stack = []
          break
        }
        case 'Tm': {
          const a = stack.slice(-6).map(x => x.num)
          if (a.length === 6 && a.every(Number.isFinite)) tlm = a
          stack = []
          break
        }
        case 'Tz': { hScale = Number(stack[stack.length - 1]?.num ?? 100) || 100; stack = []; break }
        case 'Tc': { charSpacing = Number(stack[stack.length - 1]?.num ?? 0) || 0; stack = []; break }
        case 'Tw': { wordSpacing = Number(stack[stack.length - 1]?.num ?? 0) || 0; stack = []; break }
        case 'Ts': { rise = Number(stack[stack.length - 1]?.num ?? 0) || 0; stack = []; break }
        case 'Tr': { renderMode = Math.trunc(Number(stack[stack.length - 1]?.num ?? 0)) || 0; stack = []; break }
        case 'TL': { leading = Number(stack[stack.length - 1]?.num ?? 0); stack = []; break }
        case 'T*': { tlm = mul(tlm || IDENT, [1, 0, 0, 1, 0, -leading]); stack = []; break }
        case "'": case '"': { tlm = mul(tlm || IDENT, [1, 0, 0, 1, 0, -leading]) }
        // fallthrough to Tj behaviour
        case 'Tj': {
          const st = stack[stack.length - 1]
          if (st && (st.kind === 'string' || st.kind === 'hex')) {
            const m = trm()
            if (!invisibleText()) pushRun(runs, m, decodePdfString(strOf(st), curFont, st.kind === 'hex'), effSize(m), fontSize)
          }
          stack = []
          break
        }
        case 'TJ': {
          const arr = stack[stack.length - 1]
          void arr
          let text = ''
          for (const el of (stack.find(x => x.kind === 'array')?.items ?? [])) {
            if (el.kind === 'string' || el.kind === 'hex') text += decodePdfString(strOf(el), curFont, el.kind === 'hex')
            // 词距判据：`Tc/Tw` 是真实存在的调节量，必须一起算进去，否则同一段
            // 文字在不同 `Tw` 下会被切成不同的词。
            else if (el.kind === 'num' && (-el.num / 1000) * (fontSize || 10) + charSpacing + wordSpacing >= 0.16 * (fontSize || 10)) text += ' '
          }
          const m = trm()
          if (!invisibleText()) pushRun(runs, m, text, effSize(m), fontSize)
          stack = []
          break
        }
        default: stack = []
      }
    }
  }
  return runs
}

/**
 * @param matrix 设备空间的文本渲染矩阵 `Trm`
 * @param size   设备空间的**有效**字号（= `|Trm|` 的纵向量长），不是裸 `Tf` 值 ——
 *               旧版存裸 `Tf`（某样本 209），而真实有效字号是 10.45pt，
 *               `approximateWidth` 于是把宽度算大 20 倍，run 之间的间距判断全错。
 * @param rawSize 裸 `Tf` 字号，保留给诊断（`stats`/调试），不参与几何计算。
 */
function pushRun(runs, matrix, text, size, rawSize = size) {
  if (!text) return
  runs.push({ x: matrix[4], y: matrix[5], size: Math.abs(size) || 10, text, rawSize: Math.abs(rawSize) || 10 })
}

const WINANSI_EXTRA = { 0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ', 0x8e: 'Ž', 0x91: '\u2018', 0x92: '\u2019', 0x93: '\u201c', 0x94: '\u201d', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›', 0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ' }

function hexToBytes(raw) {
  const clean = raw.replace(/[^0-9a-fA-F]/g, '')
  const out = []
  for (let i = 0; i + 1 < clean.length; i += 2) out.push(parseInt(clean.slice(i, i + 2), 16))
  return out
}

function decodePdfString(raw, font, isHex) {
  const bytes = isHex ? hexToBytes(raw) : [...unescapePDFString(raw)].map(ch => ch.charCodeAt(0) & 0xff)
  if (font && font.subtype === 'Type0') {
    const out = []
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      const code = (bytes[i] << 8) | bytes[i + 1]
      if (font.toUnicode?.has(code)) { const u = font.toUnicode.get(code); out.push(typeof u === 'string' ? u : String.fromCharCode(u || 0)); continue }
      if (code === 0) { out.push(' '); continue }
      if (code >= 0xfe00) { out.push(''); continue }
      if (code > 0x2000 || (code >= 0x8140 && code <= 0xfe4e)) out.push(String.fromCharCode(code))
      else out.push(String.fromCharCode(code))
    }
    return out.join('')
  }
  let out = ''
  for (const b of bytes) {
    if (font?.encoding && font.encoding[b] != null) { out += font.encoding[b] || ''; continue }
    if (font?.toUnicode?.has(b)) { const u = font.toUnicode.get(b); out += typeof u === 'string' ? u : String.fromCharCode(u || 0); continue }
    if (b >= 0x80) out += (WINANSI_EXTRA[b] ?? String.fromCharCode(b))
    else out += String.fromCharCode(b)
  }
  return out
}

/** CJK 及全角标点：判断"这两个 run 之间该不该补空格"。 */
function isCjkLike(ch) {
  if (!ch) return false
  const cp = ch.codePointAt(0)
  return (cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef)
}

function runsToLines(runs, page) {
  if (!runs.length) return []
  const media = (page.attrs.MediaBox && Array.isArray(page.attrs.MediaBox) ? page.attrs.MediaBox.map(Number) : [0, 0, 595, 842])
  void media
  const sizes = runs.map(r => r.size).filter(s => s > 0).sort((a, b) => a - b)
  const bodySize = sizes.length ? sizes[Math.floor(sizes.length / 2)] : 10
  const groups = new Map()
  for (const r of runs) {
    const key = Math.round(r.y / 2.5)
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key).push(r)
  }
  const lines = []
  for (const [yKey, items] of groups) {
    items.sort((a, b) => a.x - b.x)
    let text = ''
    let lastEnd = null
    let maxSize = 0
    for (const it of items) {
      // 同一行里 run 之间的空隙：只有**两侧都不是 CJK** 时才补空格。
      // 中文排版 run 与 run 之间本来就不写空格，旧规则（只对恰好以 `一` 开头的
      // run 特判）会把中文行的每个 run 之间都塞一个空格，凭空造出字符。
      if (lastEnd !== null && it.x - lastEnd > Math.max(1.2, it.size * 0.6)
        && !text.endsWith(' ') && !it.text.startsWith(' ')
        && !isCjkLike(text[text.length - 1]) && !isCjkLike(it.text[0])) text += ' '
      text += it.text
      lastEnd = it.x + approximateWidth(it.text, it.size)
      maxSize = Math.max(maxSize, it.size)
    }
    const y = yKey * 2.5
    if (text.trim()) lines.push({ y, text, size: maxSize })
  }
  lines.sort((a, b) => b.y - a.y)
  return { lines, bodySize }
}

function approximateWidth(text, size) {
  let w = 0
  for (const ch of text) w += (ch.codePointAt(0) > 0x2e80 ? 1 : 0.52) * size
  return w
}

function decodeInfoString(v) {
  if (typeof v !== 'string') return undefined
  if (v.charCodeAt(0) === 0xfe && v.charCodeAt(1) === 0xff) {
    const bytes = Buffer.from(v.slice(2), 'latin1')
    return new TextDecoder('utf-16be').decode(bytes)
  }
  // PDFDocEncoding/WinAnsi latin1
  let out = ''
  for (const ch of v) {
    const b = ch.charCodeAt(0)
    if (b >= 0x80) out += (WINANSI_EXTRA[b] ?? String.fromCharCode(b))
    else out += ch
  }
  return out
}

export function readPdfText(buf) {
  const pdf = new PdfFile(buf)
  if (pdf.isEncrypted() && !pdf.enc?.supported) {
    throw new Error(`pdf 已加密，无法读取：${pdf.enc?.note ?? '未知加密方式'}`)
  }
  const meta = {}
  const info = pdf.infoDict()
  if (info) {
    for (const [key, name] of [['Title', 'title'], ['Author', 'author'], ['Creator', 'creator'], ['Producer', 'producer'], ['CreationDate', 'created'], ['ModDate', 'modified']]) {
      let v = info[key]
      if (isRef(v)) v = pdf.value(v.__ref)
      const s = decodeInfoString(v)
      if (s) meta[name] = s
    }
  }
  const pages = pdf.pages()
  const sections = []
  let failed = 0
  pages.forEach((pg, idx) => {
    let text = ''
    try {
      const runs = extractPageText(pdf, pg)
      const { lines, bodySize } = runsToLines(runs, pg)
      const out = []
      let prevY = null
      for (const ln of lines) {
        const t = collapseSpaces(ln.text)
        if (prevY !== null && prevY - ln.y > ln.size * 1.9 && prevY - ln.y < 400) out.push('')
        if (ln.size > bodySize * 1.3 && t.length > 0 && t.length < 120) out.push(`# ${t}`)
        else out.push(t)
        prevY = ln.y
      }
      text = out.join('\n')
    } catch { failed++; text = '' }
    sections.push({ page: idx + 1, text })
  })
  const textFound = sections.some(s => s.text.replace(/\s/g, '').length)
  return { meta, pages: sections.length, sections, textFound }
}

function collapseSpaces(s) {
  return s.replace(/[ \t]{2,}/g, ' ').replace(/^\s+|\s+$/g, '')
}

// ============================ writing ============================

const WINANSI_MAP = new Map(Object.entries({ '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '\u2018': 0x91, '\u2019': 0x92, '\u201c': 0x93, '\u201d': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f }))

function winAnsiByte(ch) {
  const code = ch.codePointAt(0)
  if (code <= 255) return code
  if (WINANSI_MAP.has(ch)) return WINANSI_MAP.get(ch)
  return 0x3f
}

const HELV_W = {}
for (let c = 32; c < 127; c++) HELV_W[String.fromCharCode(c)] = 556
for (const c of ['i', 'j', 'l', '!', '|', '.', ',', ':', ';', "'", '`']) HELV_W[c] = 222
for (const c of ['f', '(', ')', '[', ']', '-', '/', '\\', '?', '"']) HELV_W[c] = 333
for (const c of ['r', 't', '{', '}', '‘', '’']) HELV_W[c] = 278
for (const c of ['m', 'w']) HELV_W[c] = 833
for (const c of ['M', 'W']) HELV_W[c] = 944
for (const c of '0123456789') HELV_W[c] = 556
for (const c of 'ABCDEFGHIJKLMNOPQRSTUVWXZ') HELV_W[c] = 700
for (const c of 'EFGHKLRSTXZ') HELV_W[c] = 667
for (const c of 'abcdeghknopqsuvxyz') HELV_W[c] = 556
for (const c of 'fhilmnrstuvw') HELV_W[c] = 500
HELV_W[' '] = 278
HELV_W['%'] = 889
void 'IJK'

function widthOf(text, size, font = 'F1') {
  let w = 0
  for (const seg of segmentText(text, font)) w += seg.em * size
  return w
}

// —— 阶段一（中文 PDF 嵌字体）：写出期的布局上下文 ——
// writePdf 是同步函数，模块级上下文不会被并发写穿；widthOf / segmentText /
// wrap 的断词都从这里取 chain（内嵌字体链）与 cjk（旧的文档级判据，仅作
// "内嵌链也没有该字形"时的 STSong 回落开关）。
let pdfLayout = null

/** 格式类字符（ZWJ/ZWNBSP/变体选择符等）：无字形语义，内嵌链下按零宽丢弃。 */
function isFormatChar(cp) {
  return cp === 0x200b || cp === 0x200c || cp === 0x200d || cp === 0x2060
    || (cp >= 0xfe00 && cp <= 0xfe0f) || (cp >= 0xe0100 && cp <= 0xe01ef)
}

/** base-14 字宽（em 小数）——沿用旧 widthOf 的查表与字体规则，逐字符口径不变。 */
function latinEm(ch, font) {
  if (font === 'F3') return 0.6
  if (font === 'F2') return ((HELV_W[ch] ?? 556) * 1.06) / 1000
  return (HELV_W[ch] ?? 556) / 1000
}

/**
 * 把一段文字切成"同一字体、同一编码方式"的连续段（阶段一 drawText/widthOf
 * 共用同一套分段，保证**测量与绘制逐段一致**）：
 *   latin — base-14（F1/F2/F3，WinAnsi 十六进制）
 *   embed — 内嵌链命中的字形（Identity-H，2 字节码 = GID）
 *   gb    — 回落 STSong-Light（UniGB-UCS2-H，旧路径；含内嵌链 miss 且 cjk 的字符）
 *   zero  — 格式字符，零宽且不落内容流（不参与提取）
 * 每段带 em（该段总宽的 em 系数）；widthOf = Σ em × size。
 */
function segmentText(text, fontName) {
  const { chain, cjk } = pdfLayout || {}
  const segs = []
  let cur = null
  for (const ch of text) {
    let k
    let f = null
    if (!chain) {
      // 无内嵌链（DSH_OFFICE_PDF_EMBED_CJK=0 / 关闭态）：旧语义逐字等价——
      // 文档含 CJK 时非 WinAnsi 字符走 gb，其余走 base-14。
      k = (!winAnsiOk(ch) && cjk) ? 'gb' : 'latin'
    } else if (isFormatChar(ch.codePointAt(0))) {
      k = 'zero'
    } else {
      const c = chain.classify(ch)
      if (c.k === 'latin') k = 'latin'
      else if (c.k === 'embed') { k = 'embed'; f = c.f }
      else k = cjk ? 'gb' : 'latin'   // miss：内嵌链没有该字形
    }
    let em
    if (k === 'embed') em = chain.emWidth(f, chain.classify(ch).g)
    else if (k === 'gb') em = ch.codePointAt(0) > 0x2e80 ? 1 : latinEm(ch, fontName)
    else if (k === 'zero') em = 0
    else em = latinEm(ch, fontName)
    if (cur && cur.k === k && cur.f === f && cur.font === fontName) {
      cur.text += ch
      cur.em += em
    } else {
      if (cur) segs.push(cur)
      cur = { k, f, font: fontName, text: ch, em }
    }
  }
  if (cur) segs.push(cur)
  return segs
}

export function needsCjk(text) {
  for (const ch of String(text)) if (ch.codePointAt(0) > 0x2e80) return true
  return false
}

function pdfHexLatin(text) {
  let out = ''
  for (const ch of text) out += winAnsiByte(ch).toString(16).padStart(2, '0')
  return out
}

function utf16Hex(text) {
  let out = ''
  for (let i = 0; i < String(text).length; i++) {
    const unit = String(text).charCodeAt(i)
    if (unit >= 0xd800 && unit <= 0xdbff && i + 1 < String(text).length) { out += unit.toString(16).padStart(4, '0'); continue }
    out += unit.toString(16).padStart(4, '0')
  }
  return out
}

/** Build a PDF from the document content model. */
export function writePdf(doc, opts = {}) {
  const d = { kind: 'document', meta: {}, blocks: [], ...doc }
  const pageW = opts.width ?? 595.28
  const pageH = opts.height ?? 841.89
  const marginX = opts.margin ?? 56
  const marginY = opts.marginY ?? 56
  const usable = pageW - marginX * 2
  const flatText = JSON.stringify(d.blocks) + (d.meta.title || '')
  const cjk = needsCjk(flatText)
  // —— 阶段一：内嵌字体链（默认开；DSH_OFFICE_PDF_EMBED_CJK=0 回退旧 STSong 路径）——
  const chain = process.env.DSH_OFFICE_PDF_EMBED_CJK === '0' ? null : new FontChain()
  const embedNames = new Map()   // SfntFont -> 'F5','F6',…（layout 期按需分配资源名）
  let usedLegacyCjk = false      // 是否真的画过 STSong 回落字符（决定 F4 是否落对象）
  pdfLayout = { chain, cjk }

  const objectStrings = []
  let nextId = 1
  const addObj = () => { const id = nextId++; objectStrings[id] = null; return id }
  const catalogId = addObj()
  const pagesId = addObj()
  const f1 = addObj(); const f2 = addObj(); const f3 = addObj()
  objectStrings[f1] = `<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>`
  objectStrings[f2] = `<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>`
  objectStrings[f3] = `<</Type/Font/Subtype/Type1/BaseFont/Courier/Encoding/WinAnsiEncoding>>`
  const fontRes = { F1: f1, F2: f2, F3: f3 }
  // 旧的"文档一含 CJK 就预建 STSong F4"挪到 layout 之后：现在 F4 只在
  // **真的画了回落字符**（内嵌链 miss）时才建对象，常态中文文档走内嵌字体。

  const pageStreams = []
  let curOps = []
  let y = pageH - marginY
  const newPage = () => { if (curOps.length) pageStreams.push(curOps); curOps = []; y = pageH - marginY }
  const ensureRoom = need => { if (y - need < marginY) { curOps.length && pageStreams.push(curOps); curOps = []; y = pageH - marginY } }

  // 内嵌字体资源名按需分配（F4 留给 STSong 回落）；page 对象在 layout 之后
  // 才组装，所以这里先记账、最后统一写进 fontRes。
  const embedResName = ef => {
    let n = embedNames.get(ef)
    if (!n) { n = `F${5 + embedNames.size}`; embedNames.set(ef, n) }
    return n
  }

  // ---------------- 图片 XObject（第十二轮 需求 1c） ----------------
  // JPEG 原样内嵌（`/DCTDecode`，不解码、不重编码，体积与质量无损）；
  // PNG 解成原始采样后 FlateDecode，带 alpha 的再挂一张 `/SMask`。zlib 是 Node 内置，合规。
  // 认不出的格式 / 读不到的文件**逐张记账**（`imagesSkipped`），绝不静默丢图。
  const imagesSkipped = []
  const imagesReused = []       // [{name, resource, reason}] 内容去重（需求 4c）
  const imagesSizing = []       // [{name, resource, px, pt, rule, capped, align}] 尺寸换算（需求 4e）
  const imageByHash = new Map() // 内容 SHA-256 → {resName, id, width, height}
  const xobjects = []          // [{name, id}]
  const addImageXObject = block => {
    const name = String(block.name ?? '')
    if (!name) { imagesSkipped.push({ name: '(无路径)', reason: 'markdown 里的图片没有可用路径' }); return null }
    let img
    try { img = readImageBytes(name) } catch (e) {
      imagesSkipped.push({ name, reason: `读不到文件：${e?.message || e}` }); return null
    }
    const imgInfo = sniffImage(img.buf)
    if (!imgInfo || !imgInfo.width || !imgInfo.height) { imagesSkipped.push({ name, reason: '不认识的图片格式（支持 PNG/JPEG/GIF/BMP）' }); return null }
    // 内容去重（需求 4c）：同图复用同一 /XObject 与资源名，不重复写对象/不重复占名字
    const imgHash = createHash('sha256').update(img.buf).digest('hex')
    const dup = imageByHash.get(imgHash)
    if (dup) {
      imagesReused.push({ name, resource: dup.resName, reason: '内容 SHA-256 与已有图片对象相同，复用同一 /XObject 与资源名' })
      return dup
    }
    const resName = `Im${xobjects.length + 1}`
    if (imgInfo.kind === 'jpeg') {
      if (imgInfo.components === 4) { imagesSkipped.push({ name, reason: 'CMYK JPEG 暂不支持（PDF 需要反相处理）' }); return null }
      const id = addObj()
      objectStrings[id] = [
        `<</Type/XObject/Subtype/Image/Width ${imgInfo.width}/Height ${imgInfo.height}`
        + `/ColorSpace /${imgInfo.components === 1 ? 'DeviceGray' : 'DeviceRGB'}/BitsPerComponent 8`
        + `/Filter/DCTDecode/Length ${img.buf.length}>>\nstream\n`,
        img.buf,
        '\nendstream',
      ]
      xobjects.push({ name: resName, id })
      const entry = { resName, id, width: imgInfo.width, height: imgInfo.height, kind: 'jpeg' }
      imageByHash.set(imgHash, entry)
      return entry
    }
    const raw = imageToRaw(img.buf)
    if (!raw) {
      imagesSkipped.push({ name, reason: `${String(imgInfo.kind).toUpperCase()} 解不出原始采样（PNG 需 8-bit 非隔行、BMP 需 8/24/32-bit BI_RGB、GIF 取首帧）` })
      return null
    }
    const id = addObj()
    const deflated = deflateSync(raw.data, { level: 6 })
    let smask = ''
    if (raw.alpha) {
      const smId = addObj()
      const aDef = deflateSync(raw.alpha, { level: 6 })
      objectStrings[smId] = [
        `<</Type/XObject/Subtype/Image/Width ${raw.width}/Height ${raw.height}/ColorSpace /DeviceGray`
        + `/BitsPerComponent 8/Filter/FlateDecode/Length ${aDef.length}>>\nstream\n`,
        aDef,
        '\nendstream',
      ]
      smask = `/SMask ${smId} 0 R`
    }
    objectStrings[id] = [
      `<</Type/XObject/Subtype/Image/Width ${raw.width}/Height ${raw.height}/ColorSpace /${raw.colorSpace}`
      + `/BitsPerComponent 8/Filter/FlateDecode/Length ${deflated.length}${smask}>>\nstream\n`,
      deflated,
      '\nendstream',
    ]
    xobjects.push({ name: resName, id })
    const entry = { resName, id, width: raw.width, height: raw.height, kind: imgInfo.kind }
    imageByHash.set(imgHash, entry)
    return entry
  }

  /** 图片排到当前页流里（按可用宽度等比缩放；`block.width` 单位磅；`block.align` = center/right）。 */
  const drawImage = block => {
    const entry = addImageXObject(block)
    if (!entry) return
    const naturalW = entry.width * 72 / 96          // 96dpi 像素 → pt
    const asked = Number(block.width)
    const explicit = Number.isFinite(asked) && asked > 0
    const wantW = explicit ? asked : naturalW
    const w = Math.min(wantW, usable)
    const h = entry.width ? w * entry.height / entry.width : 0
    // 对齐（需求 4f）：省略 align 时恒从 marginX 起排 —— 与旧版**逐字节一致**
    let x = marginX
    const align = String(block.align ?? '').toLowerCase()
    if (align === 'center') x = marginX + (usable - w) / 2
    else if (align === 'right') x = marginX + (usable - w)
    imagesSizing.push({
      name: String(block.name ?? ''),
      resource: entry.resName,
      px: `${entry.width}×${entry.height}`,
      pt: `${Math.round(w * 100) / 100}×${Math.round(h * 100) / 100}`,
      rule: explicit ? 'width 参数' : '原图像素 × 72/96',
      capped: wantW > usable,
      align: align || null,
    })
    ensureRoom(h + 8)
    y -= h
    curOps.push(`q ${fmt(w)} 0 0 ${fmt(h)} ${fmt(x)} ${fmt(y)} cm /${entry.resName} Do Q`)
    y -= 6
  }
  const drawText = (text, x, yy, size, font, colorHex) => {
    if (!text) return
    const fill = colorHex ? `${hexRgb(colorHex)} rg` : '0 0 0 rg'
    let xx = x
    for (const seg of segmentText(text, font)) {
      if (seg.k === 'zero') continue           // 格式字符：零宽、不落内容流
      let fname
      let hex
      if (seg.k === 'embed') {
        fname = embedResName(seg.f)
        hex = ''
        for (const ch of seg.text) hex += String(chain.classify(ch).g.toString(16)).padStart(4, '0')
      } else if (seg.k === 'gb') {
        fname = 'F4'
        usedLegacyCjk = true
        hex = utf16Hex(seg.text)
      } else {
        fname = seg.font
        hex = pdfHexLatin(seg.text)
      }
      curOps.push(`BT ${fill} /${fname} ${fmt(size)} Tf 1 0 0 1 ${fmt(xx)} ${fmt(yy)} Tm <${hex}> Tj ET`)
      xx += seg.em * size
    }
  }
  const drawLine = (x1, y1, x2, y2, width = 0.6, color = '000000') => {
    curOps.push(`${hexRgb(color)} RG ${fmt(width)} w ${fmt(x1)} ${fmt(y1)} m ${fmt(x2)} ${fmt(y2)} l S`)
  }
  const drawRect = (x, yy, w, h, fill) => {
    curOps.push(`${hexRgb(fill)} rg ${fmt(x)} ${fmt(yy)} ${fmt(w)} ${fmt(h)} re f`)
  }

  const wrap = (text, size, font) => {
    const tokens = []
    let cur = ''
    for (const ch of text) {
      // 断词口径与 segmentText 一致：内嵌链下"非 WinAnsi 单字成 token"（☆/①/emoji
      // 不再粘在英文词里）；无链时维持旧规则（cp>0x2e80 单字成 token）。
      const single = pdfLayout?.chain ? !winAnsiOk(ch) : ch.codePointAt(0) > 0x2e80
      if (single) { if (cur) { tokens.push(cur); cur = '' } tokens.push(ch) }
      else if (ch === ' ') { if (cur) { tokens.push(cur); cur = '' } tokens.push(' ') }
      else cur += ch
    }
    if (cur) tokens.push(cur)
    const lines = []
    let line = ''
    for (const tk of tokens) {
      const trial = line + tk
      if (widthOf(trial, size, font) > usable && line) { lines.push(line.replace(/\s+$/, '')); line = tk === ' ' ? '' : tk }
      else line = trial
    }
    if (line) lines.push(line)
    return lines.length ? lines : ['']
  }

  const para = (text, size, { font = 'F1', color = null, indent = 0, gapBefore = 0, gapAfter = 4 } = {}) => {
    const lines = wrap(text, size, font)
    y -= gapBefore
    for (const ln of lines) {
      ensureRoom(size * 1.4)
      y -= size * 1.35
      drawText(ln, marginX + indent, y, size, font, color)
    }
    y -= gapAfter
  }

  const runPara = (runs, size, { indent = 0, leading = 1.35 } = {}) => {
    const segs = []
    for (const r of runs || []) {
      const rr = typeof r === 'string' ? { text: r } : r
      const font = rr.code ? 'F3' : rr.bold ? 'F2' : 'F1'
      segs.push({ text: String(rr.text ?? ''), font, color: rr.link ? '0563C1' : rr.code ? '9C3265' : null })
    }
    // join consecutive, wrap across segments
    let line = []
    let lineWidth = indent
    const flush = () => {
      if (!line.length) return
      ensureRoom(size * leading)
      y -= size * leading
      let xx = marginX + indent
      for (const seg of line) { drawText(seg.text, xx, y, size, seg.font, seg.color); xx += seg.w }
      line = []
    }
    for (const seg of segs) {
      let rest = seg.text
      while (rest.length) {
        // —— 死循环修复（任务 C）——
        // 旧写法 `fit = Math.max(1, Math.floor(fit * (usable - lineWidth) / w) - 1)`
        // 把 `fit` 夹在 ≥1。一旦"本行剩余宽度装不下**一个**字符"（`lineWidth + w > usable`
        // 且 `fit === 1`），分子 `usable - lineWidth` 已经 ≤ 0 或乘完仍不足，`fit` 每次
        // 都回到 1、`w` 不变 → **判定条件永远为真**，`while` 同步死转，事件循环被占死
        // （所以 `Promise.race` 超时都触发不了，整个宿主进程冻住）。
        // 触发条件与字符种类无关，只与"某一行剩余宽度 < 单字宽度"有关：私用区码点
        // 因为按 1em 记宽而最先撞上（≥47 个就挂），汉字、扩展 B、甚至 `_`/ASCII 长串
        // 同样会挂（`_`.repeat(200) 一样冻宿主）—— 它不是"私用区专属"问题。
        // 修法：把"估算"只当起点，用真实宽度收敛；并且一旦出现"空行也放不下一个字"，
        // 就**硬放 1 个字**，保证每次循环至少推进 1 个字符（宁可溢出，绝不挂死）。
        const avail = usable - lineWidth
        let fit = rest.length
        const full = widthOf(rest, size, seg.font)
        if (full > avail) {
          fit = Math.max(0, Math.floor(rest.length * avail / Math.max(1e-6, full)))
          while (fit > 1 && widthOf(rest.slice(0, fit), size, seg.font) > avail) fit--
          while (fit < rest.length && widthOf(rest.slice(0, fit + 1), size, seg.font) <= avail) fit++
        }
        if (fit <= 0) {
          if (line.length) { flush(); lineWidth = indent; continue }
          fit = 1                                   // 空行也放不下一个字符：硬放，防挂死
        }
        const piece = rest.slice(0, fit)
        const pw = widthOf(piece, size, seg.font)
        if (lineWidth + pw > usable && line.length) { flush(); lineWidth = indent }
        line.push({ ...seg, text: piece, w: pw })
        lineWidth += pw
        rest = rest.slice(fit)
        if (!rest) break
        if (/[\u2e80-\uffff]/.test(piece[piece.length - 1] ?? '') && /[\u2e80-\uffff]/.test(rest[0] ?? '')) continue
      }
      if (/^\s+$/.test(seg.text) || / $/.test(seg.text)) { line.push({ ...seg, text: ' ', w: widthOf(' ', size, seg.font) }); lineWidth += widthOf(' ', size, seg.font) }
    }
    flush()
    y -= 4
  }

  const table = rows => {
    if (!rows.length) return
    const cols = Math.max(1, ...rows.map(r => r.length))
    const cellPad = 5
    let colW = Array.from({ length: cols }, (_, c) => Math.min(340, Math.max(...rows.map(r => widthOf(String(r[c] ?? '').slice(0, 60), 9.5)), 36) + cellPad * 2))
    const rawTotal = colW.reduce((a, b) => a + b, 0)
    if (rawTotal > usable) colW = colW.map(w => w * usable / rawTotal)
    const totalW = colW.reduce((a, b) => a + b, 0)
    const lineHeight = 12
    const wrapped = rows.map(r => r.map((cell, c) => wrapLines(String(cell ?? '').replace(/\r?\n/g, ' '), 9.5, colW[c] - cellPad * 2, 14)))
    rows.forEach((r, ri) => {
      const h = Math.max(1, ...wrapped[ri].map(l => l.length)) * lineHeight + cellPad
      ensureRoom(h)
      const top = y
      let xx = marginX
      if (ri === 0) drawRect(marginX, top - h, totalW, h, 'DCE6F1')
      wrapped[ri].forEach((lines, c) => {
        lines.forEach((ln, li) => drawText(ln, xx + cellPad, top - cellPad - 9 - li * lineHeight, 9.5, ri === 0 ? 'F2' : 'F1', null))
        xx += colW[c]
      })
      for (let c = 0, lx = marginX; c < cols + 1; c++) { drawLine(lx, top + 2, lx, top - h, 0.4, '999999'); lx += colW[c] ?? 0 }
      drawLine(marginX, top + 2, marginX + totalW, top + 2, 0.4, '999999')
      drawLine(marginX, top - h, marginX + totalW, top - h, 0.4, '999999')
      y -= h
    })
    y -= 8
  }

  const wrapLines = (text, size, maxW, maxLines) => {
    const lines = []
    let cur = ''
    for (const ch of text) {
      if (widthOf(cur + ch, size) > maxW) { lines.push(cur); cur = ch; if (lines.length >= maxLines) { lines[maxLines - 1] = lines[maxLines - 1].slice(0, -1) + '…'; return lines } }
      else cur += ch
    }
    lines.push(cur)
    return lines
  }

  for (const b of d.blocks) {
    if (b.type === 'heading') {
      const lvl = Math.min(6, b.level || 1)
      const size = [20, 16, 13.5, 12, 11, 10.5][lvl - 1]
      para(String(b.text ?? ''), size, { font: 'F2', color: lvl <= 2 ? '1F3864' : null, gapBefore: lvl <= 2 ? 10 : 6, gapAfter: 5 })
    } else if (b.type === 'paragraph') {
      runPara(b.runs || [{ text: b.text || '' }], 10.5)
    } else if (b.type === 'list') {
      ;(b.items || []).forEach((it, idx) => {
        const marker = b.ordered ? `${idx + 1}. ` : '• '
        const indent = 8 + Math.min(4, it.level || 0) * 14
        runPara([{ text: marker + String(it.text ?? '') }], 10.5, { indent })
      })
      y -= 4
    } else if (b.type === 'table') table((b.rows || []).map(r => (Array.isArray(r) ? r : [String(r)])))
    else if (b.type === 'quote') runPara([{ text: String(b.text || ''), italic: true }], 10.5, { indent: 18 })
    else if (b.type === 'code') {
      for (const ln of String(b.text || '').split('\n')) para(ln || ' ', 9, { font: 'F3', gapBefore: 0, gapAfter: 0 })
      y -= 6
    } else if (b.type === 'image') drawImage(b)
    else if (b.type === 'hr') {
      ensureRoom(12)
      y -= 6
      drawLine(marginX, y, pageW - marginX, y, 0.6, '808080')
      y -= 8
    } else if (b.type === 'pagebreak') { if (curOps.length) pageStreams.push(curOps); curOps = []; y = pageH - marginY }
  }
  if (curOps.length) pageStreams.push(curOps)
  if (!pageStreams.length) pageStreams.push(['BT 0 0 0 rg /F1 12 Tf 1 0 0 1 56 786 Tm <445348204f6666696365> Tj ET'])

  // —— 阶段一：layout 结束后按实际用量装配字体对象 ——
  const info = opts.info || {}
  if (usedLegacyCjk) {
    const fCjk = addObj()
    const desc = addObj()
    const fd = addObj()
    objectStrings[fCjk] = `<</Type/Font/Subtype/Type0/BaseFont/STSong-Light/Encoding/UniGB-UCS2-H/DescendantFonts[${desc} 0 R]>>`
    objectStrings[desc] = `<</Type/Font/Subtype/CIDFontType0/BaseFont/STSong-Light/CIDSystemInfo<< /Registry(Adobe-GB1)/Ordering(GB1)/Supplement 2>>/FontDescriptor[${fd} 0 R]/DW 1000>>`
    objectStrings[fd] = `<</Type/FontDescriptor/FontName/STSong-Light/Flags 4/FontBBox[-25 -254 1000 880]/ItalicAngle 0/Ascent 880/Descent -120/CapHeight 731/StemV 93>>`
    fontRes.F4 = fCjk
  }
  const embedReport = []
  for (const [ef, resName] of embedNames) {
    const { gids, chars } = chain.used(ef)
    const subset = ef.subset(gids)
    const packed = deflateSync(subset)
    const tag = subsetTag(ef, gids)
    const baseName = `${tag}+${pdfNameEsc(ef.name)}`
    const fileId = addObj()
    const tuId = addObj()
    const fdId = addObj()
    const descId = addObj()
    const type0Id = addObj()
    objectStrings[fileId] = `<</Length ${packed.length}/Length1 ${subset.length}>>\nstream\n${packed.toString('latin1')}\nendstream`
    const tu = toUnicodeCMap(chars)
    objectStrings[tuId] = `<</Length ${Buffer.byteLength(tu, 'latin1')}>>\nstream\n${tu}\nendstream`
    const s = 1000 / (ef.upem || 1000)
    const bbox = [ef.xMin, ef.yMin, ef.xMax, ef.yMax].map(v => Math.round(v * s))
    if (!(bbox[2] > bbox[0] && bbox[3] > bbox[1])) bbox.splice(0, 4, -1000, -1600, 2000, 1600)
    const asc = Math.round(ef.ascent * s) || 800
    const descY = Math.round(ef.descent * s) || -200
    const cap = Math.round(ef.capHeight * s) || Math.round(asc * 0.7)
    objectStrings[fdId] = `<</Type/FontDescriptor/FontName/${baseName}/Flags 4/FontBBox[${bbox.join(' ')}]/ItalicAngle ${Math.round(ef.italicAngle) || 0}/Ascent ${asc}/Descent ${descY}/CapHeight ${cap}/StemV 80/FontFile2 ${fileId} 0 R>>`
    objectStrings[descId] = `<</Type/Font/Subtype/CIDFontType2/BaseFont/${baseName}/CIDSystemInfo<< /Registry(Adobe)/Ordering(Identity)/Supplement 0 >>/FontDescriptor ${fdId} 0 R/DW 1000/W[${wArray(ef, gids)}]/CIDToGIDMap/Identity>>`
    objectStrings[type0Id] = `<</Type/Font/Subtype/Type0/BaseFont/${baseName}/Encoding/Identity-H/DescendantFonts[${descId} 0 R]/ToUnicode ${tuId} 0 R>>`
    fontRes[resName] = type0Id
    embedReport.push({ resource: resName, base: `${tag}+${ef.name}`, glyphs: gids.size, fontFileBytes: subset.length })
  }
  info.embedded = embedReport.length > 0
  info.legacyCjk = !!usedLegacyCjk
  info.fonts = embedReport
  info.notes = [...(chain?.notes || [])]
  // 插图账（第十二轮需求 1c）：嵌入成功的资源名 + **逐张的跳过原因**（调用方写进 stats/notice）
  info.images = xobjects.map(x => x.name)
  info.imagesSkipped = imagesSkipped
  // 图片记账（第二轮需求 4c/4e）：内嵌了哪几张、复用了哪些、尺寸怎么换算的
  if (xobjects.length) info.imageMedia = xobjects.map(x => x.name)
  if (imagesReused.length) info.imageReused = imagesReused
  if (imagesSizing.length) info.imageSizing = imagesSizing
  if (imagesSkipped.length) {
    info.notes.push(`${imagesSkipped.length} 张图片未能嵌入：`
      + imagesSkipped.map(x => `${x.name}（${x.reason}）`).join('；'))
  }
  pdfLayout = null

  const kidsIds = []
  const contentIds = []
  for (const ops of pageStreams) {
    const contentId = addObj()
    contentIds.push(contentId)
    const stream = ops.join('\n')
    objectStrings[contentId] = `<</Length ${Buffer.byteLength(stream, 'latin1')}>>\nstream\n${stream}\nendstream`
  }
  // 页对象字典的闭合必须精确：/Resources<< /Font<< … >> >> 一共 4 个 `>`。
  // 旧版在页面字典里写了 5 个（多一个），本机 lenient 解析器容忍，但 WinRT 的
  // PdfDocument.LoadFromFileAsync 判其畸形 → 插件产出的 PDF 一律栅格化失败
  // （"One or more errors occurred"）。这里删掉多余的一个。
  for (const [i, contentId] of contentIds.entries()) {
    const pageId = addObj()
    kidsIds.push(pageId)
    const fontResStr = Object.entries(fontRes).map(([name, id]) => `/${name} ${id} 0 R`).join(' ')
    // 图片资源只在真有图片时才追加 `/XObject<< … >>`：**没有图片时**资源字典必须仍是
    // `/Resources<< /Font<<…>>>>`（4 个连续的 `>`，多一个空格都会破掉 WinRT 那条不变式）。
    const xoResStr = xobjects.length ? `/XObject<< ${xobjects.map(x => `/${x.name} ${x.id} 0 R`).join(' ')} >>` : ''
    objectStrings[pageId] = `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${fmt(pageW)} ${fmt(pageH)}]/Resources<< /Font<< ${fontResStr} >>${xoResStr}>>/Contents ${contentId} 0 R>>`
    void i
  }
  objectStrings[pagesId] = `<</Type/Pages/Count ${kidsIds.length}/Kids[${kidsIds.map(i => `${i} 0 R`).join(' ')}]>>`
  const infoId = addObj()
  objectStrings[infoId] = `<< /Title(${pdfLiteral(d.meta.title || 'Document')})/Producer(DSH Office)>>`
  objectStrings[catalogId] = `<< /Type/Catalog/Pages ${pagesId} 0 R>>`

  const header = '%PDF-1.4\n%\xE9\xEA\xF0\xF1\n'
  const chunks = [Buffer.from(header, 'latin1')]
  const offsets = []
  // 图片流是**二进制**，不能过 latin1 字符串往返 —— 对象体允许是「字符串片段 + Buffer 片段」
  // 的数组，这里统一拍平（旧版只支持字符串）。
  const objBody = id => {
    const v = objectStrings[id]
    if (v === null || v === undefined) return Buffer.from('<<>>', 'latin1')
    if (Array.isArray(v)) return Buffer.concat(v.map(p => (typeof p === 'string' ? Buffer.from(p, 'latin1') : p)))
    return Buffer.from(v, 'latin1')
  }
  for (let id = 1; id < nextId; id++) {
    offsets[id] = chunks.reduce((a, c) => a + c.length, 0)
    chunks.push(Buffer.from(`${id} 0 obj\n`, 'latin1'), objBody(id), Buffer.from('\nendobj\n', 'latin1'))
  }
  const xrefStart = chunks.reduce((a, c) => a + c.length, 0)
  let xref = `xref\n0 ${nextId}\n0000000000 65535 f \n`
  for (let id = 1; id < nextId; id++) xref += `${String(offsets[id] ?? 0).padStart(10, '0')} 00000 n \n`
  xref += `trailer\n<< /Size ${nextId}/Root ${catalogId} 0 R/Info ${infoId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`
  chunks.push(Buffer.from(xref, 'latin1'))
  return Buffer.concat(chunks)
}

function pdfLiteral(s) {
  let out = ''
  for (const ch of String(s)) out += ch === '(' || ch === ')' || ch === '\\' ? '\\' + ch : (ch.codePointAt(0) <= 255 ? ch : '?')
  return out
}

// ---------------- 阶段一：内嵌字体的 PDF 侧装配件 ----------------

/** 6 位大写字母子集前缀（同文档不同字体/字形集不同名）。 */
function subsetTag(font, gids) {
  let h = 2166136261 ^ font.name.length
  for (const g of gids) { h ^= g; h = Math.imul(h, 16777619) >>> 0 }
  let s = ''
  let x = h || 1
  for (let i = 0; i < 6; i++) { s += String.fromCharCode(65 + (x % 26)); x = Math.floor(x / 26) + i * 7 + 11 }
  return s
}

/** PDF 名字转义（空格与非可见字符 -> #XX）。 */
function pdfNameEsc(s) {
  let out = ''
  for (const ch of String(s)) {
    const c = ch.codePointAt(0)
    if (c <= 32 || c > 0x7e || '()<>[]{}/%#'.includes(ch)) out += `#${c.toString(16).padStart(2, '0').toUpperCase()}`
    else out += ch
  }
  return out || 'Embedded'
}

/** /W 数组：按 GID 升序，把"连续且等宽"的 GID 折成 `first last width` 三元组。 */
function wArray(font, gids) {
  const sorted = [...gids].sort((a, b) => a - b)
  const wOf = g => Math.round(font.advance(g) * 1000 / (font.upem || 1000))
  const parts = []
  let start = -1
  let prev = -1
  let prevW = -1
  const flush = () => {
    if (start >= 0) parts.push(`${start} ${prev} ${prevW}`)
  }
  for (const g of sorted) {
    const w = wOf(g)
    if (start >= 0 && g === prev + 1 && w === prevW) { prev = g; continue }
    flush()
    start = g
    prev = g
    prevW = w
  }
  flush()
  return parts.join(' ')
}

/** 完整 ToUnicode CMap（GID -> 原字符；bfchar 每批 ≤100 条）。 */
function toUnicodeCMap(chars) {
  const entries = []
  for (const [gid, ch] of chars) entries.push(`<${gid.toString(16).padStart(4, '0').toUpperCase()}> <${utf16Hex(ch).toUpperCase()}>`)
  const blocks = []
  for (let i = 0; i < entries.length; i += 100) {
    blocks.push(`${Math.min(100, entries.length - i)} beginbfchar\n${entries.slice(i, i + 100).join('\n')}\nendbfchar`)
  }
  return [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange',
    ...blocks,
    'endcmap',
    'CMapName currentdict /CMap defineresource pop',
    'end',
    'end',
    '',
  ].join('\n')
}

function hexRgb(hex) {
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return '0 0 0'
  return `${(parseInt(hex.slice(0, 2), 16) / 255).toFixed(3)} ${(parseInt(hex.slice(2, 4), 16) / 255).toFixed(3)} ${(parseInt(hex.slice(4, 6), 16) / 255).toFixed(3)}`
}

function fmt(n) { return (Math.round(n * 100) / 100).toString() }

// expose page text assembly for the markdown layer
export function pdfSectionsToMarkdown(sections) {
  const out = []
  for (const sec of sections) {
    out.push(`<!-- 第 ${sec.page} 页 -->\n${sec.text}`)
  }
  return out.join('\n\n')
}
