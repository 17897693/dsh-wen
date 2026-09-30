// Legacy OLE2 (CFB) compound documents: Word 97-2003 (.doc/.wps),
// Excel 97-2003 (.xls/.et), PowerPoint 97-2003 (.ppt/.dps/.pps/.pot).
// Pure JS per [MS-CFB], [MS-DOC], [MS-XLS], PPT97 binary format.

const dec16 = new TextDecoder('utf-16le', { fatal: false })

const CP1252 = {}
for (let i = 0x80; i < 0x100; i++) CP1252[i] = String.fromCodePoint(i)
Object.assign(CP1252, { 0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ', 0x8e: 'Ž', 0x91: '\u2018', 0x92: '\u2019', 0x93: '\u201c', 0x94: '\u201d', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›', 0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ' })

function decodeAnsiBytes(buf) {
  let out = ''
  for (const b of buf) out += b >= 0x80 ? (CP1252[b] ?? '?') : String.fromCharCode(b)
  return out
}

// ---------------- OLE2 container ----------------

export function parseOle2(buf) {
  if (buf.length < 512) throw new Error('ole2: file too small')
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const sectorShift = dv.getUint16(30, true)
  const miniShift = dv.getUint16(32, true)
  const numDifat = dv.getUint32(72, true)
  const dirStart = dv.getInt32(48, true)
  const miniStart = dv.getInt32(60, true)
  const miniCutoff = dv.getUint32(56, true)
  const difatStart = dv.getInt32(68, true)
  if (sectorShift !== 9) throw new Error('ole2: only 512-byte sectors supported')
  const SECT = 512
  const sectorOff = s => (s + 1) * SECT

  const fatSectIds = []
  for (let i = 0; i < 109; i++) {
    const id = dv.getInt32(76 + i * 4, true)
    if (id >= 0 && id < 0x8000000) fatSectIds.push(id)
  }
  let next = difatStart
  for (let chainGuard = 0; chainGuard < numDifat && next >= 0 && next < 0x8000000; chainGuard++) {
    const base = sectorOff(next)
    if (base + SECT > buf.length) break
    for (let i = 0; i < SECT / 4 - 1; i++) {
      const id = dv.getInt32(base + i * 4, true)
      if (id >= 0 && id < 0x8000000) fatSectIds.push(id)
    }
    next = dv.getInt32(base + SECT - 4, true)
  }
  const fat = new Int32Array(Math.max(1, fatSectIds.length) * (SECT / 4))
  fatSectIds.forEach((sid, k) => {
    const base = sectorOff(sid)
    for (let i = 0; i < SECT / 4 && base + i * 4 + 4 <= buf.length; i++) fat[k * (SECT / 4) + i] = dv.getInt32(base + i * 4, true)
  })
  const chain = start => {
    const out = []
    let s = start
    const seen = new Set()
    while (s >= 0 && s < fat.length && !seen.has(s) && out.length < 500000) { seen.add(s); out.push(s); s = fat[s] }
    return out
  }
  const readChain = (start, size) => {
    const out = Buffer.alloc(Math.max(0, Math.min(size, 64 * 1024 * 1024)))
    let pos = 0
    for (const s of chain(start)) {
      const base = sectorOff(s)
      if (base >= buf.length) break
      const n = Math.min(SECT, out.length - pos, buf.length - base)
      if (n <= 0) break
      buf.copy(out, pos, base, base + n)
      pos += n
    }
    return out.subarray(0, pos)
  }
  const dirBytes = Buffer.concat(chain(dirStart).map(s => Buffer.from(buf.subarray(sectorOff(s), Math.min(sectorOff(s) + SECT, buf.length)))))
  const entries = []
  for (let i = 0; i + 128 <= dirBytes.length; i += 128) {
    const e = dirBytes.subarray(i, i + 128)
    const nameLen = e.readUInt16LE(64)
    if (nameLen < 2 || nameLen > 64) continue
    const name = dec16.decode(e.subarray(0, nameLen - 2))
    const type = e[66]
    const start = e.readInt32LE(116)
    const size = e.readUInt32LE(120) | (e.readUInt32LE(124) << 16) << 16
    entries.push({ name, type, start, size })
  }
  const root = entries.find(e => e.type === 5) || { start: -1, size: 0 }
  const miniSects = root.size > 0 ? readChain(root.start, root.size) : Buffer.alloc(0)
  const miniFatBytes = readChain(miniStart, 1 << 26)
  const miniFat = new Int32Array(miniFatBytes.buffer, miniFatBytes.byteOffset, Math.floor(miniFatBytes.length / 4))
  const miniSectSize = 1 << miniShift

  const byName = new Map()
  for (const e of entries) if (e.name && (e.type === 2 || e.type === 1 || e.type === 5)) byName.set(e.name, e)

  return {
    names: [...byName.keys()],
    has: name => byName.has(name),
    read(name) {
      const e = byName.get(name)
      if (!e) return undefined
      if (e.size >= miniCutoff) return readChain(e.start, e.size)
      const out = Buffer.alloc(Math.max(0, Math.min(e.size, 32 * 1024 * 1024)))
      let s = e.start
      let pos = 0
      const seen = new Set()
      while (s >= 0 && !seen.has(s) && pos < out.length) {
        seen.add(s)
        const base = s * miniSectSize
        if (base >= miniSects.length) break
        const n = Math.min(miniSectSize, out.length - base >= 0 ? out.length - pos : 0, miniSects.length - base)
        if (n <= 0) break
        miniSects.copy(out, pos, base, base + n)
        pos += n
        s = miniFat[s] ?? -2
      }
      return out.subarray(0, pos)
    },
  }
}

// ---------------- Word (.doc / WPS-Word) ----------------

export function readLegacyDoc(ole) {
  const wd = ole.read('WordDocument')
  if (!wd || wd.length < 0x200) throw new Error('doc/wps: 未找到 WordDocument 流（该 .wps 可能为新版 WPS 专有格式，建议在 WPS 中另存为 .docx）')
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength)
  const flags = dv.getUint16(0x000A, true)
  const whichTable = (flags & 0x0200) !== 0 ? '1Table' : '0Table'
  const table = ole.read(whichTable) || ole.read(whichTable === '1Table' ? '0Table' : '1Table')
  const fcClx = dv.getUint32(0x01A2, true)
  const lcbClx = dv.getUint32(0x01A6, true)
  let text = ''
  if (table && lcbClx > 0 && fcClx + lcbClx <= table.length) {
    const pieces = parsePieceTable(table.subarray(fcClx, fcClx + lcbClx))
    if (pieces) {
      for (const p of pieces) {
        if (p.compressed) {
          const slice = wd.subarray(p.fc, Math.min(p.fc + p.length, wd.length))
          text += decodeAnsiBytes(slice)
        } else {
          const slice = wd.subarray(p.fc, Math.min(p.fc + p.length * 2, wd.length))
          text += dec16.decode(slice)
        }
      }
      text = cleanDocControl(text)
    }
  }
  if (!text.trim()) {
    // naive fallbacks: UTF-16 scan then CP1252 run scan
    const as16 = cleanDocControl(dec16.decode(wd))
    if (as16.replace(/\s/g, '').length > 20) text = as16
    else text = cleanDocControl(decodeAnsiBytes(wd.subarray(0x200)))
  }
  if (!text.trim()) throw new Error('doc/wps: 未能提取文本（可能为受保护或异常格式）')
  return { kind: 'document', meta: {}, blocks: textToBlocks(text) }
}

function parsePieceTable(clx) {
  let i = 0
  while (i < clx.length) {
    const t = clx[i]
    if (t === 1) { const cb = clx.readUInt16LE(i + 1); i += 3 + cb; continue }
    if (t === 2) {
      const lcb = clx.readUInt32LE(i + 1)
      const plc = clx.subarray(i + 5, i + 5 + lcb)
      const n = Math.floor((plc.length - 4) / 12)
      const pieces = []
      for (let k = 0; k < n; k++) {
        const pcd = plc.subarray(4 * (n + 1) + k * 8)
        const fc = pcd.readUInt32LE(2)
        const compressed = (fc & 0x40000000) !== 0
        const cpStart = plc.readUInt32LE(k * 4)
        const cpEnd = plc.readUInt32LE((k + 1) * 4)
        const length = cpEnd - cpStart
        let realFc = fc & 0x3FFFFFFF
        if (compressed) realFc = Math.floor(realFc / 2)
        pieces.push({ fc: realFc, length, compressed })
      }
      return pieces
    }
    i++
  }
  return null
}

function cleanDocControl(text) {
  let out = ''
  let fieldDepth = 0
  for (const ch of text) {
    const c = ch.codePointAt(0)
    if (c === 0x13) { fieldDepth++; continue }
    if (c === 0x14) { fieldDepth = 0; continue }
    if (c === 0x15) { if (fieldDepth > 0) fieldDepth--; continue }
    if (fieldDepth) continue
    if (c === 0x0D || c === 0x0B || c === 0x1E) { out += '\n'; continue }
    if (c === 0x07) { out += ' | '; continue }
    if (c === 0x09 || c === 0x1F) { out += '\t'; continue }
    if (c === 0x01 || c === 0x02 || c === 0x08 || c === 0x0C || c === 0x05 || c === 0x0E || c === 0x0F || c === 0x10) continue
    if (c < 0x20 && c !== 10) continue
    out += ch
  }
  return out.replace(/ {2,}/g, ' ').replace(/\t+/g, ' ')
}

export function textToBlocks(text) {
  const blocks = []
  for (const line of String(text).replace(/\r\n?/g, '\n').split('\n')) {
    const t = line.replace(/\u0000/g, '').trim()
    if (t) blocks.push({ type: 'paragraph', runs: [{ text: t }] })
  }
  return blocks
}

// ---------------- Excel (.xls / WPS ET) ----------------

const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 22, 45, 46, 47])

export function readLegacyXls(ole) {
  const wb = ole.read('Workbook') || ole.read('Book')
  if (!wb) throw new Error('xls/et: 未找到 Workbook 流（该 .et 文件可能为新版 WPS 专有格式，建议在 WPS 中另存为 .xlsx）')
  const recs = []
  let p = 0
  while (p + 4 <= wb.length) {
    const type = wb.readUInt16LE(p)
    const len = wb.readUInt16LE(p + 2)
    const data = wb.subarray(p + 4, Math.min(p + 4 + len, wb.length))
    recs.push({ type, data })
    if (len === 0 && type === 0) break
    p += 4 + len
  }
  // gather SST with CONTINUE payloads
  const sst = []
  let sstIx = -1
  for (let i = 0; i < recs.length; i++) if (recs[i].type === 0x00FC) { sstIx = i; break }
  if (sstIx >= 0) {
    const payloads = [recs[sstIx].data]
    for (let i = sstIx + 1; i < recs.length && recs[i].type === 0x003C; i++) payloads.push(recs[i].data)
    readSst(payloads, sst)
  }
  const fmts = new Map()
  let custom = 0
  for (const r of recs) if (r.type === 0x0513 && r.data.length >= 3) {
    const code = readXlsString(r.data, 2, () => {})
    fmts.set(164 + custom++, code)
  }
  const cellXfFmt = []
  for (const r of recs) if (r.type === 0x00E0 && r.data.length >= 6) cellXfFmt.push(r.data.readUInt16LE(4))

  const sheets = []
  let current = null
  let bof = 0
  const isDateXf = xf => {
    const id = cellXfFmt[xf] ?? 0
    if (BUILTIN_DATE_FMTS.has(id)) return true
    const code = fmts.get(id)
    if (!code) return false
    const stripped = code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
    return /[ymdhs]/i.test(stripped)
  }
  for (const r of recs) {
    if (r.type === 0x0085) { // BOUNDSHEET
      const name = readXlsString(r.data, 6, () => {})
      sheets.push({ name: name || `Sheet${sheets.length + 1}`, rows: new Map() })
      continue
    }
    if (r.type === 0x0809) { bof++; current = bof === 1 ? null : sheets[sheets.length - 1] || null; continue }
    if (r.type === 0x000A) { bof--; continue }
    if (!current) continue
    cellFrom(r, current.rows, sst, isDateXf)
  }
  const out = sheets.map(sh => flattenSheetRows(sh))
  return { kind: 'workbook', meta: {}, sheets: out.length ? out : [{ name: 'Sheet1', rows: [] }] }
}

function readSst(payloads, sst) {
  let vi = 0
  let pos = 8 // total(4) unique(4)
  const unique = payloads[0].readUInt32LE(4)
  const ensure = n => vi < payloads.length && pos + n > payloads[vi].length
  let guard = 0
  while (sst.length < unique && guard++ < 500000) {
    if (vi >= payloads.length) break
    if (pos >= payloads[vi].length) {
      if (vi + 1 >= payloads.length) break
      vi++; pos = 0
      // continuation begins with a fresh option byte
    }
    if (pos + 3 > payloads[vi].length) {
      if (vi + 1 >= payloads.length) break
      vi++; pos = 0
    }
    if (pos + 3 > payloads[vi].length) break
    const cch = payloads[vi].readUInt16LE(pos)
    const flags = payloads[vi][pos + 2]
    pos += 3
    let fHigh = (flags & 1) !== 0
    if (flags & 8) pos += 2 // rich runs (skip counts; text below is plain)
    if (flags & 4) pos += 4 // ext
    let text = ''
    let remaining = cch
    let localGuard = 0
    while (remaining > 0 && localGuard++ < 200000) {
      if (vi >= payloads.length) break
      if (pos >= payloads[vi].length) {
        vi++
        pos = 0
        if (vi < payloads.length) {
          const newFlags = payloads[vi][pos]
          if (newFlags !== undefined) fHigh = (newFlags & 1) !== 0
          pos += 1
          continue
        }
        break
      }
      if (fHigh) {
        const avail = Math.floor((payloads[vi].length - pos) / 2)
        const take = Math.min(remaining, avail)
        for (let k = 0; k < take; k++) { text += String.fromCharCode(payloads[vi].readUInt16LE(pos)); pos += 2 }
        remaining -= take
        if (remaining > 0 && take === 0) { vi++; pos = 0 }
      } else {
        const take = Math.min(remaining, payloads[vi].length - pos)
        text += decodeAnsiBytes(payloads[vi].subarray(pos, pos + take))
        pos += take
        remaining -= take
      }
    }
    void ensure
    sst.push(text)
  }
  return sst
}

function readXlsString(data, offset) {
  if (offset + 3 > data.length) return ''
  const cb = data.readUInt16LE(offset)
  const flags = data[offset + 2]
  const fHigh = (flags & 1) !== 0
  let pos = offset + 3
  if (flags & 8) pos += 2
  if (flags & 4) pos += 4
  let chars = ''
  if (fHigh) {
    for (let i = 0; i < cb && pos + 1 < data.length; i++) { chars += String.fromCharCode(data.readUInt16LE(pos)); pos += 2 }
  } else {
    for (let i = 0; i < cb && pos < data.length; i++) { chars += decodeAnsiBytes(data.subarray(pos, pos + 1)); pos++ }
  }
  return chars
}

function cellFrom(r, rows, sst, isDateXf) {
  const { type, data } = r
  if (data.length < 6) return
  const row = data.readUInt16LE(0)
  const col = data.readUInt16LE(2)
  const key = `${row},${col}`
  if (type === 0x00FD) { // LABELSST
    const idx = data.readUInt32LE(6)
    rows.set(key, { v: sst[idx] ?? '', t: 's' })
  } else if (type === 0x0203) { // NUMBER
    const xf = data.readUInt16LE(4)
    const v = data.readDoubleLE(6)
    rows.set(key, numCell(v, isDateXf(xf)))
  } else if (type === 0x027E) { // RSTRING
    rows.set(key, { v: readXlsString(data, 6), t: 's' })
  } else if (type === 0x0204) { // LABEL (BIFF2-ish)
    rows.set(key, { v: readXlsString(data, 4), t: 's' })
  } else if (type === 0x0205) { // BOOLERR
    rows.set(key, { v: data[6] ? 'TRUE' : 'FALSE', t: 'b' })
  } else if (type === 0x0006) { // FORMULA cached value
    const xf = data.readUInt16LE(4)
    const b6 = data[6], b7 = data[7]
    if (b7 === 0 && b6 === 1) rows.set(key, { v: data[8] ? 'TRUE' : 'FALSE', t: 'b' })
    else if (b7 === 0 && b6 === 2) rows.set(key, { v: errorName(data[8]), t: 's' })
    else rows.set(key, numCell(data.readDoubleLE(6), isDateXf(xf)))
  } else if (type === 0x00BD) { // MULRK
    const c0 = col
    const pairs = Math.floor((data.length - 6) / 6)
    for (let i = 0; i < pairs; i++) {
      const xf = data.readUInt16LE(4 + i * 6)
      const rk = data.readInt32LE(6 + i * 6)
      rows.set(`${row},${c0 + i}`, numCell(decodeRK(rk), isDateXf(xf)))
    }
  }
}

function errorName(code) {
  return { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x17: '#VALUE!', 0x18: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A' }[code & 0xff] ?? '#ERR'
}

function decodeRK(rk) {
  const fInt = (rk & 2) !== 0
  const fDiv = (rk & 1) !== 0
  let v
  if (fInt) v = rk >> 2
  else {
    const buf = new DataView(new ArrayBuffer(8))
    buf.setInt32(4, rk & 0x7FFFFFFC, true)
    v = buf.getFloat64(0, true)
  }
  return fDiv ? v / 100 : v
}

function numCell(v, isDate) {
  if (!Number.isFinite(v)) return { v: '', t: 's' }
  if (isDate) {
    const d = serialToIso(v)
    if (d) return { v: d, t: 'd' }
  }
  return { v, t: 'n' }
}

function serialToIso(serial) {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) return null
  let days = serial
  if (days >= 61) days -= 2
  else if (days >= 60) days = days === 60 ? 59.5 : days - 1
  else days -= 1
  const d = new Date(Date.UTC(1900, 0, 1) + days * 86400000)
  const iso = d.toISOString()
  return iso.includes('T00:00:00') || /[1-9]/.test(iso.slice(11, 16)) ? (/[1-9]/.test(iso.slice(11, 16)) ? iso.slice(0, 16).replace('T', ' ') : iso.slice(0, 10)) : iso.slice(0, 10)
}

function flattenSheetRows(sh) {
  let maxRow = -1
  let maxCol = -1
  for (const key of sh.rows.keys()) {
    const [r, c] = key.split(',').map(Number)
    if (r > maxRow) maxRow = r
    if (c > maxCol) maxCol = c
  }
  const rows = []
  for (let r = 0; r <= maxRow; r++) {
    const out = []
    for (let c = 0; c <= maxCol; c++) out.push(sh.rows.get(`${r},${c}`) ?? { v: '', t: 's' })
    while (out.length && String(out[out.length - 1].v) === '') out.pop()
    rows.push(out)
  }
  while (rows.length && !rows[rows.length - 1].length) rows.pop()
  return { name: sh.name, rows }
}

// ---------------- PowerPoint (.ppt) ----------------

export function readLegacyPpt(ole) {
  const stream = ole.read('PowerPoint Document')
  if (!stream) throw new Error('ppt/dps: 未找到 PowerPoint Document 流（该 .dps 可能为新版 WPS 专有格式，建议在 WPS 中另存为 .pptx）')
  const slides = []
  let current = null
  const walk = buf => {
    let p = 0
    let guard = 0
    while (p + 8 <= buf.length && guard++ < 200000) {
      const verType = buf.readUInt16LE(p)
      const recVer = verType & 0x000f
      const recInstance = (verType & 0xfff0) >> 4
      const recType = buf.readUInt16LE(p + 2)
      const recLen = buf.readUInt32LE(p + 4)
      const end = Math.min(p + 8 + recLen, buf.length)
      const body = buf.subarray(p + 8, end)
      if (recVer === 0xF) {
        if (recType === 0x03F0) { current = { layout: 'content', title: '', bullets: [] }; slides.push(current) }
        walk(body)
      } else {
        if (recType === 0x0FA8) pushPptText(current, decodeUtf16(body))
        else if (recType === 0x0FA0) pushPptText(current, decodeBytesAnsi(body, (recInstance & 0x000f) === 0 ? 0 : 8))
      }
      if (end === p + 8 && recLen === 0) break
      p = end
    }
  }
  walk(stream)
  const clean = slides.filter(s => s.title || s.bullets.length)
  if (!clean.length) throw new Error('ppt/dps: 未找到文本内容（可能是纯图片演示文稿）')
  return { kind: 'slides', meta: {}, slides: clean }
}

function pushPptText(slide, text) {
  if (!slide || !text || !text.trim()) return
  for (const ln of text.replace(/\r/g, '\n').split(/[\r\n]+/)) {
    const t = ln.trim()
    if (!t) continue
    if (!slide.title) slide.title = t
    else slide.bullets.push({ text: t, level: 0 })
  }
}

function decodeUtf16(buf) {
  let out = ''
  for (let i = 0; i + 1 < buf.length; i += 2) out += String.fromCharCode(buf.readUInt16LE(i))
  return out
}

function decodeBytesAnsi(buf) {
  let out = ''
  for (let i = 0; i < buf.length; i += 2) {
    const b = buf[i]
    out += b >= 0x80 ? (CP1252[b] ?? '?') : String.fromCharCode(b)
  }
  return out
}
