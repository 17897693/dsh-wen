// Excel (.xlsx) reader/writer plus CSV interop on the OOXML zip container.
import { openZip, makeZip } from './zip.js'
import { parseXML, findAll, find as findOne, children, textOf, encodeEntities } from './xml.js'

const SPREAD_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

// ---------- cell utilities ----------

export function colName(index) {
  let s = ''
  index += 1
  while (index > 0) {
    const rem = (index - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    index = Math.floor((index - 1) / 26)
  }
  return s
}

export function parseCellRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref).trim())
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { col: col - 1, row: Number(m[2]) - 1 }
}

// Excel serial -> ISO date (1900 system with the 1900-02-29 bug included).
function serialToDate(serial) {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) return null
  const msPerDay = 86400000
  let days = serial
  if (days >= 61) days -= 2 // after the phantom Feb 29 1900
  else if (days >= 60) days = days === 60 ? 59.5 : days - 1
  else days -= 1
  const d = new Date(Date.UTC(1900, 0, 1) + days * msPerDay)
  return d.toISOString().replace('T', ' ').replace(/:\d\d\.\d+Z$/, m => m.length > 5 ? m.slice(1, 6) : '00:00').replace(':00:00', ':00')
}

const DATE_FMT_IDS = new Set([14, 15, 16, 17, 22, 45, 46, 47])

function collectDateStyleIndexes(stylesXml) {
  // returns Set of cellXf indexes whose numFmt is a date format
  const dateXfs = new Set()
  if (!stylesXml) return dateXfs
  try {
    const st = parseXML(stylesXml)
    const custom = new Map()
    for (const f of findAll(st, 'numFmt')) custom.set(Number(f.attrs.numFmtId), f.attrs.formatCode || '')
    const numFmtDate = id => {
      if (DATE_FMT_IDS.has(id)) return true
      const code = custom.get(id)
      if (!code) return false
      const stripped = code.replace(/\[[^\]]*\]/g, '')
      return /[ymdhs]/i.test(stripped) && !/^[^ymdhs]*$/i.test(stripped)
    }
    const cellXfs = findOne(st, 'cellXfs')
    if (cellXfs) children(cellXfs, 'xf').forEach((xf, i) => {
      if (numFmtDate(Number(xf.attrs.numFmtId || 0))) dateXfs.add(i)
    })
  } catch { /* styles unsupported → treat numbers as numbers */ }
  return dateXfs
}

// ---------- reading ----------

export function readXlsx(zipInput) {
  const zipObj = typeof zipInput === 'object' && zipInput.names ? zipInput : openZip(zipInput)
  const wb = zipObj.getText('xl/workbook.xml')
  if (wb === undefined) throw new Error('xlsx: xl/workbook.xml missing')
  const wbDoc = parseXML(wb)
  const rels = parseRelationships(zipObj.getText('xl/_rels/workbook.xml.rels'))
  const shared = readSharedStrings(zipObj)
  const sheets = []
  for (const sh of children(findOne(wbDoc, 'sheets') || { children: [] }, 'sheet')) {
    const name = sh.attrs.name || `Sheet${sheets.length + 1}`
    const rid = sh.attrs['r:id'] || sh.attrs.Id
    const target = rid && rels.get(rid)?.target
    const path = resolvePart(target, 'xl/')
    const sheetXml = path && zipObj.getText(path)
    if (sheetXml === undefined) { sheets.push({ name, rows: [] }); continue }
    sheets.push({ name, rows: parseSheet(zipObj, path, sheetXml, shared) })
  }
  const meta = readCoreProps(zipObj)
  return { kind: 'workbook', meta, sheets: sheets.length ? sheets : [{ name: 'Sheet1', rows: [] }] }
}

function resolvePart(target, base) {
  if (!target) return undefined
  let p = String(target).replace(/\\/g, '/')
  if (p.startsWith('/')) return p.slice(1)
  if (p.startsWith('../')) {
    const parts = base.split('/').filter(Boolean)
    while (p.startsWith('../')) { p = p.slice(3); parts.pop() }
    return [...parts, p].join('/')
  }
  return base + p
}

function parseSheet(zipObj, path, sheetXml, shared) {
  const doc = parseXML(sheetXml)
  const dateXfs = collectDateStyleIndexes(zipObj.getText('xl/styles.xml'))
  const sheetData = findOne(doc, 'sheetData') || { children: [] }
  const rowsMap = new Map() // rowIdx -> { colIdx: cell }
  let maxRow = -1
  let maxCol = -1
  let autoRow = 0
  for (const row of children(sheetData, 'row')) {
    const rowIdx = (row.attrs.r ? Number(row.attrs.r) : ++autoRow) - 1
    autoRow = Math.max(autoRow, rowIdx + 1)
    let colCursor = 0
    const map = rowsMap.get(rowIdx) || {}
    for (const c of children(row, 'c')) {
      let col = colCursor
      if (c.attrs.r) { const p = parseCellRef(c.attrs.r); if (p) col = p.col }
      map[col] = parseCell(c, shared, dateXfs)
      colCursor = col + 1
      if (rowIdx > maxRow) maxRow = rowIdx
      if (col > maxCol) maxCol = col
    }
    rowsMap.set(rowIdx, map)
  }
  const rows = []
  for (let r = 0; r <= maxRow; r++) {
    const map = rowsMap.get(r) || {}
    const out = []
    for (let c = 0; c <= maxCol; c++) out.push(map[c] ?? { v: '', t: 's' })
    trimRow(out)
    rows.push(out)
  }
  return rows
}

function trimRow(row) {
  while (row.length && (row[row.length - 1].v === '' || row[row.length - 1].v === undefined)) row.pop()
}

function parseCell(c, shared, dateXfs, sheetRels) {
  const t = c.attrs.t || 'n'
  const s = Number(c.attrs.s || 0)
  const vNode = findOne(c, 'v')
  const fNode = findOne(c, 'f')
  const formula = fNode ? textOf(fNode) : undefined
  if (t === 'inlineStr') {
    const is = findOne(c, 'is')
    return { v: is ? readRichText(is, shared) : '', t: 's', f: formula }
  }
  if (t === 's') {
    const idx = Number(textOf(vNode || { children: [] }))
    return { v: shared[idx] ?? '', t: 's', f: formula }
  }
  if (t === 'str') return { v: textOf(vNode || { children: [] }), t: 's', f: formula }
  if (t === 'b') return { v: textOf(vNode) === '1' ? 'TRUE' : 'FALSE', t: 'b', f: formula }
  if (t === 'e') return { v: textOf(vNode) || '#ERR', t: 'e', f: formula }
  const raw = vNode ? textOf(vNode) : ''
  if (raw === '') return { v: formula ? `=${formula}` : '', t: 's', f: formula }
  const num = Number(raw)
  if (Number.isFinite(num)) {
    if (dateXfs.has(s)) {
      const d = serialToDate(num)
      if (d) return { v: d, t: 'd', f: formula }
    }
    return { v: num, t: 'n', f: formula }
  }
  return { v: raw, t: 's', f: formula }
}

function readRichText(isNode) {
  const parts = []
  for (const t of findAll(isNode, 't')) parts.push(textOf(t))
  return parts.join('')
}

function readSharedStrings(zipObj) {
  const xml = zipObj.getText('xl/sharedStrings.xml')
  if (!xml) return []
  const doc = parseXML(xml)
  return children(doc, 'si').map(si => {
    const ts = findAll(si, 't')
    return ts.map(t => textOf(t)).join('')
  })
}

function parseRelationships(xml) {
  const map = new Map()
  if (!xml) return map
  for (const r of findAll(parseXML(xml), 'Relationship')) map.set(r.attrs.Id, { type: r.attrs.Type, target: r.attrs.Target })
  return map
}

function readCoreProps(zipObj) {
  const meta = {}
  const core = zipObj.getText('docProps/core.xml')
  if (!core) return meta
  const d = parseXML(core)
  const get = name => { const n = findOne(d, name); return n ? textOf(n) : undefined }
  for (const [k, v] of Object.entries({ title: get('dc:title'), author: get('dc:creator'), created: get('dcterms:created'), modified: get('dcterms:modified') })) if (v) meta[k] = v
  return meta
}

// ---------- writing ----------

export function writeXlsx(wb) {
  const w = { kind: 'workbook', meta: {}, sheets: [], ...wb }
  const sheetEntries = []
  const wbRels = []
  const overrides = []
  w.sheets.forEach((sh, idx) => {
    const num = idx + 1
    sheetEntries.push({ name: `xl/worksheets/sheet${num}.xml`, data: sheetXml(sh) })
    wbRels.push(`<Relationship Id="rId${num}" Type="${R_NS}/worksheet" Target="worksheets/sheet${num}.xml"/>`)
    overrides.push(`<Override PartName="/xl/worksheets/sheet${num}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
  })
  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${SPREAD_NS}" xmlns:r="${R_NS}"><sheets>
${w.sheets.map((sh, i) => `<sheet name="${encodeEntities(uniqueName(sh.name, w.sheets, i), true)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}
</sheets></workbook>`

  return makeZip([
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${overrides.join('\n')}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>` },
    { name: 'xl/workbook.xml', data: wbXml },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wbRels.join('\n')}
<Relationship Id="rId${w.sheets.length + 1}" Type="${R_NS}/styles" Target="styles.xml"/>
</Relationships>` },
    { name: 'xl/styles.xml', data: STYLES_XML },
    ...sheetEntries,
    { name: 'docProps/core.xml', data: coreXml(w.meta) },
  ])
}

function uniqueName(name, all, self) {  let base = String(name || `Sheet${self + 1}`).replace(/[\\/*?:[\]]/g, '_').slice(0, 31) || `Sheet${self + 1}`
  let out = base
  let n = 1
  while (all.some((s, i) => i !== self && String(s.name).toLowerCase() === out.toLowerCase())) out = `${base.slice(0, 28)}(${n++})`
  return out
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${SPREAD_NS}">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`

function sheetXml(sh) {
  const prepared = withHeaderRow(sh)
  const rows = prepared.rows || []
  let cols = ''
  if (prepared.columns?.length) {
    cols = `<cols>${prepared.columns.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${Number(c.width) || autoWidth(c.title ?? (rows[0] || [])[i])}" customWidth="1"/>`).join('')}</cols>`
  }
  let body = ''
  rows.forEach((r, ri) => {
    let cells = ''
    let any = false
    r.forEach((c, ci) => {
      const cell = normalizeCell(c)
      if (cell.v === '' && !cell.f) return
      any = true
      cells += cellXml(colName(ci) + (ri + 1), cell, ri === 0 && prepared.__headerPrepended)
    })
    if (any || ri === 0) body += `<row r="${ri + 1}">${cells}</row>`
  })
  const dim = rows.length ? `<dimension ref="A1:${colName(Math.max(1, ...rows.map(r => r.length)))}${rows.length}"/>` : '<dimension ref="A1"/>'
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${SPREAD_NS}" xmlns:r="${R_NS}">${dim}<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${body}</sheetData></worksheet>`
}

function normalizeCell(c) {
  if (c === null || c === undefined) return { v: '', t: 's' }
  if (typeof c === 'object') return { v: c.v ?? '', t: c.t ?? (typeof c.v === 'number' ? 'n' : 's'), f: c.f, style: c.style }
  if (typeof c === 'number') return { v: c, t: 'n' }
  return { v: String(c), t: 's' }
}

function cellXml(ref, cell, isHeader = false) {
  const style = cell.style !== undefined ? cell.style : isHeader ? 1 : undefined
  const styleAttr = style !== undefined ? ` s="${style}"` : ''
  if (cell.f) {
    return `<c r="${ref}"${styleAttr}><f>${encodeEntities(cell.f)}</f>${cell.v !== '' ? `<v>${encodeEntities(String(cell.v))}</v>` : ''}</c>`
  }
  if (cell.t === 'n' && Number.isFinite(Number(cell.v))) return `<c r="${ref}"${styleAttr}><v>${Number(cell.v)}</v></c>`
  if (cell.t === 'b') return `<c r="${ref}"${styleAttr} t="b"><v>${cell.v === 'TRUE' || cell.v === true || cell.v === '1' ? 1 : 0}</v></c>`
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${encodeEntities(String(cell.v))}</t></is></c>`
}

/**
 * `columns` declares the sheet's header: when the first row does not already
 * carry exactly those titles, prepend a styled header row so callers get the
 * header they asked for in the sheet itself (not just as column widths).
 */
function withHeaderRow(sh) {
  const titles = (sh.columns || []).map(c => String((c && typeof c === 'object' ? c.title : c) ?? ''))
  if (!titles.length) return sh
  const rows = sh.rows || []
  const first = (rows[0] || []).map(c => String((c && typeof c === 'object' ? c.v : c) ?? ''))
  const already = first.length >= titles.length && titles.every((t, i) => first[i] === t)
  if (already) return sh
  return { ...sh, rows: [titles.map(t => ({ v: t, t: 's' })), ...rows], __headerPrepended: true }
}

function autoWidth(title) {
  const s = String(title ?? '')
  let w = 10
  for (const ch of s) w += ch.codePointAt(0) > 0x2e80 ? 2 : 1
  return Math.max(10, Math.min(60, w + 2))
}

function coreXml(meta) {
  const now = new Date().toISOString().replace(/\.\d+Z/, 'Z')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${meta.title ? `<dc:title>${encodeEntities(meta.title, true)}</dc:title>` : ''}
${meta.author ? `<dc:creator>${encodeEntities(meta.author, true)}</dc:creator>` : ''}
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`
}
