// OpenDocument (.odt/.ods/.odp) reader/writer.
import { openZip, makeZip } from './zip.js'
import { parseXML, serializeXML, findAll, children, textOf, encodeEntities } from './xml.js'

function findOne(node, name) { return findAll(node, name)[0] }

const REPETITION_CAP = 400
const rep = n => Math.min(Number(n || 1) || 1, REPETITION_CAP)

// ---------- reading ----------

function openIfNeeded(zipInput) {
  return typeof zipInput === 'object' && zipInput.names ? zipInput : openZip(zipInput)
}

function contentDoc(zipObj) {
  const xml = zipObj.getText('content.xml')
  if (xml === undefined) throw new Error('odf: content.xml missing')
  return parseXML(xml)
}

function plain(node) {
  let s = ''
  const walk = n => {
    if (typeof n === 'string') { s += n; return }
    if (n.raw) return
    if (n.cdata !== undefined) { s += n.cdata; return }
    if (n.name === 'text:tab') { s += '\t'; return }
    if (n.name === 'text:s') { s += ' '.repeat(Math.min(64, Number(n.attrs['text:c'] || 1))); return }
    if (n.name === 'text:line-break') { s += '\n'; return }
    for (const c of n.children) walk(c)
  }
  walk(node)
  return decodeNumeric(s)
}

/** R19 任务 D：数字实体守卫（`&#` + 超长数字 ⇒ `Infinity` ⇒ `String.fromCodePoint` 抛
 *  `RangeError: Invalid code point`）。**本函数有真实调用点**（`elementText` 的 `decodeNumeric(s)`），
 *  所以这条是可达路径：含 `&#999…9;`（几百位）的 `.odt` 文本过去会让整条读取崩掉。 */
const codePointOrRaw = (all, code) => (Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all)
function decodeNumeric(s) {
  return s
    .replace(/&#(\d+);/g, (all, d) => codePointOrRaw(all, Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (all, h) => codePointOrRaw(all, parseInt(h, 16)))
}

function trimStrings(row) {
  while (row.length && row[row.length - 1] === '') row.pop()
}

export function readOdt(zipInput) {
  const zipObj = openIfNeeded(zipInput)
  const content = contentDoc(zipObj)
  const bodyNode = findOne(content, 'office:text') || { children: [] }
  const blocks = []
  const walk = (node, listCtx) => {
    if (typeof node === 'string' || node.raw || node.cdata) return
    const n = node
    switch (n.name) {
      case 'text:h': {
        const level = Math.min(6, Number(n.attrs['text:outline-level'] || 1))
        blocks.push({ type: 'heading', level, text: plain(n) })
        return
      }
      case 'text:p': {
        const t = plain(n)
        if (listCtx) listCtx.items.push({ text: t, level: listCtx.level })
        else blocks.push({ type: 'paragraph', runs: [{ text: t }] })
        return
      }
      case 'text:list': {
        const styleName = String(n.attrs['text:style-name'] || '')
        const ordered = /L2|Num|num/.test(styleName) && !/Bullet|bul|L1/.test(styleName)
        const items = []
        for (const c of children(n)) {
          if (c.name === 'text:list-item') for (const cc of children(c)) walk(cc, { items, level: 0 })
          else if (c.name === 'text:list') {
            for (const cc of children(c)) {
              if (cc.name === 'text:list-item') for (const pc of children(cc)) {
                if (pc.name === 'text:p') items.push({ text: plain(pc), level: 1 })
              }
            }
          }
        }
        blocks.push({ type: 'list', ordered, items: items.length ? items : [{ text: '', level: 0 }] })
        return
      }
      case 'text:table-of-content': return
      case 'table:table': {
        const rows = []
        for (const row of children(n, 'table:table-row')) {
          const cells = []
          const rRep = rep(row.attrs['table:number-rows-repeated'])
          for (const cell of children(row, 'table:table-cell')) {
            const cRep = rep(cell.attrs['table:number-columns-repeated'])
            let val = ''
            const officeText = findOne(cell, 'office:text')
            if (officeText) val = children(officeText, 'text:p').map(p => plain(p)).join('\n')
            else val = plain(findOne(cell, 'text:p') || { children: [] })
            const floatVal = cell.attrs['office:value']
            if (!val && floatVal !== undefined) val = String(floatVal)
            for (let i = 0; i < cRep && cells.length < 128; i++) cells.push(val)
          }
          trimStrings(cells)
          if (rRep > 1 && !cells.length) break
          for (let i = 0; i < rRep && rows.length < 4000; i++) rows.push([...cells])
        }
        while (rows.length && rows[rows.length - 1].length === 0) rows.pop()
        if (rows.length) blocks.push({ type: 'table', header: false, rows })
        return
      }
      case 'text:a': return
      default:
        for (const c of n.children) walk(c, listCtx)
    }
  }
  for (const c of bodyNode.children) walk(c, null)
  const meta = readMeta(zipObj)
  const doc = { kind: 'document', meta, blocks: blocks.filter(b => b.type !== 'paragraph' || (b.runs || []).some(r => r.text !== '')) }
  doc.blocks = doc.blocks.filter(b => b.type !== 'paragraph' || String(plainTextOf(b.runs)).trim() !== '')
  return doc
}

function plainTextOf(runs) {
  return (runs || []).map(r => (typeof r === 'string' ? r : r.text)).join('')
}

export function readOds(zipInput) {
  const zipObj = openIfNeeded(zipInput)
  const content = contentDoc(zipObj)
  const sheets = []
  for (const t of findAll(content, 'table:table')) {
    const rows = []
    for (const row of children(t, 'table:table-row')) {
      const cells = []
      const rRep = rep(row.attrs['table:number-rows-repeated'])
      for (const cell of children(row, 'table:table-cell')) {
        let model
        const vt = cell.attrs['office:value-type']
        const formula = findOne(cell, 'table:formula')?.attrs['table:string']
        if (vt === 'float' || vt === 'percentage' || vt === 'currency') {
          const num = Number(cell.attrs['office:value'])
          model = { v: Number.isFinite(num) ? num : (cell.attrs['office:value'] ?? ''), t: 'n', f: formula }
        } else if (vt === 'boolean') model = { v: cell.attrs['office:boolean-value'] === 'true' ? 'TRUE' : 'FALSE', t: 'b', f: formula }
        else if (vt === 'date') model = { v: cell.attrs['office:date-value'] || '', t: 's', f: formula }
        else if (vt === 'time') model = { v: cell.attrs['office:time-value'] || '', t: 's', f: formula }
        else {
          const text = children(cell, 'text:p').map(p => plain(p)).join('\n')
          model = { v: text, t: 's', f: formula }
        }
        const cRep = rep(cell.attrs['table:number-columns-repeated'])
        for (let i = 0; i < cRep && cells.length < 128; i++) cells.push({ ...model })
      }
      while (cells.length && cells[cells.length - 1].v === '') cells.pop()
      if (rRep > 1 && !cells.length) break
      for (let i = 0; i < rRep && rows.length < 4000; i++) rows.push(cells.map(c => ({ ...c })))
    }
    while (rows.length && rows[rows.length - 1].length === 0) rows.pop()
    sheets.push({ name: t.attrs['table:name'] || `Sheet${sheets.length + 1}`, rows })
  }
  return { kind: 'workbook', meta: readMeta(zipObj), sheets }
}

export function readOdp(zipInput) {
  const zipObj = openIfNeeded(zipInput)
  const content = contentDoc(zipObj)
  const slides = []
  const body = findOne(content, 'office:presentation')
  if (!body) return { kind: 'slides', meta: readMeta(zipObj), slides: [] }
  for (const page of children(body, 'draw:page')) {
    const slide = { layout: 'content', title: '', bullets: [] }
    for (const frame of children(page, 'draw:frame')) {
      const cls = String(frame.attrs['presentation:class'] || '')
      const textBox = findOne(frame, 'draw:text-box')
      if (!textBox) continue
      const paras = []
      const collect = n => {
        if (typeof n === 'string' || n.raw) return
        if (n.name === 'text:h') { const t = plain(n); if (t.trim()) paras.push({ text: t, level: 0, heading: true }); return }
        if (n.name === 'text:p') { const t = plain(n); if (t.trim()) paras.push({ text: t, level: 0 }); return }
        if (n.name === 'text:list') {
          const go = (list, level) => {
            for (const it of children(list, 'text:list-item')) {
              const ps = children(it, 'text:p').map(p => plain(p)).join(' ')
              if (ps.trim()) paras.push({ text: ps, level })
              for (const sub of children(it, 'text:list')) go(sub, level + 1)
            }
          }
          go(n, 0)
          return
        }
        for (const c of n.children) collect(c)
      }
      for (const c of textBox.children) collect(c)
      if (/^(title|subtitle)$/i.test(cls)) {
        const joined = paras.map(p => p.text).join('\n')
        if (/subtitle/i.test(cls)) slide.subtitle = joined
        else slide.title = joined
      } else {
        slide.bullets.push(...paras)
      }
    }
    if (!slide.title && !slide.bullets.length && !slide.subtitle) slide.layout = 'blank'
    slides.push(slide)
  }
  return { kind: 'slides', meta: readMeta(zipObj), slides }
}

function readMeta(zipObj) {
  const meta = {}
  const xml = zipObj.getText('meta.xml')
  if (!xml) return meta
  const d = parseXML(xml)
  const get = name => { const n = findOne(d, name); return n ? textOf(n) : undefined }
  for (const [k, v] of Object.entries({
    title: get('dc:title'), author: get('meta:initial-creator'), description: get('dc:description'),
    keywords: get('meta:keyword'), created: get('meta:creation-date'), modified: get('dc:date'),
  })) if (v) meta[k] = v
  return meta
}

// ---------- writing ----------

const OFFICE_OPEN = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" office:version="1.2">`
const OFFICE_CLOSE = '</office:document-content>'

const AUTOMATIC_STYLES = `<office:automatic-styles>
${[1, 2, 3, 4, 5, 6].map(i => `<style:style style:name="H${i}" style:family="paragraph" style:parent-style-name="Heading" style:class="text"><style:text-properties fo:font-size="${[24, 18, 14, 12, 11, 10][i - 1]}pt" fo:font-weight="bold" fo:color="#1f3864"/></style:style>`).join('\n')}
<style:style style:name="co1" style:family="paragraph" style:parent-style-name="Preformatted_20_Text"><style:paragraph-properties fo:background-color="#f5f5f5"/></style:style>
${[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(i => `<style:style style:name="co${i}" style:family="table-column"><style:table-column-properties style:column-width="3.5cm"/></style:style>`).join('\n')}
<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:padding="0.1cm" fo:border="0.06pt solid #000000"/></style:style>
<style:style style:name="cehead" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:padding="0.1cm" fo:border="0.06pt solid #000000" fo:background-color="#4472c4"/><style:text-properties fo:font-weight="bold" fo:color="#ffffff"/></style:style>
</office:automatic-styles>`

function escRuns(runs) {
  let out = ''
  for (const r of runs || []) {
    const run = typeof r === 'string' ? { text: r } : r
    let style = ''
    if (run.bold && run.italic) style = 'T3'
    else if (run.bold) style = 'T1'
    else if (run.italic) style = 'T2'
    else if (run.underline) style = 'T4'
    else if (run.strike) style = 'T5'
    else if (run.code) style = 'TT'
    const inner = encodeEntities(String(run.text ?? '')).replace(/\n/g, '<text:line-break/>')
    const body = style ? `<text:span text:style-name="${style}">${inner}</text:span>` : inner
    out += run.link ? `<text:a xlink:type="simple" xlink:href="${encodeEntities(run.link, true)}">${body}</text:a>` : body
  }
  return out
}

const odtTable = rows => {
  const cols = Math.max(1, ...rows.map(r => r.length))
  let xml = '<table:table table:name="Table1">'
  xml += Array.from({ length: Math.min(cols, 10) }, (_, i) => `<table:table-column table:style-name="co${(i % 10) + 1}"/>`).join('')
  if (cols > 10) xml += `<table:table-column table:style-name="co1" table:number-columns-repeated="${cols - 10}"/>`
  rows.forEach((r, ri) => {
    xml += '<table:table-row>'
    for (let c = 0; c < cols; c++) {
      const cellStyle = ri === 0 ? 'cehead' : 'ce1'
      xml += `<table:table-cell table:style-name="${cellStyle}" office:value-type="string"><office:text><text:p text:style-name="Table_20_1">${encodeEntities(String(r[c] ?? ''))}</text:p></office:text></table:table-cell>`
    }
    xml += '</table:table-row>'
  })
  return xml + '</table:table>'
}

export function writeOdt(doc) {
  const d = { kind: 'document', meta: {}, blocks: [], ...doc }
  const body = []
  for (const b of d.blocks) {
    if (b.type === 'heading') body.push(`<text:h text:style-name="H${Math.min(6, b.level || 1)}" text:outline-level="${Math.min(6, b.level || 1)}">${encodeEntities(b.text)}</text:h>`)
    else if (b.type === 'paragraph') body.push(`<text:p>${escRuns(b.runs)}</text:p>`)
    else if (b.type === 'quote') body.push(`<text:p text:style-name="Quot_20_Escape">${encodeEntities(String(b.text || ''))}</text:p>`)
    else if (b.type === 'code') body.push(String(b.text || '').split('\n').map(l => `<text:p text:style-name="co1">${encodeEntities(l || ' ')}</text:p>`).join('\n'))
    else if (b.type === 'list') {
      const items = (b.items || []).map(it => `<text:list-item><text:p>${encodeEntities(String(it.text ?? ''))}</text:p></text:list-item>`).join('')
      body.push(`<text:list text:style-name="L${b.ordered ? 2 : 1}">${items}</text:list>`)
    } else if (b.type === 'table') body.push(odtTable(b.rows || []))
    else if (b.type === 'pagebreak') body.push('<text:p><text:line-break/></text:p>')
    else if (b.type === 'hr') body.push('<text:p>————————————</text:p>')
  }
  const content = `${OFFICE_OPEN}${AUTOMATIC_STYLES}<office:body><office:text>${body.join('\n')}</office:text></office:body>${OFFICE_CLOSE}`
  return makeZip([
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true },
    { name: 'META-INF/manifest.xml', data: manifest('application/vnd.oasis.opendocument.text') },
    { name: 'content.xml', data: content },
    { name: 'styles.xml', data: STYLES_XML },
    { name: 'meta.xml', data: metaXml(d.meta) },
  ])
}

export function writeOds(wb) {
  const w = { kind: 'workbook', meta: {}, sheets: [{ name: 'Sheet1', rows: [] }], ...wb }
  const tables = w.sheets.map(sh => {
    const cols = Math.max(1, ...sh.rows.map(r => r.length))
    let xml = `<table:table table:name="${encodeEntities(sanitizeSheetName(sh.name), true)}">`
    xml += Array.from({ length: Math.min(cols, 3) }, () => '<table:table-column/>').join('')
    if (cols > 3) xml += `<table:table-column table:number-columns-repeated="${cols - 3}"/>`
    for (const r of sh.rows) {
      xml += '<table:table-row>'
      for (let c = 0; c < cols; c++) {
        const cell = r[c] ?? { v: '', t: 's' }
        const v = typeof cell === 'object' && cell !== null ? cell : { v: cell }
        if (v.f) xml += `<table:table-cell table:formula="${encodeEntities(String(v.f).startsWith('=') ? String(v.f) : '=' + String(v.f), true)}" office:value-type="string"><text:p>${encodeEntities(String(v.v ?? ''))}</text:p></table:table-cell>`
        else if (typeof v.v === 'number' && Number.isFinite(v.v)) xml += `<table:table-cell office:value-type="float" office:value="${v.v}"><text:p>${v.v}</text:p></table:table-cell>`
        else if (String(v.v ?? '') === '') xml += '<table:table-cell/>'
        else xml += `<table:table-cell office:value-type="string"><text:p>${encodeEntities(String(v.v))}</text:p></table:table-cell>`
      }
      xml += '</table:table-row>'
    }
    return xml + '</table:table>'
  })
  const content = `${OFFICE_OPEN}${AUTOMATIC_STYLES}<office:body><office:spreadsheet>${tables.join('\n')}</office:spreadsheet></office:body>${OFFICE_CLOSE}`
  return makeZip([
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', store: true },
    { name: 'META-INF/manifest.xml', data: manifest('application/vnd.oasis.opendocument.spreadsheet') },
    { name: 'content.xml', data: content },
    { name: 'styles.xml', data: STYLES_XML },
    { name: 'meta.xml', data: metaXml(w.meta) },
  ])
}

function sanitizeSheetName(name) {
  let out = String(name || 'Sheet').replace(/[.*#/\\]/g, '_').slice(0, 31) || 'Sheet'
  if (/^[aA][bB][cC]\d+$/.test(out)) out = `_${out}`
  return out
}

export function writeOdp(sl) {
  const s = { kind: 'slides', meta: {}, slides: [], ...sl }
  const pages = s.slides.map((sl2, i) => {
    const frames = []
    if (sl2.title) frames.push(`<draw:frame draw:layer="layout" svg:x="2cm" svg:y="1cm" svg:width="24cm" svg:height="3.8cm" presentation:class="title"><draw:text-box><text:p>${encodeEntities(String(sl2.title))}</text:p></draw:text-box></draw:frame>`)
    if (sl2.subtitle) frames.push(`<draw:frame draw:layer="layout" svg:x="2cm" svg:y="4.8cm" svg:width="24cm" svg:height="1.6cm" presentation:class="subtitle"><draw:text-box><text:p>${encodeEntities(String(sl2.subtitle))}</text:p></draw:text-box></draw:frame>`)
    const paras = []
    for (const b of sl2.bullets || []) paras.push(`<text:p>${encodeEntities(String(b.text ?? ''))}</text:p>`)
    if (sl2.table?.rows) paras.push(odtTable(sl2.table.rows.map(r => Array.isArray(r) ? r : [String(r)])))
    if (paras.length) frames.push(`<draw:frame draw:layer="layout" svg:x="2cm" svg:y="6cm" svg:width="24cm" svg:height="11cm" presentation:class="outline"><draw:text-box>${paras.join('\n')}</draw:text-box></draw:frame>`)
    return `<draw:page draw:name="page${i + 1}" draw:master-page-name="Default">${frames.join('\n')}</draw:page>`
  })
  const content = `${OFFICE_OPEN}${AUTOMATIC_STYLES}<office:body><office:presentation>${pages.join('\n')}</office:presentation></office:body>${OFFICE_CLOSE}`
  return makeZip([
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.presentation', store: true },
    { name: 'META-INF/manifest.xml', data: manifest('application/vnd.oasis.opendocument.presentation') },
    { name: 'content.xml', data: content },
    { name: 'styles.xml', data: STYLES_XML },
    { name: 'meta.xml', data: metaXml(s.meta) },
  ])
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.2">
<office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri"/><style:font-face style:name="宋体" svg:font-family="宋体"/><style:font-face style:name="Consolas" svg:font-family="Consolas"/></office:font-face-decls>
<office:styles>
<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.21cm"/><style:text-properties fo:font-size="12pt" style:font-name="Calibri" style:font-name-asian="宋体"/></style:default-style>
<style:style style:name="Standard" style:family="paragraph"/>
<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_Body"><style:paragraph-properties fo:margin-top="0.35cm" fo:margin-bottom="0.15cm"/><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>
<style:style style:name="Text_20_Body" style:family="paragraph" style:parent-style-name="Standard"/>
<style:style style:name="Preformatted_20_Text" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties style:font-name="Consolas" fo:font-size="10pt"/></style:style>
<style:style style:name="Quot_20_Escape" style:family="paragraph" style:parent-style-name="Text_20_Body"><style:paragraph-properties fo:margin-left="1cm"/><style:text-properties fo:font-style="italic"/></style:style>
<style:style style:name="Table_20_1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/></style:style>
<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
<style:style style:name="T3" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic"/></style:style>
<style:style style:name="T4" style:family="text"><style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
<style:style style:name="T5" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>
<style:style style:name="TT" style:family="text"><style:text-properties style:font-name="Consolas" fo:font-size="10pt"/></style:style>
<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:padding="0.05cm"/></style:style>
</office:styles>
<office:automatic-styles>
<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.699cm" style:num-format="1" fo:margin="2cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>
<style:page-layout style:name="Mpm1"><style:page-layout-properties fo:page-width="27.94cm" fo:page-height="17.78cm" style:num-format="1" fo:margin="0cm" style:print-orientation="landscape"/></style:page-layout>
</office:automatic-styles>
<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles>
</office:document-styles>`

function manifest(mime) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="${mime}"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`
}

function metaXml(meta) {
  const now = new Date().toISOString().replace(/\.\d+Z/, 'Z')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">
<office:meta>
${meta.title ? `<dc:title>${encodeEntities(meta.title, true)}</dc:title>` : ''}
${meta.author ? `<meta:initial-creator>${encodeEntities(meta.author, true)}</meta:initial-creator>` : ''}
${meta.description ? `<dc:description>${encodeEntities(meta.description, true)}</dc:description>` : ''}
${meta.keywords ? `<meta:keyword>${encodeEntities(meta.keywords, true)}</meta:keyword>` : ''}
<meta:creation-date>${now}</meta:creation-date>
<dc:date>${now}</dc:date>
</office:meta></office:document-meta>`
}

// ---------- editing ----------

export function replaceTextInOdfContent(contentXmlText, find, replace, { regex = false } = {}) {
  const doc = parseXML(contentXmlText)
  const ts = findAll(doc, ['text:p', 'text:h'])
  let count = 0
  let re = null
  if (regex) { try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`invalid regex: ${e.message}`) } }
  const patch = t => {
    const text = plain(t)
    if (re) { const hits = text.match(re); if (hits) { count += hits.length; t.children = [encodeEntities(text.replace(re, () => replace))] } }
    else if (text.includes(find)) { count += text.split(find).length - 1; t.children = [encodeEntities(text.split(find).join(replace))] }
  }
  for (const t of ts) patch(t)
  const cells = findAll(doc, 'table:table-cell').map(c => findOne(c, 'text:p')).filter(Boolean)
  for (const c of cells) patch(c)
  if (!count) return { xml: null, count }
  return { xml: serializeXML(doc), count }
}
