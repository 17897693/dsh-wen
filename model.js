// The shared office content model: format sniffing, text decoding, CSV,
// markdown <-> model, and helpers every reader/writer uses.

// ---------- format sniffing ----------

export function sniff(buf, ext) {
  const e = (ext || '').toLowerCase().replace(/^\./, '')
  if (buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b) return 'zip'
  if (buf.length >= 8 && buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0) return 'ole2'
  if (buf.length >= 5 && buf.subarray(0, 5).toString('latin1') === '%PDF-') return 'pdf'
  if (e === 'csv' || e === 'tsv') return 'delimited'
  if (e === 'rtf') return 'rtf'
  return 'text'
}

/** Extension → canonical family kind. */
export const KIND_BY_EXT = {
  docx: 'docx', docm: 'docx', dotx: 'docx', dotm: 'docx',
  xlsx: 'xlsx', xlsm: 'xlsx', xltx: 'xlsx', xltm: 'xlsx',
  pptx: 'pptx', pptm: 'pptx', potx: 'pptx', ppsx: 'pptx',
  odt: 'odt', ods: 'ods', odp: 'odp',
  pdf: 'pdf',
  doc: 'doc', wps: 'doc',
  xls: 'xls', et: 'xls',
  ppt: 'ppt', dps: 'ppt', pps: 'ppt', pot: 'ppt',
  csv: 'csv', tsv: 'tsv',
  md: 'md', markdown: 'md',
  txt: 'txt', text: 'txt',
  json: 'json', jsonl: 'jsonl',
  rtf: 'rtf',
}

/** Resolve which handler kind applies: sniff first, extension ties in. */
export function resolveKind(buf, ext, zipNames = null) {
  const e = (ext || '').toLowerCase().replace(/^\./, '')
  const byExt = KIND_BY_EXT[e]
  const container = sniff(buf, e)
  if (container === 'zip') {
    const has = p => zipNames.some(n => n.startsWith(p))
    if (has('word/')) return 'docx'
    if (has('xl/')) return 'xlsx'
    if (has('ppt/')) return 'pptx'
    if (zipNames.includes('mimetype') || has('META-INF/')) {
      const mime = zipNames.includes('mimetype') ? '' : ''
      if (byExt === 'odt' || byExt === 'ods' || byExt === 'odp') return byExt
      return 'odt'
    }
    return byExt && ['docx', 'xlsx', 'pptx'].includes(byExt) ? byExt : 'ooxml-unknown'
  }
  if (container === 'ole2') {
    if (byExt === 'et' || byExt === 'xls') return 'xls'
    if (byExt === 'dps' || byExt === 'ppt' || byExt === 'pps' || byExt === 'pot') return 'ppt'
    if (byExt === 'wps' || byExt === 'doc') return 'doc'
    return 'ole2'
  }
  if (container === 'pdf') return 'pdf'
  if (container === 'delimited') return byExt === 'tsv' ? 'tsv' : 'csv'
  if (container === 'rtf') return 'rtf'
  if (byExt) return byExt
  return 'text'
}

// ---------- text decoding (UTF-8 / UTF-16 / GBK-aware) ----------

const utf8Validator = new TextDecoder('utf-8', { fatal: true })

export function decodeTextBytes(buf) {
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return new TextDecoder('utf-16le').decode(buf.subarray(2))
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) return new TextDecoder('utf-16be').decode(buf.subarray(2))
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.subarray(3).toString('utf8')
  try { utf8Validator.decode(buf); return buf.toString('utf8') } catch {}
  for (const label of ['gb18030', 'big5', 'shift-jis', 'windows-1252']) {
    try {
      const s = new TextDecoder(label).decode(buf)
      if (!s.includes('\ufffd')) return s
    } catch {}
  }
  return buf.toString('latin1')
}

// ---------- CSV / TSV ----------

export function parseDelimited(text, delim) {
  if (!delim) delim = guessDelimiter(text)
  const rows = []
  let row = []
  let field = ''
  let quoted = false
  const pushField = () => { row.push(field); field = '' }
  const pushRow = () => { pushField(); rows.push(row); row = [] }
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++ } else quoted = false
      } else field += c
    } else if (c === '"') quoted = true
    else if (c === delim) pushField()
    else if (c === '\n') pushRow()
    else if (c === '\r') { if (text[i + 1] !== '\n') pushRow() }
    else field += c
  }
  if (field.length || row.length) pushRow()
  return rows.filter(r => !(r.length === 1 && r[0] === ''))
}

export function guessDelimiter(text) {
  const head = text.slice(0, 4096)
  const counts = { ',': 0, '\t': 0, ';': 0 }
  let quoted = false
  for (const c of head) {
    if (c === '"') quoted = !quoted
    if (!quoted && counts[c] !== undefined) counts[c]++
  }
  let best = ','
  for (const d of Object.keys(counts)) if (counts[d] > counts[best]) best = d
  return best
}

export function renderDelimited(rows, delim = ',') {
  const esc = v => {
    const s = v == null ? '' : String(v)
    return /[",\r\n]|^\s|\s$/.test(s) && delim === ',' ? `"${s.replace(/"/g, '""')}"` : s.replace(new RegExp(`[${delim}"\\r\\n]`), m => ({ '\r': '', '\n': ' ' }[m] ?? m))
  }
  return rows.map(r => r.map(esc).join(delim)).join('\r\n') + (rows.length ? '\r\n' : '')
}

// ---------- markdown -> document model ----------

export function markdownToDocument(md) {
  const blocks = []
  const lines = md.replace(/\r\n?/g, '\n').split('\n')
  let paragraph = []
  let i = 0
  const flushParagraph = () => {
    if (!paragraph.length) return
    blocks.push({ type: 'paragraph', runs: parseInlineRuns(paragraph.join(' ').trim()) })
    paragraph = []
  }
  const isTableLine = s => /^\s*\|.*\|\s*$/.test(s)
  // R19 任务 D：原判据是 `/^\s*([-*_])( *\1){2,}\s*$/` —— 复合体量词 `( *\1){2,}` 在
  // "单行 2 MiB 的 `-`"（**合法 markdown**）上抛 `RangeError: Maximum call stack size exceeded`。
  // 手写扫描逐条对齐：去掉首尾空白后，第一个字符必须是 `-` / `*` / `_`，之后只允许**空格**
  // 或同一个字符，且该字符出现 ≥3 次。
  const isHrLine = s => {
    const t = s.trim()
    if (t.length < 3) return false
    const c = t[0]
    if (c !== '-' && c !== '*' && c !== '_') return false
    let count = 0
    for (let k = 0; k < t.length; k++) {
      const ch = t[k]
      if (ch === c) { count++; continue }
      if (ch === ' ') continue
      return false
    }
    return count >= 3
  }
  while (i < lines.length) {
    const line = lines[i]
    if (/^\s*$/.test(line)) { flushParagraph(); i++; continue }
    let m
    if ((m = /^<pagebreak\s*\/?>\s*$/i.exec(line)) || /^\\newpage\s*$/i.test(line) || /^---page---\s*$/i.test(line)) {
      flushParagraph(); blocks.push({ type: 'pagebreak' }); i++; continue
    }
    if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) {
      flushParagraph(); blocks.push({ type: 'heading', level: m[1].length, text: stripInline(m[2]) }); i++; continue
    }
    if (isHrLine(line)) { flushParagraph(); blocks.push({ type: 'hr' }); i++; continue }
    if (/^```/.test(line)) {
      flushParagraph()
      const code = []
      i++
      while (i < lines.length && !/^```/.test(lines[i])) { code.push(lines[i]); i++ }
      i++
      blocks.push({ type: 'code', text: code.join('\n') })
      continue
    }
    if (isTableLine(line) && i + 1 < lines.length && /^\s*\|?[\s:|-]+\|?\s*$/.test(lines[i + 1]) && lines[i + 1].includes('-')) {
      flushParagraph()
      const rows = [splitRow(line)]
      const headerAlign = lines[i + 1]
      i += 2
      while (i < lines.length && isTableLine(lines[i])) { rows.push(splitRow(lines[i])); i++ }
      blocks.push({ type: 'table', header: headerAlign.includes('-'), rows })
      continue
    }
    if (/^\s*>/.test(line)) {
      flushParagraph()
      const quote = []
      while (i < lines.length && /^\s*>/.test(lines[i])) { quote.push(lines[i].replace(/^\s*>\s?/, '')); i++ }
      blocks.push({ type: 'quote', text: quote.join('\n') })
      continue
    }
    // 独立成行的图片 `![alt](path)` → image 块（第十二轮需求 1a）。
    // 这是 markdown 里图片的**块级**写法，走 markdownToDocument 的段落通道会被
    // parseInlineRuns 当成「`!` + 链接」而丢掉图片语义（旧版就是这样）。
    if ((m = /^\s*!\[([^\]]*)\]\(\s*(\S*?)\s*(?:"[^"]*"|'[^']*')?\s*\)\s*$/.exec(line))) {
      flushParagraph()
      blocks.push({ type: 'image', alt: m[1] || '图片', name: m[2] || '' })
      i++
      continue
    }
    const bullet = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(line)
    if (bullet) {
      flushParagraph()
      const ordered = /\d/.test(bullet[2][0])
      const list = { type: 'list', ordered, items: [] }
      while (i < lines.length) {
        const bm = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(lines[i])
        if (!bm || (/\d/.test(bm[2][0])) !== ordered) break
        const level = Math.min(8, Math.floor(bm[1].replace(/\t/g, '    ').length / 2))
        let text = bm[3]
        i++
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !/^(\s*)([-*+]|\d+[.)])\s+/.test(lines[i])) { text += ' ' + lines[i].trim(); i++ }
        list.items.push({ level, text: stripInline(text) })
      }
      blocks.push(list)
      continue
    }
    paragraph.push(line.trim())
    i++
  }
  flushParagraph()
  return { kind: 'document', meta: {}, blocks }
}

function splitRow(line) {
  let s = line.trim().replace(/^\|/, '').replace(/\|$/, '')
  const cells = []
  let cur = ''
  let escaped = false
  for (const c of s) {
    if (escaped) { cur += c; escaped = false; continue }
    if (c === '\\') { escaped = true; continue }
    if (c === '|') { cells.push(cur.trim()); cur = ''; continue }
    cur += c
  }
  cells.push(cur.trim())
  return cells
}

export function stripInline(s) {
  return s
    .replace(/\*\*\*([^*]+)\*\*\*/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*\n]+)\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/~~([^~]+)~~/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1')
    .trim()
}

export function parseInlineRuns(s) {
  const runs = []
  let cur = { text: '' }
  const flush = () => { if (cur.text) runs.push(cur); cur = { text: '' } }
  let i = 0
  while (i < s.length) {
    const rest = s.slice(i)
    let m
    if ((m = /^\*\*\*([^*]+)\*\*\*/.exec(rest))) { flush(); runs.push({ text: m[1], bold: true, italic: true }); i += m[0].length; continue }
    if ((m = /^\*\*([^*]+)\*\*/.exec(rest))) { flush(); runs.push({ text: m[1], bold: true }); i += m[0].length; continue }
    if ((m = /^\*([^*\n]+)\*/.exec(rest))) { flush(); runs.push({ text: m[1], italic: true }); i += m[0].length; continue }
    if ((m = /^__([^_]+)__/.exec(rest))) { flush(); runs.push({ text: m[1], underline: true }); i += m[0].length; continue }
    if ((m = /^~~([^~]+)~~/.exec(rest))) { flush(); runs.push({ text: m[1], strike: true }); i += m[0].length; continue }
    if ((m = /^`([^`]+)`/.exec(rest))) { flush(); runs.push({ text: m[1], code: true }); i += m[0].length; continue }
    if ((m = /^\[([^\]]+)\]\(([^)\s]+)\)/.exec(rest))) { flush(); runs.push({ text: m[1], link: m[2] }); i += m[0].length; continue }
    cur.text += s[i++]
  }
  flush()
  return runs.length ? runs : [{ text: s }]
}

// ---------- document model -> markdown ----------

export function documentToMarkdown(doc) {
  const out = []
  for (const b of doc.blocks || []) {
    switch (b.type) {
      case 'heading': out.push(`${'#'.repeat(Math.min(6, b.level || 1))} ${b.text}\n`); break
      case 'paragraph': out.push(`${(b.runs || []).map(runMd).join('')}\n`); break
      case 'list': {
        for (const it of b.items || []) {
          const indent = '  '.repeat(Math.max(0, Math.min(8, it.level || 0)))
          out.push(b.ordered ? `${indent}1. ${it.text}` : `${indent}- ${it.text}`)
        }
        out.push('')
        break
      }
      case 'table': out.push(tableMd(b.rows, b.header)); break
      case 'code': out.push('```\n' + b.text + '\n```\n'); break
      case 'quote': out.push(String(b.text || '').split('\n').map(l => `> ${l}`).join('\n') + '\n'); break
      case 'hr': out.push('---\n'); break
      case 'pagebreak': out.push('<!-- pagebreak -->\n'); break
      case 'image': out.push(`![${b.alt || '图片'}](${b.name || ''})\n`); break
      default: if (b.text) out.push(`${b.text}\n`)
    }
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

function runMd(r) {
  let t = r.text
  if (r.code) t = `\`${t}\``
  if (r.bold) t = `**${t}**`
  if (r.italic) t = `*${t}*`
  if (r.strike) t = `~~${t}~~`
  if (r.link) t = `[${t}](${r.link})`
  return t
}

export function tableMd(rows, header = true) {
  if (!rows.length) return ''
  const w = Math.max(...rows.map(r => r.length))
  const grid = rows.map(r => Array.from({ length: w }, (_, i) => String(r[i] ?? '').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')))
  const lines = [`| ${grid[0].join(' | ')} |`, `|${grid[0].map(() => ' --- ').join('|')}|`]
  if (!header) lines.splice(1, 0, ...[])
  for (let i = 1; i < grid.length; i++) lines.push(`| ${grid[i].join(' | ')} |`)
  return lines.join('\n') + '\n'
}

// ---------- generic model coercion for writers ----------

/** Accept structured or markdown input and return a normalized model of the wanted kind. */
export function coerceModel(input, want) {
  const { document, slides, workbook, markdown, table } = input
  if (want === 'document') {
    if (document) return normalizeDocument(document)
    if (typeof markdown === 'string') return markdownToDocument(markdown)
    if (slides) return slidesToDocument(slides)
    if (workbook) return workbookToDocument(workbook)
    if (table) return tableToDocument(table)
    return null
  }
  if (want === 'slides') {
    if (slides) return normalizeSlides(slides)
    if (typeof markdown === 'string') return markdownToSlides(markdown)
    if (document) return documentToSlides(document)
    return null
  }
  if (want === 'workbook') {
    if (workbook) return normalizeWorkbook(workbook)
    if (table) return normalizeWorkbook({ sheets: [{ name: 'Sheet1', rows: table.rows, columns: table.columns }] })
    if (typeof markdown === 'string') {
      const doc = markdownToDocument(markdown)
      const tables = doc.blocks.filter(b => b.type === 'table')
      if (tables.length) return normalizeWorkbook({ sheets: tables.map((t, i) => ({ name: `Sheet${i + 1}`, rows: t.rows })) })
    }
    if (document) return documentToWorkbook(document)
    return null
  }
  throw new Error(`unknown model kind "${want}"`)
}

export function normalizeDocument(d) {
  return { kind: 'document', meta: d.meta || {}, blocks: (d.blocks || []).map(normalizeBlock) }
}

function normalizeBlock(b) {
  if (b.type === 'paragraph' || b.type === 'quote') {
    const runs = b.runs ?? [{ text: String(b.text ?? '') }]
    return { ...b, runs: runs.map(r => typeof r === 'string' ? { text: r } : r) }
  }
  if (b.type === 'list') return { ...b, items: (b.items || []).map(it => typeof it === 'string' ? { text: it, level: 0 } : { level: 0, ...it }) }
  if (b.type === 'table') return { ...b, rows: (b.rows || []).map(r => r.map(c => Array.isArray(c) ? String(c[0] ?? '') : typeof c === 'object' ? String(c.v ?? '') : String(c ?? ''))) }
  return { ...b }
}

export function normalizeSlides(s) {
  return {
    kind: 'slides', meta: s.meta || {},
    slides: (s.slides || []).map(sl => ({
      layout: sl.layout || (sl.subtitle !== undefined ? 'title' : sl.title ? 'content' : 'blank'),
      title: sl.title ?? '', subtitle: sl.subtitle,
      bullets: (sl.bullets || []).map(b => typeof b === 'string' ? { text: b, level: 0 } : { level: 0, ...b }),
      table: sl.table, notes: sl.notes,
    })),
  }
}

export function normalizeWorkbook(w) {
  return {
    kind: 'workbook', meta: w.meta || {},
    sheets: (w.sheets || []).map(sh => ({
      name: sh.name || 'Sheet',
      columns: sh.columns || null,
      rows: (sh.rows || []).map(r => r.map(c => normalizeCell(c))),
    })),
  }
}

export function normalizeCell(c) {
  if (c === null || c === undefined) return { v: '', t: 's' }
  if (typeof c === 'number') return { v: c, t: 'n' }
  if (typeof c === 'boolean') return { v: c ? '1' : '0', t: 'b' }
  if (typeof c === 'object') {
    const t = c.t ?? (typeof c.v === 'number' ? 'n' : 's')
    return { v: c.v ?? '', t, f: c.f, style: c.style }
  }
  const s = String(c)
  if (s !== '' && /^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(s.trim()) && !/^0\d/.test(s.trim())) {
    return { v: Number(s), t: 'n' }
  }
  return { v: s, t: 's' }
}

// ---------- cross-kind conversions ----------

export function slidesToDocument(sl) {
  const doc = normalizeSlides(sl)
  const blocks = []
  for (const s of doc.slides) {
    if (s.title) blocks.push({ type: 'heading', level: 2, text: s.title })
    if (s.subtitle) blocks.push({ type: 'paragraph', runs: [{ text: s.subtitle }] })
    for (const b of s.bullets) blocks.push({ type: 'paragraph', runs: [{ text: '• ' + b.text }], indent: b.level })
    if (s.table) blocks.push({ type: 'table', rows: s.table.rows || s.table, header: true })
    if (s.notes) blocks.push({ type: 'paragraph', runs: [{ text: `备注: ${s.notes}`, italic: true }] })
  }
  return { kind: 'document', meta: doc.meta, blocks }
}

export function documentToSlides(doc) {
  const d = normalizeDocument(doc)
  const slides = []
  let current = null
  const ensure = title => {
    if (!current) { current = { title, bullets: [] }; slides.push(current) }
    return current
  }
  for (const b of d.blocks) {
    if (b.type === 'heading' && b.level <= 2) {
      current = { title: b.text, bullets: [] }
      slides.push(current)
    } else if (b.type === 'hr') current = null
    else if (b.type === 'paragraph') ensure('').bullets.push({ text: plainOf(b.runs), level: 0 })
    else if (b.type === 'list') for (const it of b.items) ensure('').bullets.push(it)
    else if (b.type === 'table') ensure('').table = { rows: b.rows }
  }
  if (!slides.length) slides.push({ title: '演示文稿', bullets: [] })
  slides[0] = { ...slides[0], layout: 'title', subtitle: slides[0].subtitle }
  return { kind: 'slides', meta: d.meta, slides: slides.map(s => ({ layout: s.title && slides.indexOf(s) === 0 ? 'title' : s.title ? 'content' : 'blank', ...s })) }
}

export function documentToWorkbook(doc) {
  const d = normalizeDocument(doc)
  const tables = d.blocks.filter(b => b.type === 'table')
  const sheets = tables.length
    ? tables.map((t, i) => ({ name: `表${i + 1}`, rows: t.rows }))
    : [{ name: 'Sheet1', rows: d.blocks.filter(b => b.type === 'paragraph' || b.type === 'heading').map(b => [plainOf(b.runs) || b.text || '']) }]
  return normalizeWorkbook({ meta: d.meta, sheets })
}

export function workbookToDocument(wb) {
  const w = normalizeWorkbook(wb)
  const blocks = []
  for (const sh of w.sheets) {
    blocks.push({ type: 'heading', level: 2, text: sh.name })
    if (sh.rows.length) {
      const rows = sh.rows.map(r => r.map(c => typeof c === 'object' ? String(c.v ?? '') : String(c)))
      blocks.push({ type: 'table', header: true, rows })
    }
  }
  return { kind: 'document', meta: w.meta, blocks }
}

export function tableToDocument(t) {
  return { kind: 'document', meta: {}, blocks: [{ type: 'table', header: true, rows: (t.rows || []).map(r => r.map(String)) }] }
}

export function plainOf(runs) {
  return (runs || []).map(r => (typeof r === 'string' ? r : r.text)).join('')
}

// ---------- markdown -> slides (H1/H2 = new slide) ----------

export function markdownToSlides(md) {
  const doc = markdownToDocument(md)
  const slides = []
  let current = null
  let titleSlideDone = false
  for (const b of doc.blocks) {
    if (b.type === 'heading' && (b.level === 1 || b.level === 2) && (b.level === 1 || current)) {
      if (b.level === 1 && !titleSlideDone && slides.length === 0) {
        current = { layout: 'title', title: b.text, subtitle: '', bullets: [] }
        slides.push(current)
        titleSlideDone = true
        continue
      }
      current = { layout: 'content', title: b.text, bullets: [] }
      slides.push(current)
      continue
    }
    if (!current) { current = { layout: 'content', title: doc.meta?.title || '内容', bullets: [] }; slides.push(current) }
    if (b.type === 'paragraph') current.bullets.push({ text: plainOf(b.runs), level: 0 })
    else if (b.type === 'list') for (const it of b.items) current.bullets.push(it)
    else if (b.type === 'table') current.table = { rows: b.rows }
    else if (b.type === 'heading') current.bullets.push({ text: `${'#'.repeat(b.level)} ${b.text}`, level: 0 })
  }
  if (!slides.length) slides.push({ layout: 'title', title: '演示文稿', subtitle: '', bullets: [] })
  return { kind: 'slides', meta: {}, slides }
}

// ---------- workbook -> csv rows (read side helper) ----------

export function workbookToCsvRows(wb, sheetName) {
  const w = normalizeWorkbook(wb)
  const sheet = sheetName ? w.sheets.find(s => s.name === sheetName) || w.sheets[0] : w.sheets[0]
  if (!sheet) return []
  return sheet.rows.map(r => r.map(c => {
    if (c === null || c === undefined) return ''
    return typeof c === 'object' ? String(c.v ?? '') : String(c)
  }))
}
