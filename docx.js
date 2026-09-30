// Word (.docx) reader/writer on the OOXML zip container.
import { openZip, makeZip } from './zip.js'
import { parseXML, serializeXML, findAll, children, textOf, encodeEntities } from './xml.js'
import { plainOf } from './model.js'
import { readImageBytes, sniffImage, imageExt, imageMime } from './image.js'
import { createHash } from 'node:crypto'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function findEl(node, name) {
  return findAll(node, name)[0]
}

/** R19 任务 D：数字实体必须守卫 —— `&#` + 400 位十进制数字会算出 `Infinity`，
 *  `String.fromCodePoint(Infinity)` 抛 `RangeError: Invalid code point`。
 *  本函数当前**零调用点**（`grep -n decodeEntitiesRaw docx.js` 只有定义这一行），
 *  守卫是为"将来接上调用点"不再踩；口径与 `xml.js::decodeEntities` 一致（越界/非法 → 原样保留）。 */
const codePointOrRaw = (all, code) => (Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all)
function decodeEntitiesRaw(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (all, d) => codePointOrRaw(all, Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (all, h) => codePointOrRaw(all, parseInt(h, 16)))
}

// ---------- reading ----------

export function readDocx(zipInput) {
  const zipObj = typeof zipInput === 'object' && zipInput.names ? zipInput : openZip(zipInput)
  const docXml = zipObj.getText('word/document.xml')
  if (docXml === undefined) throw new Error('docx: word/document.xml missing')
  const doc = parseXML(docXml)
  const rels = parseRelationships(zipObj.getText('word/_rels/document.xml.rels'))
  const body = findEl(doc, 'w:body') || { children: [] }
  const blocks = []
  for (const el of children(body)) {
    if (el.name === 'w:p') blocks.push(...paragraphBlock(el, rels))
    else if (el.name === 'w:tbl') blocks.push(tableBlock(el, rels))
    else if (el.name === 'w:sdt') for (const c of findAll(el, 'w:p')) blocks.push(...paragraphBlock(c, rels))
  }
  const doc_ = { kind: 'document', meta: readCoreProps(zipObj), blocks }
  normalizeListRuns(doc_)
  return doc_
}

function normalizeListRuns(doc) {
  // merge consecutive single-item lists
  const out = []
  for (const b of doc.blocks) {
    const prev = out[out.length - 1]
    if (b.type === 'list' && prev && prev.type === 'list' && prev.ordered === b.ordered && prev._single && b._single) {
      prev.items.push(...b.items)
    } else out.push(b)
  }
  for (const b of out) delete b._single
  doc.blocks = out
}

function parseRelationships(xml) {
  const map = new Map()
  if (!xml) return map
  for (const r of findAll(parseXML(xml), 'Relationship')) map.set(r.attrs.Id, { type: r.attrs.Type, target: r.attrs.Target })
  return map
}

function paragraphBlock(p, rels) {
  const out = []
  const pPr = findEl(p, 'w:pPr')
  const style = (pPr && findEl(pPr, 'w:pStyle')?.attrs['w:val']) || ''
  const numPr = pPr && findEl(pPr, 'w:numPr')
  const runs = []
  let hasPageBreak = false

  const runFrom = (r, rPr) => {
    let text = ''
    for (const c of r.children) {
      if (c.name === 'w:t') text += c.children.map(x => typeof x === 'string' ? x : '').join('')
      else if (c.name === 'w:tab') text += '\t'
      else if (c.name === 'w:br' || c.name === 'w:cr') text += '\n'
    }
    if (!text) return null
    const run = { text }
    if (rPr) {
      if (findEl(rPr, 'w:b') && findEl(rPr, 'w:b').attrs['w:val'] !== '0') run.bold = true
      if (findEl(rPr, 'w:i') && findEl(rPr, 'w:i').attrs['w:val'] !== '0') run.italic = true
      if (findEl(rPr, 'w:strike')) run.strike = true
      if (findEl(rPr, 'w:u') && findEl(rPr, 'w:u').attrs['w:val'] !== 'none') run.underline = true
    }
    return run
  }

  const collect = node => {
    if (typeof node === 'string' || node.raw || node.cdata) return
    if (node.name === 'w:r') {
      const run = runFrom(node, findEl(node, 'w:rPr'))
      if (run) runs.push(run)
      // 只含图形/域的 run（没有 w:t，`runFrom` 会返回 null）不能让内容凭空消失：
      // 第十二轮起 `office_edit` 能插图，读回时至少要看得见"这里有一张图"。
      for (const c of node.children) {
        if (typeof c === 'string') continue
        if (c.name === 'w:drawing' || c.name === 'w:pict' || c.name === 'w:object') {
          runs.push({ text: run ? ' [图片]' : '[图片]', italic: true })
          break
        }
        if (c.name === 'w:fldSimple') { runs.push({ text: textOf(c) || '[域]', italic: true }); break }
      }
      return
    }
    if (node.name === 'w:hyperlink') {
      const link = node.attrs['r:id'] && rels.get(node.attrs['r:id'])?.target
      for (const r of children(node, 'w:r')) {
        const run = runFrom(r, findEl(r, 'w:rPr'))
        if (run) { if (link) run.link = link; runs.push(run) }
      }
      return
    }
    if (node.name === 'w:br' && node.attrs['w:type'] === 'page') { hasPageBreak = true; return }
    if (node.name === 'w:drawing' || node.name === 'w:pict' || node.name === 'w:object') {
      runs.push({ text: '[图片]', italic: true })
      return
    }
    if (node.name === 'w:fldSimple') {
      runs.push({ text: textOf(node) || '[域]', italic: true })
      return
    }
    if (node.name === 'w:proofErr' || node.name === 'w:bookmarkStart' || node.name === 'w:bookmarkEnd' || node.name === 'w:commentRangeStart' || node.name === 'w:commentRangeEnd') return
    for (const c of node.children) collect(c)
  }
  for (const c of p.children) collect(c)
  if (hasPageBreak) out.push({ type: 'pagebreak' })
  const text = plainOf(runs)
  if (!text.trim()) { if (runs.length) out.push({ type: 'paragraph', runs }); return out }
  const hm = /(?:heading|标题)\s*([1-6])|^h([1-6])$/i.exec(style)
  const level = hm ? Number(hm[1] || hm[2]) : null
  if (level && level <= 6) out.push({ type: 'heading', level, text })
  else if (numPr) {
    const ilvl = Number(findEl(numPr, 'w:ilvl')?.attrs['w:val'] || 0)
    out.push({ type: 'list', ordered: false, items: [{ text, level: ilvl }], _single: true })
  } else if (/^Quote/i.test(style) || /^引用/.test(style)) out.push({ type: 'quote', text })
  else out.push({ type: 'paragraph', runs })
  return out
}

function tableBlock(tbl, rels) {
  const rows = []
  for (const tr of children(tbl, 'w:tr')) {
    const cells = []
    for (const tc of children(tr, 'w:tc')) {
      const parts = []
      for (const p of children(tc, 'w:p')) {
        const blocks = paragraphBlock(p, rels)
        parts.push(blocks.map(b =>
          b.type === 'heading' ? b.text
            : b.type === 'paragraph' ? plainOf(b.runs)
              : b.type === 'list' ? b.items.map(i => i.text).join('; ')
                : b.text || '').join(' '))
      }
      cells.push(parts.join('\n'))
    }
    rows.push(cells)
  }
  return { type: 'table', header: false, rows }
}

function readCoreProps(zipObj) {
  const meta = {}
  const core = zipObj.getText('docProps/core.xml')
  if (!core) return meta
  const d = parseXML(core)
  const get = name => { const n = findEl(d, name); return n ? textOf(n) : undefined }
  for (const [k, v] of Object.entries({
    title: get('dc:title'), author: get('dc:creator'), subject: get('dc:subject'),
    keywords: get('cp:keywords'), description: get('dc:description'),
    created: get('dcterms:created'), modified: get('dcterms:modified'),
  })) if (v) meta[k] = v
  return meta
}

// ---------- writing ----------

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="200" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/><w:spacing w:before="160" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:b/><w:color w:val="595959"/><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:fill="F5F5F5"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`

function numberingXml() {
  const bulletLvls = [0, 1, 2, 3, 4, 5, 6, 7, 8].map(i => {
    const glyph = i % 3 === 0 ? '●' : i % 3 === 1 ? '○' : '▪'
    return `<w:lvl w:ilvl="${i}"><w:numFmt w:val="bullet"/><w:lvlText w:val="${glyph}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" w:hint="default"/></w:rPr></w:lvl>`
  }).join('')
  const numLvls = [0, 1, 2, 3, 4, 5, 6, 7, 8].map(i =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${numLvls}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`
}

/**
 * 文档网格（申论稿纸，第十二轮需求 4c）：`meta.grid` → `<w:docGrid>`。
 *
 * 接受 `{chars, lines}`、`number`（= 每行字数，行数按同值）或 `'20x25'` 字符串。
 * 只写 `w:docGrid`（Word 的"文档网格"）：它把每行字数/每页行数**钉死**，
 * 稿纸要的等宽节奏就有了；但**可见的方格线框**属于页面背景，Word 里要靠
 * "页面边框/背景"，不在本参数范围内（README/SKILL 已注明）。
 * 版面按 A4（11906×16838 twips）+ 1440 twips 页边距算可用区。
 */
export function docGridXml(grid) {
  const g = normalizeGrid(grid)
  if (!g) return ''
  const usableW = 11906 - 1440 * 2
  const usableH = 16838 - 1440 * 2
  const charPitch = usableW / g.chars            // twips/字
  const linePitch = Math.round(usableH / g.lines) // twips/行
  // charSpace 的单位是 1/20 pt：charPitch(twips) → pt = /20 → 1/20pt = 原值
  const charSpace = Math.max(0, Math.round(charPitch) - 240)
  return `<w:docGrid w:type="linesAndChars" w:linePitch="${linePitch}" w:charSpace="${charSpace}"/>`
}

/** 归一化 grid 参数；无法识别时返回 null（调用方据此报错，不静默忽略）。 */
export function normalizeGrid(grid) {
  if (grid === undefined || grid === null || grid === '') return null
  const clamp = n => Math.max(2, Math.min(60, Math.round(n)))
  if (typeof grid === 'number') return { chars: clamp(grid), lines: clamp(grid) }
  if (typeof grid === 'string') {
    const m = /^(\d{1,2})\s*[x×*]\s*(\d{1,2})$/i.exec(grid.trim())
    if (m) return { chars: clamp(Number(m[1])), lines: clamp(Number(m[2])) }
    if (/^\d{1,2}$/.test(grid.trim())) return { chars: clamp(Number(grid)), lines: clamp(Number(grid)) }
    return null
  }
  if (typeof grid === 'object') {
    const chars = Number(grid.chars ?? grid.cols ?? grid.perLine)
    const lines = Number(grid.lines ?? grid.rows ?? grid.perPage ?? chars)
    if (!Number.isFinite(chars) || !Number.isFinite(lines)) return null
    return { chars: clamp(chars), lines: clamp(lines) }
  }
  return null
}

export function writeDocx(doc, opts = {}) {
  const d = { kind: 'document', meta: {}, blocks: [], ...doc }
  const ctx = newImageContext(1000)
  const paragraphs = []
  const hyperlinks = []
  let ridSeq = 2 // rId1=styles, rId2=numbering

  const nextRid = () => `rId${++ridSeq}`

  for (const b of d.blocks) {
    if (b.type === 'heading') {
      paragraphs.push(`<w:p><w:pPr><w:pStyle w:val="Heading${Math.min(6, b.level || 1)}"/></w:pPr>${runsXml([{ text: b.text }], hyperlinks, nextRid)}</w:p>`)
    } else if (b.type === 'paragraph') {
      paragraphs.push(`<w:p>${runsXml(b.runs, hyperlinks, nextRid)}</w:p>`)
    } else if (b.type === 'quote') {
      for (const l of String(b.text || '').split('\n')) {
        paragraphs.push(`<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${runsXml([{ text: l }], hyperlinks, nextRid)}</w:p>`)
      }
    } else if (b.type === 'code') {
      for (const l of String(b.text || '').split('\n')) {
        paragraphs.push(`<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${runsXml([{ text: l || ' ' }], hyperlinks, nextRid)}</w:p>`)
      }
    } else if (b.type === 'list') {
      for (const it of b.items || []) {
        const numId = b.ordered ? 2 : 1
        const ilvl = Math.max(0, Math.min(8, it.level || 0))
        paragraphs.push(`<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${runsXml([{ text: it.text }], hyperlinks, nextRid)}</w:p>`)
      }
    } else if (b.type === 'table') {
      paragraphs.push(tableXml(b.rows || [], hyperlinks, nextRid))
      paragraphs.push('<w:p/>')
    } else if (b.type === 'hr') {
      paragraphs.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>')
    } else if (b.type === 'pagebreak') {
      paragraphs.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    } else if (b.type === 'image') {
      paragraphs.push(imageBlockXml(b, hyperlinks, nextRid, ctx))
    } else if (b.text) {
      paragraphs.push(`<w:p>${runsXml([{ text: String(b.text) }], hyperlinks, nextRid)}</w:p>`)
    }
  }

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${docGridXml(d.meta?.grid)}</w:sectPr></w:body></w:document>`

  const entries = [
    { name: '[Content_Types].xml', data: ctx.media.length ? contentTypesWithImages(CONTENT_TYPES_XML, ctx.media) : CONTENT_TYPES_XML },
    { name: '_rels/.rels', data: ROOT_RELS_XML },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: docRelsXml(hyperlinks, ctx.rels) },
    { name: 'word/styles.xml', data: STYLES_XML },
    { name: 'word/numbering.xml', data: numberingXml() },
    { name: 'docProps/core.xml', data: coreXml(d.meta) },
    { name: 'docProps/app.xml', data: APP_XML },
  ]
  for (const m of ctx.media) entries.push({ name: m.name, data: m.data })
  flushImageInfo(ctx, opts.info)
  return makeZip(entries)
}

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`

const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>DSH Office</Application></Properties>`

function docRelsXml(hyperlinks, images = []) {
  let body = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
  body += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>\n'
  for (const h of hyperlinks) {
    body += `<Relationship Id="${h.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${encodeEntities(h.target, true)}" TargetMode="External"/>\n`
  }
  for (const im of images) {
    body += `<Relationship Id="${im.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${encodeEntities(im.target, true)}"/>\n`
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${body}</Relationships>`
}

function coreXml(meta) {
  const now = new Date().toISOString().replace(/\.\d+Z/, 'Z')
  const s = v => encodeEntities(v ?? '', true)
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${meta.title ? `<dc:title>${s(meta.title)}</dc:title>` : ''}
${meta.author ? `<dc:creator>${s(meta.author)}</dc:creator>` : ''}
${meta.subject ? `<dc:subject>${s(meta.subject)}</dc:subject>` : ''}
${meta.keywords ? `<cp:keywords>${s(meta.keywords)}</cp:keywords>` : ''}
${meta.description ? `<dc:description>${s(meta.description)}</dc:description>` : ''}
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`
}

function singleRunXml(run, hyperlinks, nextRid) {
  const props = []
  if (run.bold) props.push('<w:b/>')
  if (run.italic) props.push('<w:i/>')
  if (run.underline) props.push('<w:u w:val="single"/>')
  if (run.strike) props.push('<w:strike/>')
  if (run.code) props.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')
  if (run.color && /^[0-9a-fA-F]{6}$/.test(run.color)) props.push(`<w:color w:val="${run.color}"/>`)
  if (run.size) props.push(`<w:sz w:val="${2 * run.size}"/>`)
  const rPr = props.length ? `<w:rPr>${props.join('')}</w:rPr>` : ''
  const pieces = String(run.text ?? '').split('\n')
  let out = ''
  pieces.forEach((piece, idx) => {
    out += `<w:r>${rPr}${idx ? '<w:br/>' : ''}${piece ? `<w:t xml:space="preserve">${encodeEntities(piece)}</w:t>` : '<w:t xml:space="preserve"> </w:t>'}</w:r>`
  })
  return out
}

function runsXml(runs, hyperlinks, nextRid) {
  let out = ''
  for (const r of runs || []) {
    const run = typeof r === 'string' ? { text: r } : r
    if (run.link) {
      const id = nextRid()
      hyperlinks.push({ id, target: run.link })
      const body = singleRunXml({ ...run, link: undefined, color: '0563C1' }, hyperlinks, nextRid)
      out += `<w:hyperlink r:id="${id}">${body}</w:hyperlink>`
      continue
    }
    out += singleRunXml(run, hyperlinks, nextRid)
  }
  return out
}

function tableXml(rows, hyperlinks, nextRid) {
  const grid = Math.max(1, (rows[0] || []).length)
  let xml = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${Array.from({ length: grid }, () => `<w:gridCol w:w="${Math.floor(9000 / grid)}"/>`).join('')}</w:tblGrid>`
  for (const r of rows) {
    xml += '<w:tr>'
    for (let c = 0; c < grid; c++) {
      const cellText = String(r[c] ?? '')
      const paras = cellText.split(/\r?\n/).map(t => `<w:p>${runsXml([{ text: t }], hyperlinks, nextRid)}</w:p>`).join('')
      xml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${paras || '<w:p/>'}</w:tc>`
    }
    xml += '</w:tr>'
  }
  return xml + '</w:tbl>'
}

// ---------- editing ----------

/** Replace text across w:t nodes (document part level); returns new XML or null. */
export function replaceTextInDocument(documentXml, find, replace, { regex = false } = {}) {
  const doc = parseXML(documentXml)
  const ts = findAll(doc, 'w:t')
  let count = 0
  let re = null
  if (regex) {
    try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`invalid regex: ${e.message}`) }
  }
  for (const t of ts) {
    const raw = t.children.map(c => typeof c === 'string' ? c : '').join('')
    const decoded = raw
    if (re) {
      const hits = decoded.match(re)
      if (hits) { count += hits.length; t.children = [encodeEntities(decoded.replace(re, () => replace))] }
    } else if (decoded.includes(find)) {
      count += decoded.split(find).length - 1
      t.children = [encodeEntities(decoded.split(find).join(replace))]
    }
  }
  return { xml: serializeXML(doc), count }
}

/** Append model blocks to the end of body (before sectPr). */
export function appendBlocksToDocument(documentXml, blocks, opts = {}) {
  const doc = parseXML(documentXml)
  const body = findEl(doc, 'w:body') || doc
  void body
  const hyperlinks = []
  const ctx = newImageContext(2000)
  // 追加到**既有**文档时：媒体序号必须接着既有 `word/media/imageN` 编（绝不覆盖既有图片），
  // 内容去重表由调用方带进来（同一个 Map 结构 {part, rid}）。
  if (Number.isFinite(opts.mediaSeqStart) && opts.mediaSeqStart > 0) ctx.mediaSeq = opts.mediaSeqStart
  if (opts.byHash instanceof Map) ctx.byHash = opts.byHash
  let seq = 0
  const nextRid = () => `dshRel${++seq}`
  const clone = JSON.parse(JSON.stringify({ blocks }))
  let xml = ''
  for (const b of clone.blocks) xml += blockToXml(b, hyperlinks, nextRid, ctx)
  // splice raw text before closing sectPr/body: rebuild via string insertion
  const insertAt = documentXml.lastIndexOf('<w:sectPr') >= 0
    ? documentXml.lastIndexOf('<w:sectPr')
    : documentXml.lastIndexOf('</w:body>')
  if (insertAt < 0) throw new Error('docx: cannot locate body end')
  return {
    xml: documentXml.slice(0, insertAt) + xml + documentXml.slice(insertAt),
    hyperlinks,
    media: ctx.media,
    imageRels: ctx.rels,
    imagesSkipped: ctx.skipped,
    imageReused: ctx.reused,
    imageSizing: ctx.sizing,
  }
}

function blockToXml(b, hyperlinks, nextRid, imgCtx) {
  if (b.type === 'heading') return `<w:p><w:pPr><w:pStyle w:val="Heading${Math.min(6, b.level || 1)}"/></w:pPr>${runsXml([{ text: b.text }], hyperlinks, nextRid)}</w:p>`
  if (b.type === 'image') {
    // 没有图片上下文（老调用方）时退化为字面文本，**绝不静默丢图**
    return imgCtx
      ? imageBlockXml(b, hyperlinks, nextRid, imgCtx)
      : `<w:p>${runsXml([{ text: `![${b.alt || '图片'}](${b.name || ''})` }], hyperlinks, nextRid)}</w:p>`
  }
  if (b.type === 'paragraph') return `<w:p>${runsXml(b.runs || [{ text: b.text || '' }], hyperlinks, nextRid)}</w:p>`
  if (b.type === 'list') {
    return (b.items || []).map(it => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${Math.max(0, Math.min(8, it.level || 0))}"/><w:numId w:val="${b.ordered ? 2 : 1}"/></w:numPr></w:pPr>${runsXml([{ text: it.text }], hyperlinks, nextRid)}</w:p>`).join('')
  }
  if (b.type === 'table') return tableXml(b.rows || [], hyperlinks, nextRid) + '<w:p/>'
  if (b.type === 'pagebreak') return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
  if (b.type === 'hr') return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
  if (b.type === 'quote') return String(b.text || '').split('\n').map(l => `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${runsXml([{ text: l }], hyperlinks, nextRid)}</w:p>`).join('')
  if (b.type === 'code') return String(b.text || '').split('\n').map(l => `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${runsXml([{ text: l || ' ' }], hyperlinks, nextRid)}</w:p>`).join('')
  return ''
}

// ---------- 插图（第十二轮 需求 1b） ----------
//
// `office_edit` 的 docx 插图操作：zip 级新增 `word/media/*` 部件 + image 关系 +
// 一段 `<w:drawing>`。**命名空间就地声明**在这段 XML 自己身上（`wp`/`a`/`pic`），
// 不动既有文档的根元素 —— 目标文件可能是 Word 产出的、根上带一堆 `mc:Ignorable`，
// 改根反而容易踩雷。
// ⚠ URI 必须是 ECMA-376 的标准值 `…/drawingml/2006/wordprocessingDrawing`。
// 之前写的 `…/drawingWordprocessingDrawing/2006/main` 是**不存在的命名空间**：
// 包结构体检、rId 唯一性、Word 打开**全都会**通过字面检查，但 Word 解析 `wp:inline`
// 时找不到元素定义 → 直接判包损坏（0x80070570 / "文件已损坏"）。
// 第二轮 R13 实测：含图 docx 100% 被 Word 16.0 拒开，改这一行后即恢复正常。
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

/**
 * 内联图片段落 XML。
 * @param {string} rid image 关系 id（`r:embed`）
 * @param {number} cx 宽（EMU） @param {number} cy 高（EMU）
 * @param {number} id 文档内唯一图形 id @param {string} descr 替代文字
 */
export function imageParagraphXml(rid, cx, cy, id, descr = '') {
  const d = encodeEntities(descr, true)
  return '<w:p><w:r><w:drawing>'
    + `<wp:inline xmlns:wp="${WP_NS}" distT="0" distB="0" distL="0" distR="0">`
    + `<wp:extent cx="${cx}" cy="${cy}"/>`
    + '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    + `<wp:docPr id="${id}" name="图片 ${id}" descr="${d}"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic xmlns:a="${A_NS}">`
    + `<a:graphicData uri="${PIC_NS}">`
    + `<pic:pic xmlns:pic="${PIC_NS}">`
    + `<pic:nvPicPr><pic:cNvPr id="${id}" name="图片 ${id}" descr="${d}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    + '</pic:pic>'
    + '</a:graphicData>'
    + '</a:graphic>'
    + '</wp:inline>'
    + '</w:drawing></w:r></w:p>'
}

/**
 * 把图片段落插进 `word/document.xml`。
 * @param {string} documentXml
 * @param {string} paraXml `imageParagraphXml()` 的产物
 * @param {{after?:string, at?:'start'|'end'}} place
 * @returns {{xml:string, inserted:boolean}}
 */
export function insertImageParagraph(documentXml, paraXml, place = {}) {
  const at = place.at || (place.after ? 'after' : 'end')
  const bodyEnd = documentXml.lastIndexOf('<w:sectPr') >= 0
    ? documentXml.lastIndexOf('<w:sectPr')
    : documentXml.lastIndexOf('</w:body>')
  if (bodyEnd < 0) throw new Error('docx: 找不到 w:body 结束位置，无法插入图片')
  if (at === 'end' || !place.after) {
    if (at === 'end') return { xml: documentXml.slice(0, bodyEnd) + paraXml + documentXml.slice(bodyEnd), inserted: true }
    // at=start：插到 <w:body> 之后
    const m = /<w:body[^>]*>/.exec(documentXml)
    if (!m) throw new Error('docx: 找不到 w:body，无法插入图片')
    const pos = m.index + m[0].length
    return { xml: documentXml.slice(0, pos) + paraXml + documentXml.slice(pos), inserted: true }
  }
  // after=<文本>：在该文本所在**段落**之后插入（取第一个命中；命中不到退化为文末）
  const needle = encodeEntities(String(place.after))
  const hit = documentXml.indexOf(needle)
  if (hit < 0) return { xml: documentXml.slice(0, bodyEnd) + paraXml + documentXml.slice(bodyEnd), inserted: false }
  const close = documentXml.indexOf('</w:p>', hit)
  if (close < 0) return { xml: documentXml.slice(0, bodyEnd) + paraXml + documentXml.slice(bodyEnd), inserted: false }
  const pos = close + '</w:p>'.length
  return { xml: documentXml.slice(0, pos) + paraXml + documentXml.slice(pos), inserted: true }
}

/** `[Content_Types].xml` 里补一个图片扩展名的 Default（已有则原样返回）。 */
export function ensureContentTypeDefault(ctXml, ext, mime) {
  if (new RegExp(`<Default\\s+Extension="${ext}"\\b`, 'i').test(ctXml)) return ctXml
  const add = `<Default Extension="${ext}" ContentType="${mime}"/>`
  return ctXml.includes('</Types>') ? ctXml.replace('</Types>', `${add}</Types>`) : ctXml + add
}

export function docBlocksToXml(blocks) {
  const hyperlinks = []
  const ctx = newImageContext(3000)
  let ridSeq = 200
  const nextRid = () => `rId${++ridSeq}`
  const xml = (blocks || []).map(b => blockToXml(b, hyperlinks, nextRid, ctx)).join('')
  return {
    xml,
    hyperlinks,
    media: ctx.media,
    imageRels: ctx.rels,
    imagesSkipped: ctx.skipped,
    imageReused: ctx.reused,
    imageSizing: ctx.sizing,
  }
}

// ---------- 插图写出端（第二轮需求 1a / 4c / 4e） ----------
//
// 输入侧（`model.js::markdownToDocument`）会把独立成行的 `![alt](path)` 变成 image 块；
// 写出端此前只有 PDF / HTML 会内嵌，docx 落进 `blockToXml` 的 default 分支 → 静默丢图。
// 这里补上 zip 级内嵌：`word/media/imageN.<ext>` + image 关系 + 一段 `<w:drawing>`。
//
// 三条硬规则（都对应"不静默"红线）：
//   ① 媒体部件按**内容 SHA-256 去重**（同图复用同一部件与同一 rId，逐次进 `imageReused`）；
//   ② 尺寸按显式规则算（省略 width → 原图像素 × 72/96；超出 A4 可用宽等比缩到可用宽），
//      换算过程进 `imageSizing`，由调用方转成 notice；
//   ③ 读不到 / 不认识的图片**绝不静默丢**：写回字面 `![alt](path)` 文本并进 `imagesSkipped`。
export const DOCX_USABLE_WIDTH_PT = 451.3   // A4 宽 11906twips − 左右边距 1440×2 = 9026twips = 451.3pt

/** 图片尺寸策略：原图像素 → pt；显式 width 优先；超可用宽等比缩（换算过程可回放）。 */
export function imageSizeFor(widthPx, heightPx, widthPt, opts = {}) {
  const usable = Number(opts.usablePt) > 0 ? Number(opts.usablePt) : DOCX_USABLE_WIDTH_PT
  const px = Number(widthPx) > 0 ? Number(widthPx) : 0
  const py = Number(heightPx) > 0 ? Number(heightPx) : 0
  const asked = Number(widthPt)
  const explicit = Number.isFinite(asked) && asked > 0
  const naturalPt = px ? (px * 72) / 96 : 0
  let width = explicit ? asked : naturalPt
  let capped = false
  if (width > usable) { width = usable; capped = true }
  const height = px ? (width * py) / px : 0
  const r2 = n => Math.round(n * 100) / 100
  return {
    widthPt: r2(width),
    heightPt: r2(height),
    naturalWidthPt: r2(naturalPt),
    rule: explicit ? 'width 参数' : '原图像素 × 72/96',
    capped,
  }
}

/** 一次写出/追加过程里的图片记账上下文（media 部件、关系、跳过、复用、尺寸）。 */
function newImageContext(graphId = 1000) {
  return {
    media: [],          // [{name, data, kind}]
    rels: [],           // [{id, target, kind}]
    skipped: [],        // [{name, alt, reason}]
    reused: [],         // [{name, part, reason}]
    sizing: [],         // [{name, part, px, pt, rule, capped}]
    byHash: new Map(),  // 内容哈希 → {part, rid}（同图复用同一部件与同一关系 id）
    mediaSeq: 0,
    graphId,
  }
}

/** 一个 image 块 → `<w:p><w:drawing>`；失败时退化为字面文本并记账（不静默）。 */
function imageBlockXml(b, hyperlinks, nextRid, ctx) {
  const alt = b.alt || '图片'
  const src = String(b.name ?? b.path ?? b.src ?? '')
  const literal = () => `<w:p>${runsXml([{ text: `![${alt}](${src})` }], hyperlinks, nextRid)}</w:p>`
  let img
  try { img = readImageBytes(src) } catch (e) {
    ctx.skipped.push({ name: src, alt, reason: `读不到图片：${e?.message || e}` })
    return literal()
  }
  const meta = sniffImage(img.buf)
  if (!meta || !meta.width || !meta.height) {
    ctx.skipped.push({ name: src, alt, reason: '不认识的图片格式（只支持 PNG/JPEG/GIF/BMP）' })
    return literal()
  }
  const hash = createHash('sha256').update(img.buf).digest('hex')
  const dup = ctx.byHash.get(hash)
  let part, rid
  if (dup) {
    part = dup.part
    rid = dup.rid
    ctx.reused.push({ name: src, part, reason: '内容 SHA-256 与已有媒体部件相同，复用同一部件与关系 id' })
  } else {
    const ext = imageExt(meta)
    ctx.mediaSeq += 1
    part = `word/media/image${ctx.mediaSeq}.${ext}`
    rid = nextRid()
    ctx.media.push({ name: part, data: img.buf, kind: ext })
    ctx.rels.push({ id: rid, target: `media/image${ctx.mediaSeq}.${ext}`, kind: ext })
    ctx.byHash.set(hash, { part, rid })
  }
  const size = imageSizeFor(meta.width, meta.height, b.width ?? b.widthPt)
  ctx.sizing.push({
    name: src,
    part,
    px: `${meta.width}×${meta.height}`,
    pt: `${size.widthPt}×${size.heightPt}`,
    rule: size.rule,
    capped: size.capped,
  })
  ctx.graphId += 1
  const cx = Math.round(size.widthPt * 12700)
  const cy = Math.round(size.heightPt * 12700)
  return imageParagraphXml(rid, cx, cy, ctx.graphId, alt)
}

/** 有图片时把 `[Content_Types].xml` 的图片 Default 补齐（无图时调用方直接用原串，逐字节不变）。 */
function contentTypesWithImages(ctXml, media) {
  let out = ctXml
  for (const kind of new Set(media.map(m => m.kind))) out = ensureContentTypeDefault(out, kind, imageMime(kind))
  return out
}

/** 把图片记账并进调用方的 info 对象（只在真有内容时写字段，干净文件不加键）。 */
function flushImageInfo(ctx, info) {
  if (!info || typeof info !== 'object') return
  if (ctx.media.length) info.imageMedia = ctx.media.map(m => m.name)
  if (ctx.skipped.length) info.imagesSkipped = ctx.skipped
  if (ctx.reused.length) info.imageReused = ctx.reused
  if (ctx.sizing.length) info.imageSizing = ctx.sizing
}
