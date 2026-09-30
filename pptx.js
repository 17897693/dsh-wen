// PowerPoint (.pptx) reader/writer on the OOXML zip container.
// The writer emits a complete minimal package (presentation, master, two
// layouts, theme, notes master) so the file opens in PowerPoint, WPS and
// LibreOffice.
import { openZip, makeZip } from './zip.js'
import { parseXML, serializeXML, findAll, children, textOf, encodeEntities } from './xml.js'
import { serializeSlideXmlForEdit, slideRelsForEdit, notesXmlForEdit, notesRelsForEdit } from './pptx-edit.js'

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function findOne(node, name) { return findAll(node, name)[0] }

/** R19 任务 D：同 `docx.js::decodeEntitiesRaw` —— 数字实体要守卫（`&#` + 超长数字 ⇒ `Infinity`
 *  ⇒ `String.fromCodePoint` 抛 `RangeError`）。本函数当前**零调用点**，守卫是为将来接上时不再踩。 */
const codePointOrRaw = (all, code) => (Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all)
function decodeRaw(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (all, d) => codePointOrRaw(all, Number(d))).replace(/&#x([0-9a-f]+);/gi, (all, h) => codePointOrRaw(all, parseInt(h, 16)))
}

// ---------- reading ----------

export function readPptx(zipInput) {
  const zipObj = typeof zipInput === 'object' && zipInput.names ? zipInput : openZip(zipInput)
  const presXml = zipObj.getText('ppt/presentation.xml')
  if (presXml === undefined) throw new Error('pptx: ppt/presentation.xml missing')
  const pres = parseXML(presXml)
  const presRels = parseRelationships(zipObj.getText('ppt/_rels/presentation.xml.rels'))
  const sldIdLst = findOne(pres, 'p:sldIdLst')
  const slidePaths = []
  if (sldIdLst) for (const s of children(sldIdLst, 'p:sldId')) {
    const target = s.attrs['r:id'] && presRels.get(s.attrs['r:id'])?.target
    if (target) slidePaths.push('ppt/' + String(target).replace(/^\.\.\//, '').replace(/^\//, ''))
  }
  const slides = []
  for (const path of slidePaths) {
    const xml = zipObj.getText(path)
    if (xml === undefined) { slides.push({ layout: 'blank', title: '', bullets: [] }); continue }
    slides.push(readSlide(xml, path, zipObj))
  }
  return { kind: 'slides', meta: readCoreProps(zipObj), slides }
}

function readSlide(xml, path, zipObj) {
  const doc = parseXML(xml)
  const slide = { layout: 'content', title: '', bullets: [] }
  const cSld = findOne(doc, 'p:cSld')
  const spTree = cSld && findOne(cSld, 'p:spTree')
  const walk = node => {
    if (typeof node === 'string' || node.raw || node.cdata) return
    if (node.name === 'p:sp') {
      const phType = (findOne(node, 'p:ph') || {}).attrs?.type || ''
      const txBody = findOne(node, 'p:txBody')
      if (txBody) {
        const paras = txBodyParagraphs(txBody)
        const joined = paras.map(p => p.text).join('\n').trim()
        if (/^(ctrTitle|title)$/.test(phType)) slide.title = joined || slide.title
        else if (phType === 'subTitle') slide.subtitle = joined
        else for (const p of paras) if (p.text.trim()) slide.bullets.push({ text: p.text, level: p.level })
      }
      return
    }
    if (node.name === 'p:graphicFrame') {
      const tbl = findOne(node, 'a:tbl')
      if (tbl) {
        const rows = []
        for (const tr of children(tbl, 'a:tr')) {
          const cells = []
          for (const tc of children(tr, 'a:tc')) cells.push(txBodyParagraphs(findOne(tc, ['a:txBody', 'p:txBody'])).map(p => p.text).join('\n'))
          rows.push(cells)
        }
        slide.table = { rows }
      }
      return
    }
    if (node.name === 'p:grpSp') { for (const c of node.children) walk(c); return }
    if (node.name === 'p:pic') { slide.images = (slide.images || 0) + 1; return }
    for (const c of node.children) walk(c)
  }
  if (spTree) for (const c of spTree.children) walk(c)
  // speaker notes
  const relPath = path.replace(/slides\/([^/]+)\.xml$/, 'slides/_rels/$1.xml.rels')
  const rels = parseRelationships(zipObj.getText(relPath))
  for (const rel of rels.values()) {
    if (/notesSlide\d+\.xml$/.test(String(rel.target))) {
      const nxml = zipObj.getText('ppt/' + String(rel.target).replace(/^\.\.\//, ''))
      if (nxml) {
        const ndoc = parseXML(nxml)
        const bodies = findAll(ndoc, 'p:sp').map(sp => findOne(sp, 'p:txBody')).filter(Boolean)
          .map(tb => txBodyParagraphs(tb).map(p => p.text).join('\n').trim()).filter(Boolean)
        if (bodies.length) slide.notes = bodies[bodies.length - 1]
      }
    }
  }
  if (!slide.bullets.length && !slide.title && !slide.table) slide.layout = 'blank'
  return slide
}

function txBodyParagraphs(txBody) {
  const paras = []
  for (const p of children(txBody, 'a:p')) {
    let text = ''
    const level = Number(findOne(p, 'a:pPr')?.attrs.lvl || 0)
    for (const node of p.children) {
      if (node.name === 'a:r') text += findAll(node, 'a:t').map(t => textOf(t)).join('')
      else if (node.name === 'a:br') text += '\n'
      else if (node.name === 'a:fld') text += textOf(node)
    }
    paras.push({ text, level })
  }
  return paras
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
  for (const [k, v] of Object.entries({ title: get('dc:title'), author: get('dc:creator'), modified: get('dcterms:modified') })) if (v) meta[k] = v
  return meta
}

// ---------- writing ----------

const SLIDE_W = 12192000
const SLIDE_H = 6858000

export function writePptx(slidesInput) {
  const s = { kind: 'slides', meta: {}, slides: [], ...slidesInput }
  const n = Math.max(1, s.slides.length)
  const entries = [
    { name: '[Content_Types].xml', data: contentTypes(n, s.slides) },
    { name: '_rels/.rels', data: ROOT_RELS },
    { name: 'ppt/presentation.xml', data: presentationXml(n) },
    { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels(n) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: SLIDE_MASTER },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: MASTER_RELS },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: LAYOUT_TITLE },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: LAYOUT_RELS },
    { name: 'ppt/slideLayouts/slideLayout2.xml', data: LAYOUT_CONTENT },
    { name: 'ppt/slideLayouts/_rels/slideLayout2.xml.rels', data: LAYOUT_RELS },
    { name: 'ppt/notesMasters/notesMaster1.xml', data: NOTES_MASTER },
    { name: 'ppt/notesMasters/_rels/notesMaster1.xml.rels', data: NOTES_MASTER_RELS },
    { name: 'ppt/theme/theme1.xml', data: THEME },
    // 第二轮 R13 / 需求 7：notesMaster **必须**有自己独立的 theme 部件。
    // PowerPoint 16.0 会因为 notesMaster 与 slideMaster 共用同一个 theme 部件而判包损坏（0x80070570）。
    { name: 'ppt/theme/theme2.xml', data: THEME_NOTES },
    { name: 'ppt/presProps.xml', data: PRES_PROPS },
    { name: 'ppt/tableStyles.xml', data: TABLE_STYLES },
    { name: 'ppt/viewProps.xml', data: VIEW_PROPS },
    { name: 'docProps/core.xml', data: coreXml(s.meta) },
    { name: 'docProps/app.xml', data: appXml(n) },
  ]
  for (let i = 0; i < n; i++) {
    const sl = s.slides[i]
    entries.push({ name: `ppt/slides/slide${i + 1}.xml`, data: serializeSlideXmlForEdit(sl) })
    entries.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: slideRelsForEdit(i + 1, sl) })
    if (sl.notes) {
      entries.push({ name: `ppt/notesSlides/notesSlide${i + 1}.xml`, data: notesXmlForEdit(sl, i + 1) })
      entries.push({ name: `ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`, data: notesRelsForEdit(i + 1) })
    }
  }
  return makeZip(entries)
}

function contentTypes(n, slides = []) {
  let overrides = `
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`
  for (let i = 1; i <= n; i++) overrides += `\n<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  // speaker-note parts need their own override, or PowerPoint reads them as plain xml
  for (let i = 0; i < slides.length; i++) {
    if (!slides[i]?.notes) continue
    overrides += `\n<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>${overrides}
</Types>`
}

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="${R_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

// rel id layout: rId1 master, rId2 theme, rId3 presProps, rId4..(3+n) slides,
// then notesMaster rId(n+4), tableStyles rId(n+5), viewProps rId(n+6).
function presentationXml(n) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:notesMasterIdLst><p:notesMasterId r:id="rId${n + 4}"/></p:notesMasterIdLst>
<p:sldIdLst>${Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 4}"/>`).join('')}</p:sldIdLst>
<p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/>
<p:notesSz cx="${SLIDE_H}" cy="${SLIDE_W}"/>
<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle>
</p:presentation>`
}

function presentationRels(n) {
  let body = `<Relationship Id="rId1" Type="${R_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="${R_NS}/theme" Target="theme/theme1.xml"/>
<Relationship Id="rId3" Type="${R_NS}/presProps" Target="presProps.xml"/>`
  for (let i = 0; i < n; i++) body += `\n<Relationship Id="rId${i + 4}" Type="${R_NS}/slide" Target="slides/slide${i + 1}.xml"/>`
  body += `
<Relationship Id="rId${n + 4}" Type="${R_NS}/notesMaster" Target="notesMasters/notesMaster1.xml"/>
<Relationship Id="rId${n + 5}" Type="${R_NS}/tableStyles" Target="tableStyles.xml"/>
<Relationship Id="rId${n + 6}" Type="${R_NS}/viewProps" Target="viewProps.xml"/>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
}




function txBody(paras) {
  return `<p:txBody><a:bodyPr rtlCol="0" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody>`
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

function appXml(n) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>DSH Office</Application><Slides>${n}</Slides></Properties>`
}

// ---------- static package parts ----------

const SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>${txBody('<a:p><a:endParaRPr lang="zh-CN"/></a:p>')}</p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="文本占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>${txBody('<a:p><a:endParaRPr lang="zh-CN"/></a:p>')}</p:sp>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400"><a:defRPr sz="4400" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="•"/><a:defRPr sz="2000"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="742950" indent="-285750"><a:buChar char="–"/><a:defRPr sz="1800"/></a:lvl2pPr><a:lvl3pPr marL="1143000" indent="-228600"><a:buChar char="▪"/><a:defRPr sz="1600"/></a:lvl3pPr></p:bodyStyle>
<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>
</p:sldMaster>`

const MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
<Relationship Id="rId3" Type="${R_NS}/theme" Target="../theme/theme1.xml"/>
</Relationships>`

function layoutXml(type, titlePh, bodyPh, bodyGeom) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="${type}" preserve="1"><p:cSld name="DSH-${type}"><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${titlePh}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>${txBody('<a:p><a:endParaRPr lang="zh-CN"/></a:p>')}</p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="内容占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${bodyPh}" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${bodyGeom[0]}" y="${bodyGeom[1]}"/><a:ext cx="${bodyGeom[2]}" cy="${bodyGeom[3]}"/></a:xfrm></p:spPr>${txBody('<a:p><a:endParaRPr lang="zh-CN"/></a:p>')}</p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
}

const LAYOUT_TITLE = layoutXml('title', 'ctrTitle', 'subTitle', [838200, 1825625, 10515600, 1238250])
const LAYOUT_CONTENT = layoutXml('cust', 'title', 'body', [838200, 1825625, 10515600, 4351338])

const LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`

const NOTES_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="图像占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="文本占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1435100"/><a:ext cx="5486400" cy="6096000"/></a:xfrm></p:spPr>${txBody('<a:p><a:endParaRPr lang="zh-CN"/></a:p>')}</p:sp>
</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:notesMaster>`

const NOTES_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/theme" Target="../theme/theme2.xml"/>
</Relationships>`

const PRES_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"/>`

const TABLE_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="${A_NS}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`

const VIEW_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"/>`

const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A_NS}" name="DSH Office"><a:themeElements>
<a:clrScheme name="DSH"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="DSH"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>`

/** 备注母版的独立主题（内容与 THEME 同构，仅 name 区分）—— 见 writePptx 的 theme2 部件。 */
const THEME_NOTES = THEME.replace('name="DSH Office"', 'name="DSH Office Notes"')

// ---------- editing ----------

export function replaceTextInPptxPart(partXml, find, replace, { regex = false } = {}) {
  const doc = parseXML(partXml)
  const ts = findAll(doc, 'a:t')
  let count = 0
  let re = null
  if (regex) {
    try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`invalid regex: ${e.message}`) }
  }
  for (const t of ts) {
    const raw = t.children.map(c => typeof c === 'string' ? c : '').join('')
    const decoded = raw
    if (re) { const hits = decoded.match(re); if (hits) { count += hits.length; t.children = [encodeEntities(decoded.replace(re, () => replace))] } }
    else if (decoded.includes(find)) { count += decoded.split(find).length - 1; t.children = [encodeEntities(decoded.split(find).join(replace))] }
  }
  return { xml: count ? serializeXML(doc) : null, count }
}
