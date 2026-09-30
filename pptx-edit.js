// Slide-level XML generators shared by the pptx writer and the editor.
import { encodeEntities } from './xml.js'

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export function aRuns(text, { bold = false, size = null, color = null } = {}) {
  const sz = size ? ` sz="${Math.round(size * 100)}"` : ''
  const b = bold ? ' b="1"' : ''
  // 颜色必须写在 a:rPr 内部：DrawingML 里 a:r / a:br 的子节点只有 (rPr?, t)，
  // 把 <a:solidFill> 放成 rPr 的兄弟会让 PowerPoint 判整个包损坏（0x80070570）。
  const rPr = `<a:rPr lang="zh-CN" altLang="en-US"${sz}${b}>${color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : ''}</a:rPr>`
  return String(text).split('\n').map((line, idx) =>
    `${idx ? `<a:br>${rPr}</a:br>` : ''}<a:r>${rPr}<a:t>${encodeEntities(line)}</a:t></a:r>`).join('')
}

export function txBody(paras) {
  return `<p:txBody><a:bodyPr rtlCol="0" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody>`
}

export function graphicTable(table, shapeId) {
  const rows = (table.rows || []).map(r => Array.isArray(r) ? r : [String(r)])
  const cols = Math.max(1, ...rows.map(r => r.length))
  const colW = Math.floor(10058400 / cols)
  const tbl = `<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('')}</a:tblGrid>${rows.map((r, ri) => `<a:tr h="${ri === 0 ? 370840 : 320040}">${Array.from({ length: cols }, (_, ci) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${aRuns(String(r[ci] ?? ''), { bold: ri === 0, size: 12 })}</a:p></a:txBody><a:tcPr/></a:tc>`).join('')}</a:tr>`).join('')}</a:tbl>`
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="表格"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="1069165" y="1800000"/><a:ext cx="${colW * cols}" cy="${Math.min(3600000, 370840 + rows.length * 320040)}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tbl}</a:graphicData></a:graphic>
</p:graphicFrame>`
}

export function serializeSlideXmlForEdit(sl) {
  const shapes = []
  let idSeq = 1
  const nid = () => ++idSeq
  const titleIsCtr = sl.layout === 'title'
  if (sl.title) shapes.push(
    `<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="标题 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${titleIsCtr ? 'ctrTitle' : 'title'}"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody(`<a:p>${aRuns(sl.title, { bold: true, size: titleIsCtr ? 40 : 28, color: '1F3864' })}</a:p>`)}</p:sp>`)
  if (sl.subtitle) shapes.push(
    `<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="副标题 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody(`<a:p>${aRuns(sl.subtitle, { size: 18, color: '404040' })}</a:p>`)}</p:sp>`)
  const bullets = (sl.bullets || []).filter(b => String(b.text ?? '').trim())
  if (bullets.length) {
    const paras = bullets.map(b => {
      const lvl = Math.max(0, Math.min(4, b.level || 0))
      return `<a:p><a:pPr lvl="${lvl}"/>${aRuns(String(b.text), { size: Math.max(12, 18 - lvl * 2) })}</a:p>`
    }).join('')
    shapes.push(
      `<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="内容占位符 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody(paras)}</p:sp>`)
  }
  if (sl.table?.rows?.length) shapes.push(graphicTable(sl.table, nid()))
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${shapes.join('\n')}
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
}

export function slideRelsForEdit(slideNo, sl, relPrefix = 'rId') {
  let body = `<Relationship Id="${relPrefix}1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout${sl.layout === 'title' ? 1 : 2}.xml"/>`
  if (sl.notes) body += `\n<Relationship Id="${relPrefix}2" Type="${R_NS}/notesSlide" Target="../notesSlides/notesSlide${slideNo}.xml"/>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
}

export function notesXmlForEdit(sl, slideNo) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="幻灯片图像占位符"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="备注占位符"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody(String(sl.notes).split('\n').map(l => `<a:p>${aRuns(l, { size: 12 })}</a:p>`).join(''))}</p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
}

export function notesRelsForEdit(slideNo, relPrefix = 'rId') {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="${relPrefix}1" Type="${R_NS}/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
<Relationship Id="${relPrefix}2" Type="${R_NS}/slide" Target="../slides/slide${slideNo}.xml"/>
</Relationships>`
}
