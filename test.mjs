// dsh-office development test harness: exercises the registered tools
// directly (create → read → edit → convert) plus real-world file reads.
import { apply, cleanVisionText, ocrPageBodies, ocrSourceLabel, visionConcurrency, visionBudget, pdfMemoStats, sweepRenderCache, chunkLocalBatches,
  sanitizeTextForReturn, losslessJsonProblem, finalizeToolValue, textQuality, readSidecarPath,
  capWithOffset, ocrCachePath, cacheDirState, renderDirFor, ocrRetryScales, currentRenderScale,
  sidecarCoverage, PDF_PARSER_VERSION, pdfOutputGate, registerSourceIdentity, pathKeyOf,
  writeFileAtomic, writeFileAtomicSync, pdfGateStagingRoot, runRenderScript } from './index.js'
import { scanBadCodePoints, utf8RoundTrips, isNfc, describeStringProblems } from './repro.mjs'
import { readPdfText, writePdf, PdfFile, matrixMul } from './pdf.js'
import { markdownToDocument } from './model.js'
import { parseXML } from './xml.js'
import { readOdt } from './odf.js'
import { hash2B, rc4 } from './pdfcrypt.js'
import { openZip, makeZip, maxEntryBytes, maxTotalBytes, zip64ExtraField } from './zip.js'
import { randomBytes, createHash, createCipheriv } from 'node:crypto'
import { findEngine, parseResultLine, gateResult, boxesToText, engineArgs, LOCAL_MAX_IMAGES } from './rapidocr.js'
import { mkdir, rm, readFile, copyFile } from 'node:fs/promises'
import { existsSync, readFileSync, writeFileSync, readdirSync, statSync, rmSync, utimesSync, mkdirSync, chmodSync, openSync, closeSync, renameSync } from 'node:fs'
import { spawn, spawnSync } from 'node:child_process'
import { once } from 'node:events'
import { homedir, tmpdir } from 'node:os'
import { resolve, join, basename, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readPngInfo } from './png.js'
import { inflateRawSync, deflateRawSync, deflateSync, createDeflateRaw, crc32 as zlibCrc32 } from 'node:zlib'
import { imageSizeFor, DOCX_USABLE_WIDTH_PT } from './docx.js'
import { bmpToRaw, gifToRaw, imageToRaw, pngToRaw } from './image.js'

/** 本文件所在目录 = 插件目录（pdf-render.ps1 与 vendor/ 都在这里）。 */
const HERE = dirname(fileURLToPath(import.meta.url))

let OUT = process.env.DSH_OFFICE_TEST_OUT || resolve(process.cwd(), 'test-out')
const results = []
let failures = 0

function ok(name, cond, detail = '') {
  results.push(`${cond ? 'PASS' : 'FAIL'}  ${name}${detail ? '  — ' + detail : ''}`)
  if (!cond) failures++
  return cond
}

/**
 * The artifacts dir has to be writable. Running `node test.mjs` straight from the
 * plugin dir is the natural thing to do and gets EPERM under a workspace-write
 * sandbox — fall back to the system temp dir instead of dying half-way through.
 */
async function ensureOut() {
  const tried = []
  for (const dir of [OUT, join(tmpdir(), 'dsh-office-test-out')]) {
    try { await rm(dir, { recursive: true, force: true }); await mkdir(dir, { recursive: true }); OUT = dir; return OUT }
    catch (e) { tried.push(`${dir}（${e.code || e.message}）`) }
  }
  throw new Error(`没有可写的测试输出目录：${tried.join('、')}；可用 DSH_OFFICE_TEST_OUT 指定`)
}

// --- capture the registered tool definitions -------------------------------
const captured = new Map()
const fakeCtx = {
  get(name) {
    if (name !== 'tools') return undefined
    return { register: def => { captured.set(def.name, def); return () => {} } }
  },
}
apply(fakeCtx)

const exec = {
  agent: { session: { header: { cwd: process.cwd() } } },
  signal: new AbortController().signal,
}
const call = (tool, args) => {
  const def = captured.get(tool)
  if (!def) throw new Error(`tool ${tool} not registered`)
  return def.execute(args, exec)
}

async function textOf(file) {
  const r = await call('office_read', { path: file, as: 'markdown' })
  return r.content
}

/**
 * 任务一：手写一份**身份相符**的 OCR sidecar 夹具。
 *
 * 新规则下，manifest 缺 `srcpath`/`srcsha256`（或与源文件不符）一律整份作废，
 * 所以夹具必须带上身份 —— 否则测到的是"缓存作废"而不是"缓存命中"，
 * 那就等于把原来那些用例的意图悄悄换掉了。
 */
function writeOcrSidecar(pdfPath, manifestLine, body) {
  const p = ocrCachePath(pdfPath)
  mkdirSync(dirname(p), { recursive: true })
  const id = registerSourceIdentity(pdfPath, readFileSync(pdfPath))
  const identity = id ? `<!-- srcpath: ${id.pathHash} | srcsha256: ${id.contentHash} | srcsize: ${id.size} -->\n` : ''
  writeFileSync(p, `<!-- dsh-office OCR cache -->\n<!-- ${manifestLine} -->\n${identity}${body}`, 'utf8')
  return p
}

// ===========================================================================
// 第三轮（pdf.js 解析层修复）的夹具生成器 —— 零外部依赖、零"用户盘上的样本"依赖
// ===========================================================================

/**
 * **真·缺 `/ToUnicode`** 的乱码 PDF（35 页）。
 *
 * `writePdf` 产出的 Type0 字体是
 * `<<…/Subtype/Type0/BaseFont/STSong-Light/Encoding/UniGB-UCS2-H/DescendantFonts[…]>>`
 * —— **本来就没有 `/ToUnicode`**。所以只要正文里放**私用区码点**（会被编成
 * `<E0A1E0A2…>` 这样的 UTF-16BE 十六进制），读回来必然是 `String.fromCharCode(code)`
 * 的 PUA 字符，质量门就会按"私用区码点占 X%（CID 字体缺 ToUnicode 的典型产物）"判乱码。
 * 这不是"模拟乱码"，这就是那类文档本身。
 *
 * ⚠ 可见文字必须走 Helvetica（base-14 + WinAnsi）：未内嵌的 STSong-Light 在 WinRT 里
 * 渲染不出字形，整页会是**空白**（实测 OCR 全部 code=101），OCR 段就失去意义了。
 * PUA 那一行本来就没有字形，只是文字层里的乱码载体 —— 页面看起来仍然正常。
 */
/**
 * 自造 PNG（零依赖）：给插图用例当夹具，避免依赖用户盘上的图片。
 * colorType 0/2/6（灰度/RGB/RGBA），8-bit、非隔行 —— 与 `png.js` 支持面一致。
 */
const TEST_CRC_TABLE = (() => {
  const t = new Int32Array(256)
  for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[i] = c }
  return t
})()

function makeTestPng(w, h, colorType, pixel) {
  const ch = colorType === 6 ? 4 : colorType === 0 ? 1 : 3
  const crc = buf => {
    let c = 0xffffffff
    for (let i = 0; i < buf.length; i++) c = TEST_CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    return (c ^ 0xffffffff) >>> 0
  }
  const chunk = (type, data) => {
    const out = Buffer.alloc(12 + data.length)
    out.writeUInt32BE(data.length, 0)
    out.write(type, 4, 'latin1')
    data.copy(out, 8)
    out.writeUInt32BE(crc(Buffer.concat([Buffer.from(type, 'latin1'), data])), 8 + data.length)
    return out
  }
  const rows = []
  for (let y = 0; y < h; y++) {
    const r = Buffer.alloc(w * ch)
    for (let x = 0; x < w; x++) pixel(r, x, y, ch)
    rows.push(Buffer.concat([Buffer.from([0]), r]))
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8
  ihdr[9] = colorType
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(Buffer.concat(rows), { level: 6 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

/** 自造 8-bit 索引 PNG（colorType 3 + PLTE/tRNS）：第二轮需求 4a 的调色板解码夹具。 */
function makeIndexedPng(w, h, palette, indices, trns = null) {
  const crc = buf => {
    let c = 0xffffffff
    for (let i = 0; i < buf.length; i++) c = TEST_CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    return (c ^ 0xffffffff) >>> 0
  }
  const chunk = (type, data) => {
    const out = Buffer.alloc(12 + data.length)
    out.writeUInt32BE(data.length, 0)
    out.write(type, 4, 'latin1')
    data.copy(out, 8)
    out.writeUInt32BE(crc(Buffer.concat([Buffer.from(type, 'latin1'), data])), 8 + data.length)
    return out
  }
  const rows = []
  for (let y = 0; y < h; y++) {
    const r = Buffer.alloc(w + 1)
    for (let x = 0; x < w; x++) r[x + 1] = indices[y * w + x] & 0xff
    rows.push(r)
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8
  ihdr[9] = 3
  const plte = Buffer.alloc(palette.length * 3)
  palette.forEach((c, i) => { plte[i * 3] = c[0]; plte[i * 3 + 1] = c[1]; plte[i * 3 + 2] = c[2] })
  const parts = [
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('PLTE', plte),
  ]
  if (trns) parts.push(chunk('tRNS', Buffer.from(trns)))
  parts.push(chunk('IDAT', deflateSync(Buffer.concat(rows), { level: 6 })), chunk('IEND', Buffer.alloc(0)))
  return Buffer.concat(parts)
}

/** 自造 24-bit BMP（自下而上、BI_RGB）：第二轮需求 4b 的 BMP 解码夹具。 */
function makeTestBmp24(w, h, pixel) {
  const rowSize = Math.ceil((w * 3) / 4) * 4
  const dataSize = rowSize * h
  const buf = Buffer.alloc(54 + dataSize)
  buf.write('BM', 0, 'latin1')
  buf.writeUInt32LE(54 + dataSize, 2)
  buf.writeUInt32LE(54, 10)
  buf.writeUInt32LE(40, 14)
  buf.writeInt32LE(w, 18)
  buf.writeInt32LE(h, 22)
  buf.writeUInt16LE(1, 26)
  buf.writeUInt16LE(24, 28)
  buf.writeUInt32LE(0, 30)
  buf.writeUInt32LE(dataSize, 34)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const [r, g, b] = pixel(x, y)
      const o = 54 + (h - 1 - y) * rowSize + x * 3
      buf[o] = b; buf[o + 1] = g; buf[o + 2] = r
    }
  }
  return buf
}

/**
 * 自造 GIF（首帧，含可选透明索引）：LZW 位流用"每个像素前发一个 Clear"的**合法最简形态**
 * —— 解码器每次 clear 后字典长度恒定、code size 不变，编码端无需实现字典同步
 * （压缩率差，但对夹具足够，且产物是标准 GIF）。
 */
function makeTestGif(w, h, palette, indices, transparentIndex = -1) {
  const bits = Math.max(1, Math.ceil(Math.log2(Math.max(2, palette.length))))
  const count = 1 << bits
  const gct = Buffer.alloc(count * 3)
  palette.forEach((c, i) => { gct[i * 3] = c[0]; gct[i * 3 + 1] = c[1]; gct[i * 3 + 2] = c[2] })
  const parts = [Buffer.from('GIF89a', 'latin1')]
  const lsd = Buffer.alloc(7)
  lsd.writeUInt16LE(w, 0)
  lsd.writeUInt16LE(h, 2)
  lsd[4] = 0x80 | ((bits - 1) & 7)
  parts.push(lsd, gct)
  if (transparentIndex >= 0) parts.push(Buffer.from([0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, transparentIndex & 0xff, 0x00]))
  const desc = Buffer.alloc(10)
  desc[0] = 0x2c
  desc.writeUInt16LE(w, 5)
  desc.writeUInt16LE(h, 7)
  parts.push(desc)
  const minCode = Math.max(2, bits)
  const clear = 1 << minCode
  const eoi = clear + 1
  const codeSize = minCode + 1
  const codes = []
  for (const idx of indices) codes.push(clear, idx)
  codes.push(eoi)
  const bytes = []
  let acc = 0
  let accBits = 0
  for (const c of codes) {
    acc |= c << accBits
    accBits += codeSize
    while (accBits >= 8) { bytes.push(acc & 0xff); acc >>= 8; accBits -= 8 }
  }
  if (accBits > 0) bytes.push(acc & 0xff)
  parts.push(Buffer.from([minCode]))
  for (let i = 0; i < bytes.length; i += 255) {
    const block = bytes.slice(i, i + 255)
    parts.push(Buffer.from([block.length]), Buffer.from(block))
  }
  parts.push(Buffer.from([0x00, 0x3b]))
  return Buffer.concat(parts)
}

function buildGarbledFixture({ pages = 35 } = {}) {
  const blocks = []
  for (let p = 1; p <= pages; p++) {
    if (p > 1) blocks.push({ type: 'pagebreak' })
    blocks.push({ type: 'heading', level: 2, text: `Fixture page ${p}` })
    blocks.push({ type: 'paragraph', runs: [{ text: `This is fixture page ${p} of the text layer gate.` }] })
    blocks.push({ type: 'paragraph', runs: [{ text: '\uE0A1\uE0A2\uE0A3'.repeat(30) }] })
  }
  return writePdf({ kind: 'document', meta: { title: 'text layer fixture' }, blocks }, {})
}

/**
 * 最小 ObjStm 夹具：把一个对象塞进 `/ObjStm` 且 `First > 0`。
 * 这是 off-by-`First` 的**充要触发条件**（`First === 0` 时 `first + off === off`，新老代码
 * 逐字一致），也是"病灶样本①"会消失之后唯一能长期钉住这条修复的凭据。
 *
 * @returns {{buf: Buffer, first: number, fontBody: string, objStmText: string}}
 */
function minimalObjStmPdf({ first = 4, corruptOffset = false, omitFirst = false } = {}) {
  const fontBody = '<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>'
  // 头部索引：`objNum offset`，offset 是**相对 First** 的区域偏移（本夹具里只有 1 个对象 → 0）
  const head = '4 0 '
  const objStmText = (omitFirst ? '' : head) + fontBody
  const n = corruptOffset ? '4 999999 ' : head
  const body = (omitFirst ? '' : n) + fontBody
  const objs = {}
  objs[1] = '<</Type/Catalog/Pages 2 0 R>>'
  objs[2] = '<</Type/Pages/Kids[3 0 R]/Count 1>>'
  objs[3] = '<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>'
  const content = 'BT /F1 12 Tf 1 0 0 1 20 100 Tm (ObjStm works) Tj ET'
  objs[5] = `<</Length ${content.length}>>\nstream\n${content}\nendstream`
  objs[6] = `<</Type/ObjStm/N 1${omitFirst ? '' : `/First ${first}`}/Length ${body.length}>>\nstream\n${body}\nendstream`
  let out = '%PDF-1.4\n'
  // 注意：**故意跳过 4** —— 对象 4 只能存在于 ObjStm 里。若在此写出 `4 0 obj\nundefined\nendobj`，
  // `scanObjects()` 会先扫到这个真的 obj 4，`expandObjStms()` 又因 `objs.has(objNum)` 跳过它，
  // 夹具就根本走不到被测的 off-by-First 路径（第三轮 FAIL-1/2/3 的根因）。
  for (const i of [1, 2, 3, 5, 6]) out += `${i} 0 obj\n${objs[i]}\nendobj\n`
  out += 'trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n0\n%%EOF\n'
  return { buf: Buffer.from(out, 'latin1'), first, fontBody, objStmText }
}

/**
 * 在**有限时间**内跑一段同步代码（任务 C 的死循环回归）。
 * 必须走子进程：那个死循环是**同步**的，跑在事件循环上，`Promise.race` 连超时都触发不了
 * —— 主进程里根本没有办法把它抢回来。
 */
function runWithTimeout(seconds, code) {
  const r = spawnSync(process.execPath, ['--input-type=module', '-e', code], {
    stdio: 'ignore', timeout: seconds * 1000,
  })
  return { hung: !!r.error && r.error.code === 'ETIMEDOUT', status: r.status, timedOut: r.signal === 'SIGTERM' }
}

/**
 * R18 任务 A：找一个**非 %TEMP%** 且可写的目录，用来验证"目标目录不在 %TEMP% 时 PDF 也能产出"。
 * 刻意排除两处：%TEMP%（那正是要脱离的条件）与插件目录自身（否则会把测试产物写进**目标目录**的
 * `test-out/`）。找不到就返回 null，由调用方打印跳过原因（绝不放宽断言）。
 */
function nonTempWritableDir() {
  const tmpRoot = resolve(tmpdir()).toLowerCase()
  const here = resolve(HERE).toLowerCase()
  const cands = [
    process.env.DSH_OFFICE_TEST_NONTEMP && join(process.env.DSH_OFFICE_TEST_NONTEMP, 'dsh-office-r18'),
    join(homedir(), '.dsh', 'tmp', 'dsh-office-r18-nontemp'),
    join(resolve(process.cwd()), 'r18-nontemp'),
  ].filter(Boolean)
  for (const d of cands) {
    const r = resolve(d)
    const l = r.toLowerCase()
    if (l.startsWith(tmpRoot) || l.startsWith(here)) continue
    try {
      mkdirSync(r, { recursive: true })
      const probe = join(r, '.write-probe')
      writeFileSync(probe, 'x')
      rmSync(probe, { force: true })
      return r
    } catch { /* 试下一个候选 */ }
  }
  return null
}

async function main() {
  await ensureOut()

  ok('插件注册了 4 个工具', captured.size === 4, [...captured.keys()].join(', '))
  for (const t of ['office_read', 'office_create', 'office_edit', 'office_convert']) ok(`工具 ${t} 已注册`, captured.has(t))

  const MD = `# 季度经营报告

本季度整体**营收增长**明显，成本控制良好。

## 关键指标

- 营收 1280 万元
- 毛利率 42.5%
- 客户留存率 91%

## 明细表

| 区域 | 营收 | 同比 |
| --- | --- | --- |
| 华北 | 520 | +12% |
| 华东 | 460 | +8% |
| 华南 | 300 | -3% |

\`\`\`
calc: 520 + 460 + 300 = 1280
\`\`\`

> 结论：下季度继续投入华东市场。
`

  // ---------------- create + read round trips ----------------
  const docx = join(OUT, 'report.docx')
  let r = await call('office_create', { path: docx, markdown: MD })
  ok('创建 docx', existsSync(docx) && r.bytes > 3000, `${r.bytes} bytes`)
  let txt = await textOf(docx)
  ok('读取 docx 标题', txt.includes('季度经营报告'))
  ok('读取 docx 列表', txt.includes('营收 1280 万元'))
  ok('读取 docx 表格', txt.includes('华北') && txt.includes('+12%'))
  ok('读取 docx 引用', txt.includes('下季度继续投入华东市场'))

  const xlsx = join(OUT, 'data.xlsx')
  r = await call('office_create', {
    path: xlsx,
    workbook: {
      meta: { title: '成绩表' },
      sheets: [
        { name: '成绩', columns: [{ title: '姓名', width: 14 }, { title: '分数', width: 10 }, { title: '科目', width: 12 }], rows: [['姓名', '分数', '科目'], ['张三', 95, '数学'], ['李四', 88.5, '语文'], ['王五', 72, '英语']] },
        { name: '汇总', rows: [['平均分', 85.17], ['人数', 3]] },
      ],
    },
  })
  ok('创建 xlsx 双表', existsSync(xlsx) && r.summary.includes('成绩'), r.summary)
  txt = await textOf(xlsx)
  ok('读取 xlsx 中文与数字', txt.includes('张三') && txt.includes('95') && txt.includes('88.5'))
  ok('读取 xlsx 第二个表', txt.includes('汇总') && txt.includes('85.17'))
  const sheetOnly = await call('office_read', { path: xlsx, sheet: '汇总' })
  ok('按工作表名读取', sheetOnly.content.includes('85.17') && !sheetOnly.content.includes('张三'))

  // columns 声明表头时，应落成表格首行（而非仅列宽）
  const xlsx2 = join(OUT, 'header.xlsx')
  await call('office_create', { path: xlsx2, workbook: { sheets: [{ name: 'S', columns: [{ title: '姓名', width: 12 }, { title: '分数', width: 8 }], rows: [['张三', 95]] }] } })
  txt = await textOf(xlsx2)
  ok('columns 自动生成表头行', txt.includes('姓名') && txt.includes('分数') && txt.includes('张三'))
  const hdrJson = await call('office_read', { path: xlsx2, as: 'json' })
  const hdrRows = JSON.parse(hdrJson.content).sheets[0].rows
  ok('表头行为第一行', String(hdrRows[0][0].v) === '姓名' && String(hdrRows[1][0].v) === '张三')

  // limit 应被尊重（下限 100）
  const smallRead = await call('office_read', { path: docx, limit: 200 })
  ok('limit 参数生效', smallRead.content.length <= 400, `${smallRead.content.length} 字符`)

  // as=meta 的 render 必须输出元数据 JSON（而不是 "null"）
  const metaRead = await call('office_read', { path: xlsx, as: 'meta' })
  const metaRendered = captured.get('office_read').output.render({ path: xlsx, as: 'meta' }, metaRead)[0].text
  ok('as=meta 渲染元数据', metaRendered.includes('sheets') && metaRendered !== 'null', metaRendered.slice(0, 80).replace(/\n/g, ' '))

  const pptx = join(OUT, 'deck.pptx')
  r = await call('office_create', {
    path: pptx,
    slides: {
      meta: { title: '产品发布' },
      slides: [
        { layout: 'title', title: '2026 产品发布会', subtitle: 'DSH Office 能力演示' },
        { layout: 'content', title: '本次亮点', bullets: [{ text: '零第三方依赖', level: 0 }, { text: '全格式互转', level: 0 }, { text: '支持中文 PDF', level: 1 }], notes: '演示时强调中文支持' },
        { layout: 'content', title: '定价对比', table: { rows: [['方案', '价格', '席位'], ['基础版', '免费', '1'], ['团队版', '99/月', '20']] } },
      ],
    },
  })
  ok('创建 pptx 三页', existsSync(pptx) && r.summary.includes('3 页'), r.summary)
  txt = await textOf(pptx)
  ok('读取 pptx 标题页', txt.includes('2026 产品发布会') && txt.includes('DSH Office 能力演示'))
  ok('读取 pptx 多级要点', txt.includes('零第三方依赖') && txt.includes('支持中文 PDF'))
  ok('读取 pptx 表格', txt.includes('团队版') && txt.includes('99/月'))
  ok('读取 pptx 备注', txt.includes('演示时强调中文支持'))

  const pdf = join(OUT, 'report.pdf')
  r = await call('office_create', { path: pdf, markdown: MD })
  ok('创建 pdf(含中文)', existsSync(pdf) && r.bytes > 2000, `${r.bytes} bytes`)
  const pdfRead = await call('office_read', { path: pdf })
  ok('PDF 回读含中文标题', pdfRead.content.includes('季度经营报告'), '中文字体 STSong 映射')
  ok('PDF 回读含数字', pdfRead.content.includes('1280'))
  ok('PDF 元数据', pdfRead.stats?.pages >= 1, `pages=${pdfRead.stats?.pages}`)

  // ---------------- 扫描件 / OCR 路径 ----------------
  const blankPdf = join(OUT, 'scan-like.pdf')
  await call('office_create', { path: blankPdf, document: { blocks: [{ type: 'hr' }] } })
  const blankMeta = await call('office_read', { path: blankPdf, as: 'meta' })
  ok('无文本层 PDF 被标记为扫描页', blankMeta.stats.pagesWithText === 0 && blankMeta.stats.scannedPages.length === 1, JSON.stringify(blankMeta.stats.scannedPages))

  let ocrThrew = false
  try { await call('office_read', { path: blankPdf, ocr: 'never' }) } catch (e) { ocrThrew = /文本层/.test(e.message) }
  ok('ocr=never 明确提示需要 OCR', ocrThrew)

  const autoRead = await call('office_read', { path: blankPdf })
  ok('ocr=auto 无视觉桥时优雅降级', typeof autoRead.content === 'string' && /扫描|OCR/.test(autoRead.content), autoRead.content.replace(/\n/g, ' ').slice(0, 100))

  // P5：扫件 as="json" 只有页级空壳，必须自报原因而不是静默返回空
  const scanJson = await call('office_read', { path: blankPdf, as: 'json' })
  ok('P5 扫描件 as=json 带提示字段', /"notice":/.test(scanJson.content) && /无文本层/.test(scanJson.content) && /ocr:\\"always\\"/.test(scanJson.content), scanJson.content.slice(0, 80).replace(/\n/g, ' '))
  ok('P5 json stats 给出 pagesWithText', scanJson.stats.pagesWithText === 0, JSON.stringify(scanJson.stats))
  const textJson = await call('office_read', { path: join(OUT, 'report.pdf'), as: 'json' })
  ok('P5 有文本层的 PDF 不加多余提示', !/"notice":/.test(textJson.content))

  // 预置 OCR 缓存 → 直接命中缓存，不再需要视觉桥
  // （任务 D 起 manifest 必须带 `parser:`；任务一起还必须带**源身份** srcpath/srcsha256，
  //   缺失或不符会被整份作废，见下面的"缓存迁移 / 缓存身份"段）
  writeOcrSidecar(blankPdf, `covered: 1 | parser: ${PDF_PARSER_VERSION}`,
    '## 第 1 页（OCR）\n缓存识别文本：sample-ocr-text\n')
  const cachedRead = await call('office_read', { path: blankPdf, pages: '1' })
  ok('OCR 缓存命中并标注来源', cachedRead.content.includes('缓存识别文本') && cachedRead.content.includes('OCR 缓存'), cachedRead.content.replace(/\n/g, ' ').slice(0, 100))
  ok('OCR 页计入 stats.ocrPages', Array.isArray(cachedRead.stats.ocrPages) && cachedRead.stats.ocrPages.includes(1), JSON.stringify(cachedRead.stats.ocrPages))

  // 回归：缓存回读必须拿回整页（旧版 /m + $ 的正则只能拿回第一行）
  writeOcrSidecar(blankPdf, `covered: 1-2 | parser: ${PDF_PARSER_VERSION}`,
    '## 第 1 页（OCR）\n第一行内容\n第二行内容\n第三行内容\n\n## 第 2 页（OCR）\n另一页文本\n')
  const multiCached = await call('office_read', { path: blankPdf, pages: '1' })
  ok('OCR 缓存回读保留整页多行', multiCached.content.includes('第一行内容') && multiCached.content.includes('第三行内容'), multiCached.content.replace(/\n/g, ' ').slice(0, 90))

  // P1：指定了引擎却吃到别家缓存 → 页脚 + 脚注都要说清"指定的没跑"
  const visionWanted = await call('office_read', { path: blankPdf, pages: '1', ocr: 'always', ocrEngine: 'vision' })
  ok('P1 异源命中：页脚写明缓存来源与未执行的引擎',
    /OCR 缓存 · 来源未记录（指定的 vision 未执行）/.test(visionWanted.content),
    (visionWanted.content.match(/_（第 1 页：.*?）_/) || ['<无标注>'])[0])
  ok('P1 异源命中：脚注给出删除 sidecar 的路径',
    /无法确认指定的 vision 是否跑过/.test(visionWanted.content) && /如需强制重识别请删/.test(visionWanted.content),
    (visionWanted.content.match(/^> .*$/m) || ['<无脚注>'])[0].slice(0, 120))

  // P2：旧 sidecar 里的视觉客套话必须在命中缓存时洗掉并回写
  writeOcrSidecar(blankPdf, `covered: 1 | parser: ${PDF_PARSER_VERSION}`,
    '# OCR 文本（由 dsh-office 生成，可安全删除以重新识别）\n\n## 第 1 页（OCR）\n该图片展示了一页文档，内容为法律条文。\n\nTranscription:\n缓存正文甲\n\nUncertain: 底部被截断\n')
  const healed = await call('office_read', { path: blankPdf, pages: '1', ocr: 'always' })
  const healedFile = readFileSync(join(OUT, 'scan-like.ocr.md'), 'utf8')
  ok('P2 命中缓存时自动清洗并回写 sidecar',
    healedFile.includes('缓存正文甲') && !/Transcription:|Uncertain:|该图片展示/.test(healedFile),
    JSON.stringify(healedFile.split('\n').slice(3).join(' | ').slice(0, 90)))
  ok('P0 回写时补上 covered/total manifest',
    /^<!-- covered: 1 \| total: 1 \| parser: \d+ -->$/m.test(healedFile), healedFile.split('\n')[1])
  ok('P2 回洗后的正文不含客套话', !/Transcription:|Uncertain:|该图片展示/.test(healed.content))

  // ---------------- 本地引擎 RapidOCR-json ----------------
  const eng = findEngine()
  ok('引擎发现（vendor / DSH_OFFICE_RAPIDOCR_DIR / PATH）', true, eng ? `已装：${eng.dir}` : '未安装 → 仅验证视觉桥降级')
  ok('解析 code:100 结果行（文本 + 置信）', (() => {
    const p = parseResultLine('{"code":100,"data":[{"box":[[349,175],[556,175],[556,228],[349,228]],"score":0.978,"text":"sample-topic"},{"box":[[216,332],[690,335],[689,446],[215,444]],"score":0.992,"text":"sample-ocr-text"}]}')
    return p.code === 100 && p.boxes.length === 2 && p.avg > 0.97 && p.text === 'sample-topic\nsample-ocr-text'
  })())
  ok('code:101 判为空页且不回落视觉', (() => {
    const v = gateResult(parseResultLine('{"code":101,"data":"No text found in image."}'))
    return v.blank === true && v.hard === false
  })())
  ok('忽略横幅行与非结果噪声', parseResultLine('RapidOCR-json v1.1.0') === null && parseResultLine('') === null)
  ok('质量门：低置信/复杂版面判疑难', !!gateResult({ code: 100, boxes: new Array(60).fill(0), text: 'x', avg: 0.9, shortRatio: 0.2 }).hard && !!gateResult({ code: 100, boxes: new Array(6).fill(0), text: '一', avg: 0.7, shortRatio: 0.9 }).hard)
  ok('质量门：清晰整页放行', gateResult({ code: 100, boxes: new Array(12).fill(0), text: '聚焦国省事考纲创新理论和方针政策', avg: 0.98, shortRatio: 0.4 }).hard === false)

  // ---- 第八轮（自 WB 侧回移）：短句判据只在低置信时判硬 + retryable 分流 ----
  // 病灶：表格/数字页天然由大量 ≤6 字片段组成，旧的"短句占比高 → 判硬"会把置信度 0.99、
  // 文本完全可读的页丢成空字符串；而且它被当成"可换倍率救"的失败，让重试段白跑。
  // 新行为：只在 avg < shortRatioScore 时判硬，且 retryable=false（换倍率救不了结构特征）。
  {
    // 35 框（< hardBoxes=45，避开"版面复杂"）、avg 0.90（≥ minScore=0.88，避开"置信度低"）
    // —— 必须让"短句占比高"这个分支真的被走到，否则这条断言什么都没验。
    const shortLow = gateResult({ code: 100, boxes: new Array(35).fill(0), text: 'x', avg: 0.9, shortRatio: 0.9 })
    ok('质量门：短句占比高（低置信）判硬且标 retryable=false',
      shortLow.hard === true && shortLow.retryable === false && /短句占比高/.test(String(shortLow.reason)),
      JSON.stringify(shortLow))
    // 新行为核心（能区分新旧实现）：同样碎、但认得很准（0.96 ≥ shortRatioScore 0.95）→ 放行
    const shortHigh = gateResult({ code: 100, boxes: new Array(35).fill(0), text: 'x', avg: 0.96, shortRatio: 0.9 })
    ok('质量门：短句占比高但高置信 → 放行（表格/数字页不再被误杀）', shortHigh.hard === false, JSON.stringify(shortHigh))
    // 非短句的 hard 分支：换倍率有意义，retryable 必须恒为 true
    const lowScore = gateResult({ code: 100, boxes: new Array(6).fill(0), text: '一', avg: 0.5, shortRatio: 0.1 })
    const denseBoxes = gateResult({ code: 100, boxes: new Array(60).fill(0), text: 'x', avg: 0.9, shortRatio: 0.2 })
    ok('质量门：置信度低 / 版面复杂 两个 hard 分支 retryable 恒为 true',
      lowScore.hard === true && lowScore.retryable === true && /置信度低/.test(String(lowScore.reason))
      && denseBoxes.hard === true && denseBoxes.retryable === true && /版面复杂/.test(String(denseBoxes.reason)),
      `${JSON.stringify(lowScore)} | ${JSON.stringify(denseBoxes)}`)
    // 阈值语义：GATE 是模块级常量，同进程改 process.env 不会生效 → 用第二参数注入自定义 gate
    const G = { minScore: 0.88, hardBoxes: 45, hardBoxesScore: 0.95, shortRatio: 0.6, shortMinBoxes: 30, shortRatioScore: 0.95 }
    const r35 = { code: 100, boxes: new Array(35).fill(0), text: 'x', avg: 0.9, shortRatio: 0.9 }
    ok('质量门阈值：shortRatio 是严格大于（占比恰好等于阈值不判硬）',
      gateResult({ ...r35, shortRatio: 0.6 }, G).hard === false && gateResult({ ...r35, shortRatio: 0.61 }, G).hard === true)
    ok('质量门阈值：shortMinBoxes 是"框数 ≥"（少一个框就不判）',
      gateResult({ code: 100, boxes: new Array(30).fill(0), text: 'x', avg: 0.9, shortRatio: 0.9 }, G).hard === true
      && gateResult({ code: 100, boxes: new Array(29).fill(0), text: 'x', avg: 0.9, shortRatio: 0.9 }, G).hard === false)
    ok('质量门阈值：shortRatioScore 是"avg <"（等于即视为认得很准，放行）',
      gateResult({ ...r35, avg: 0.95 }, G).hard === false && gateResult({ ...r35, avg: 0.949 }, G).hard === true)
    ok('质量门阈值：三个变量任一收紧都能让短句分支不触发（联动而非各自为政）',
      gateResult(r35, { ...G, shortRatio: 0.95 }).hard === false
      && gateResult(r35, { ...G, shortMinBoxes: 40 }).hard === false
      && gateResult(r35, { ...G, shortRatioScore: 0.9 }).hard === false)
    // GATE 确实读 process.env：改 env 后必须**重新加载模块实例**才看得到（同进程改 env 无效）
    const envGuard = {}
    const withEnv = kv => {
      for (const [k, v] of Object.entries(kv)) {
        if (!(k in envGuard)) envGuard[k] = process.env[k]
        if (v === undefined) delete process.env[k]; else process.env[k] = v
      }
    }
    try {
      let reloadSeq = 0
      const reload = async () => {
        reloadSeq += 1
        return import(new URL(`./rapidocr.js?gate-env=${reloadSeq}`, import.meta.url).href)
      }
      withEnv({ DSH_OFFICE_OCR_SHORT_RATIO: '0.95', DSH_OFFICE_OCR_SHORT_MIN_BOXES: undefined, DSH_OFFICE_OCR_SHORT_SCORE: undefined })
      const gRatio = await reload()
      ok('质量门阈值：DSH_OFFICE_OCR_SHORT_RATIO 生效（提到 0.95 → 占比 0.9 不再判硬）',
        gRatio.GATE.shortRatio === 0.95 && gRatio.GATE.shortMinBoxes === 30 && gRatio.GATE.shortRatioScore === 0.95
        && gRatio.gateResult(r35).hard === false,
        JSON.stringify({ shortRatio: gRatio.GATE.shortRatio, hard: gRatio.gateResult(r35).hard }))
      withEnv({ DSH_OFFICE_OCR_SHORT_RATIO: undefined, DSH_OFFICE_OCR_SHORT_MIN_BOXES: '40' })
      const gBoxes = await reload()
      ok('质量门阈值：DSH_OFFICE_OCR_SHORT_MIN_BOXES 生效（提到 40 → 35 框不再判硬）',
        gBoxes.GATE.shortMinBoxes === 40 && gBoxes.gateResult(r35).hard === false,
        JSON.stringify({ shortMinBoxes: gBoxes.GATE.shortMinBoxes, hard: gBoxes.gateResult(r35).hard }))
      withEnv({ DSH_OFFICE_OCR_SHORT_MIN_BOXES: undefined, DSH_OFFICE_OCR_SHORT_SCORE: '0.90' })
      const gScore = await reload()
      ok('质量门阈值：DSH_OFFICE_OCR_SHORT_SCORE 生效（收到 0.90 → avg 0.9 视为高置信放行）',
        gScore.GATE.shortRatioScore === 0.9 && gScore.gateResult(r35).hard === false,
        JSON.stringify({ shortRatioScore: gScore.GATE.shortRatioScore, hard: gScore.gateResult(r35).hard }))
      withEnv({ DSH_OFFICE_OCR_SHORT_SCORE: 'abc', DSH_OFFICE_OCR_SHORT_RATIO: '0', DSH_OFFICE_OCR_SHORT_MIN_BOXES: '0' })
      const gFallback = await reload()
      ok('质量门阈值：0 / 非法值一律回落默认（Number(env) || 默认）',
        gFallback.GATE.shortRatio === 0.6 && gFallback.GATE.shortMinBoxes === 30 && gFallback.GATE.shortRatioScore === 0.95,
        JSON.stringify({ shortRatio: gFallback.GATE.shortRatio, shortMinBoxes: gFallback.GATE.shortMinBoxes, shortRatioScore: gFallback.GATE.shortRatioScore }))
    } finally {
      for (const [k, v] of Object.entries(envGuard)) { if (v === undefined) delete process.env[k]; else process.env[k] = v }
    }
  }

  ok('boxesToText 按 y 分行、行内按 x 排序', boxesToText([
    { text: 'B', box: [[0, 30], [9, 30], [9, 40], [0, 40]] },
    { text: 'A', box: [[0, 5], [9, 5], [9, 15], [0, 15]] },
  ]).split('\n').join('|') === 'A|B')

  // 视觉兜底的输出会原样进 .ocr.md 缓存，模型客套话必须在写盘前剪掉
  ok('视觉清洗：剥掉描述段 + Transcription 标记行', cleanVisionText(
    '该图片为一份文档的第5页，主要包含目录内容。\n\nTranscription:\n21. 负面政策类母题 ...... 38\n22. 领导讲话类母题 ...... 40',
  ) === '21. 负面政策类母题 ...... 38\n22. 领导讲话类母题 ...... 40')
  ok('视觉清洗：markdown 围栏只去围栏行', cleanVisionText('```markdown\n# 标题\n正文一\n```') === '# 标题\n正文一')
  ok('视觉清洗：无标记的裸描述最多削几行，正文照旧', cleanVisionText('图片显示了一张表格\n项目 数量\n甲 3') === '项目 数量\n甲 3')
  ok('视觉清洗：宁可不删——没有客套话时逐字原样返回', cleanVisionText(
    '第一，对于主任关键是要尊重服从\n（以上为解决问题）\n第二，对于副主任要抓住机会多请教') ===
    '第一，对于主任关键是要尊重服从\n（以上为解决问题）\n第二，对于副主任要抓住机会多请教')
  ok('视觉清洗：正文里的"以下为……："引导句不当客套话剪掉', cleanVisionText(
    '以下为具体安排：\n上午 研讨\n下午 观摩') === '以下为具体安排：\n上午 研讨\n下午 观摩')
  // 多 tile 拼接（一页切 2/4/8 片后 join）会把每一片的图说/标记都拼进来：
  // 只有开头那一片能被"开头规则"剪掉，第二片起必须靠逐片规则
  ok('视觉清洗：第二片起的图说 + Transcription + Uncertain 也剪掉', cleanVisionText(
    '第一片正文甲\n\nUncertain: 底部被截断\n这是一页中文法律辅导材料，主要内容是类型表格。\n\nTranscription:\n第二片正文乙',
  ) === '第一片正文甲\n\n第二片正文乙')
  ok('视觉清洗：没有标记行时不碰正文（提到"表格"也不剪）', cleanVisionText(
    '本页表格说明了三件事。\n甲 1\n乙 2') === '本页表格说明了三件事。\n甲 1\n乙 2')
  ok('视觉清洗幂等（缓存回洗不会越洗越少）', (() => {
    const once = cleanVisionText('该图片展示了一页文档。\n\nTranscription:\n正文甲\n\nUncertain: 底部截断\n正文乙')
    return once === '正文甲\n\n正文乙' && cleanVisionText(once) === once
  })())
  // P3/P4：页脚必须能看出"切了几片/调了几次"和"是谁让它走视觉的"
  ok('P3 页脚标注视觉切片数与调用次数', ocrSourceLabel({ engine: 'vision', via: 'user', bands: 4, calls: 7 })
    === '视觉模型识别 · 用户指定 ocrEngine:"vision" · 切 4 片（7 次调用）')
  ok('P4 用户指定 / 自动复核 / 本地缺失 三种措辞各不相同', (() => {
    const user = ocrSourceLabel({ engine: 'vision', via: 'user' })
    const review = ocrSourceLabel({ engine: 'vision', via: 'review' })
    const nolocal = ocrSourceLabel({ engine: 'vision', via: 'no-local' })
    return user.includes('用户指定') && review.includes('本地复核') && nolocal.includes('本地引擎不可用')
      && !review.includes('用户指定') && new Set([user, review, nolocal]).size === 3
  })())
  // P1：命中缓存要说清是谁的结果，且 auto 路径维持原来的短标注（不加负担）
  ok('P1 缓存来源三态：同源 / 异源 / 未记录', (() => {
    const auto = ocrSourceLabel({ cached: true, src: 'rapidocr' }, 'auto')
    const same = ocrSourceLabel({ cached: true, src: 'rapidocr' }, 'local')
    const cross = ocrSourceLabel({ cached: true, src: 'rapidocr' }, 'vision')
    const unknown = ocrSourceLabel({ cached: true }, 'vision')
    return auto === 'OCR 缓存' && same === 'OCR 缓存 · local'
      && cross === 'OCR 缓存 · local（指定的 vision 未执行）'
      && unknown.includes('来源未记录')
  })())
  // 逐页切分是缓存比对的基础：标注里带全角嵌套括号，正则必须贪婪到行尾
  ok('逐页切分：按来源标注切回逐页正文（含嵌套括号标注）', (() => {
    const m = ocrPageBodies('甲页正文\n\n_（第 1 页：本地 OCR · RapidOCR（9 框，置信 0.988））_\n\n乙页正文\n\n_（第 2 页：OCR 缓存）_\n\n> 已 OCR 第 1、2 页')
    return m.size === 2 && m.get(1) === '甲页正文' && m.get(2) === '乙页正文'
  })(), '标注之后的脚注不得混进页正文')

  let scan = process.env.DSH_OFFICE_TEST_SCAN_PDF || ''
  if (!existsSync(scan)) {
    try {
      const root = join(homedir(), '.dsh', 'attachments')
      scan = readdirSync(root, { recursive: true }).map(String).filter(f => /\.pdf$/i.test(f))
        .map(f => join(root, f)).filter(f => { try { return statSync(f).size > 5 * 1048576 } catch { return false } })
        // 最近一次的附件才是"当前在测的东西"；按路径哈希目录排序等于随机抽样本
        .sort((a, b) => statSync(b).mtimeMs - statSync(a).mtimeMs)[0] || ''
    } catch { scan = '' }
  }
  if (eng && scan) {
    // 拷一份进产物目录再冷启动：绝不 rmSync 用户真实附件旁边的 .ocr.md（那是识别成果，也不该被测试改状态）
    const sample = join(OUT, 'scan-sample.pdf')
    const userSidecar = scan.replace(/\.pdf$/i, '.ocr.md')
    const userSidecarBefore = existsSync(userSidecar)
    await copyFile(scan, sample)
    const t0 = Date.now()
    const local = await call('office_read', { path: sample, pages: '1-2', ocr: 'always', ocrEngine: 'local' })
    const coldMs = Date.now() - t0
    const secs = (coldMs / 1000).toFixed(1)
    ok('真实扫描件：本地引擎 2 页出文', /[一-龥]{6,}/.test(local.content), `${secs}s · ${basename(scan)}`)
    ok('来源标注为本地 OCR 且计入 stats', /本地 OCR/.test(local.content) && String(local.stats.ocrEngine).includes('RapidOCR'), `${local.stats.ocrEngine} | ${secs}s/2页`)
    ok('ocrEngine:"local" 下不触发视觉兜底', Array.isArray(local.stats.ocrEscalated) && local.stats.ocrEscalated.length === 0, JSON.stringify(local.stats.ocrEscalated))
    ok('sidecar 只写在样本旁边，不碰用户附件', existsSync(join(OUT, 'scan-sample.ocr.md')) && existsSync(userSidecar) === userSidecarBefore, basename(scan))
    const tw = Date.now()
    const warm = await call('office_read', { path: sample, pages: '1-2', ocr: 'always', ocrEngine: 'local' })
    const warmMs = Date.now() - tw
    const cold = ocrPageBodies(local.content)
    const hot = ocrPageBodies(warm.content)
    const shared = [...cold.keys()].filter(p => hot.has(p))
    const cjk = s => (String(s).match(/[一-龥]/g) || []).length
    ok('二次读取命中 sidecar 缓存', /OCR 缓存/.test(warm.content), `${shared.length} 页复用 · ${basename(scan)}`)
    // 断言与语料无关：凡冷读给出过正文的页，缓存必须逐字给回同一页（无文字页不出标注，故按"有的页"计数）
    ok('缓存复用后逐页正文与冷读逐字一致（不截断、不串页）',
      cold.size >= 1 && cold.size === hot.size && shared.length === cold.size
      && shared.every(p => cold.get(p) === hot.get(p)) && shared.some(p => cjk(cold.get(p)) >= 30),
      shared.map(p => `p${p}:${cjk(cold.get(p))}字${cold.get(p) === hot.get(p) ? '=' : '≠'}`).join(' ') || `冷读 0 页正文：${basename(scan)}`)
    // 冷读里含"打开整本 + 栅格化"的固定开销，缓存读只省识别那段，所以看绝对差值而不是倍数
    ok('缓存命中省下识别耗时（不重复消耗识别额度）', warmMs < coldMs && coldMs - warmMs >= 250, `省 ${coldMs - warmMs}ms（warm ${warmMs}ms / cold ${coldMs}ms）`)
    // P0：sidecar 头部必须自报覆盖范围；P1：同源命中要写明是谁的结果
    const sideText = readFileSync(join(OUT, 'scan-sample.ocr.md'), 'utf8')
    const manifest = (sideText.split('\n')[1] || '')
    ok('P0 sidecar manifest 记录 covered/total',
      /^<!-- covered: [\d,\-]+ \| total: \d+ \| parser: \d+( \| src: [^>]+)? -->$/.test(manifest), manifest)
    ok('P0 manifest 的 total 与全书页数一致', Number((/total: (\d+)/.exec(manifest) || [])[1]) === local.stats.pages, `total=${(/total: (\d+)/.exec(manifest) || [])[1]} pages=${local.stats.pages}`)
    ok('P0 stats 自报覆盖范围与剩余页', /^[\d,\-]+ \/ \d+$/.test(String(local.stats.ocrCovered)) && (local.stats.ocrUncovered === undefined || /^[\d,\-]+$/.test(String(local.stats.ocrUncovered))), `${local.stats.ocrCovered} | uncovered=${local.stats.ocrUncovered}`)
    ok('P1 同源命中写明缓存来源', /OCR 缓存 · local/.test(warm.content), (warm.content.match(/_（第 \d+ 页：.*?）_/) || ['<无标注>'])[0])
    ok('P0 再次读取 manifest 不回退', /^<!-- covered: [\d,\-]+ \| total: \d+/.test(readFileSync(join(OUT, 'scan-sample.ocr.md'), 'utf8').split('\n')[1] || ''))
  } else {
    ok('本地引擎端到端（缺引擎或缺样本 → 跳过）', true, eng ? '无真实扫描件样本' : '引擎未安装')
  }

  // ---------------- P1-3a 回归：writePdf 页字典闭合（旧版多一个 '>'） ----------------
  // 旧版每个页面字典写成 `/Resources<< /Font<< … >>>>>`（5 个 '>'）：本机 lenient 解析器容忍，
  // WinRT 的 PdfDocument.LoadFromFileAsync 判其畸形 → 插件产出的 PDF 一律栅格化失败
  // （"One or more errors occurred"）。这里用真栅格化守住这 1 个字符。
  const fixPdf = join(OUT, 'render-fix.pdf')
  await call('office_create', { path: fixPdf, document: { blocks: [{ type: 'hr' }, { type: 'pagebreak' }, { type: 'hr' }] } })
  const fixRaw = readFileSync(fixPdf, 'latin1')
  ok('P1-3 页面字典以 4 个 ">" 闭合（不再多写一个）',
    !fixRaw.includes('>>>>>/Contents') && /\/Resources<< \/Font<<[^>]*>>>>\/Contents/.test(fixRaw))
  if (process.platform === 'win32') {
    const fixDir = join(OUT, 'render-fix-png')
    rmSync(fixDir, { recursive: true, force: true })
    mkdirSync(fixDir, { recursive: true })
    spawnSync('powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(HERE, 'pdf-render.ps1'), fixPdf, fixDir, '1-2'],
      { stdio: 'ignore', windowsHide: true })
    const fp1 = join(fixDir, 'page-1.png'); const fp2 = join(fixDir, 'page-2.png')
    const fsize = f => (existsSync(f) ? statSync(f).size : 0)
    ok('P1-3 修复后 WinRT 能栅格化插件产出的 PDF（2 页 PNG 落盘）',
      fsize(fp1) > 1000 && fsize(fp2) > 1000, `page-1=${fsize(fp1)}B page-2=${fsize(fp2)}B`)
  } else {
    ok('P1-3 WinRT 栅格化（非 Windows → 跳过）', true, `platform=${process.platform}`)
  }

  // ---------------- P1-3b 渲染分辨率旋钮（opt-in，默认行为不变） ----------------
  ok('P1-3b engineArgs：设了 scale 才传 --maxSideLen，且按 1123*scale 对齐 256', (() => {
    const set = v => { if (v === undefined) delete process.env.DSH_OFFICE_RENDER_SCALE; else process.env.DSH_OFFICE_RENDER_SCALE = v; return engineArgs() }
    const side = a => (a.find(x => x.startsWith('--maxSideLen=')) || '（未传）')
    const got = [set(undefined), set('1'), set('1.5'), set('2'), set('abc'), set('0'), set('5')].map(side)
    delete process.env.DSH_OFFICE_RENDER_SCALE
    return got.join('|') === '（未传）|--maxSideLen=1024|--maxSideLen=1792|--maxSideLen=2304|（未传）|（未传）|（未传）'
  })(), '1.5 → 1792；未设/垃圾值/0/超窗(5) 一律不传')
  if (process.platform === 'win32') {
    // 口径：96dpi × scale 的 A4 像素（Size 是 DIP，故 scale=1 即 794×1123），跨机一致
    const wantScale = [[1, 794, 1123], [1.5, 1191, 1685], [2, 1588, 2246]]
    const gotScale = []
    let scaleOk = true
    for (const [scale, ew, eh] of wantScale) {
      const out = join(OUT, `scale-${scale}`)
      rmSync(out, { recursive: true, force: true })
      mkdirSync(out, { recursive: true })
      const env = { ...process.env, DSH_OFFICE_RENDER_SCALE: String(scale) }
      spawnSync('powershell.exe',
        ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(HERE, 'pdf-render.ps1'), fixPdf, out, '0', 'page-1.png'],
        { stdio: 'ignore', windowsHide: true, env })
      const png = join(out, 'page-1.png')
      if (!existsSync(png)) { scaleOk = false; gotScale.push(`scale=${scale}:无 PNG`); continue }
      const d = readPngInfo(readFileSync(png))
      const near = (v, e) => Math.abs(v - e) <= Math.max(2, e * 0.02)
      if (!near(d.width, ew) || !near(d.height, eh)) scaleOk = false
      gotScale.push(`scale=${scale}:${d.width}x${d.height}(期望≈${ew}x${eh})`)
    }
    ok('P1-3b 各档 PNG 尺寸 = 96dpi×scale 的 A4（±2%）', scaleOk, gotScale.join(' · '))
    // 不设 env → 完全不传 options，与旧版逐像素一致（本机 120dpi 原生 = 992×1403）
    const defOut = join(OUT, 'scale-default')
    rmSync(defOut, { recursive: true, force: true })
    mkdirSync(defOut, { recursive: true })
    const envD = { ...process.env }
    delete envD.DSH_OFFICE_RENDER_SCALE
    spawnSync('powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(HERE, 'pdf-render.ps1'), fixPdf, defOut, '1-2'],
      { stdio: 'ignore', windowsHide: true, env: envD })
    const dp1 = join(defOut, 'page-1.png')
    const dp2 = join(defOut, 'page-2.png')
    ok('P1-3b 未设 env 走原生路径（单页/批处理两模式都出图）',
      existsSync(dp1) && existsSync(dp2) && readPngInfo(readFileSync(dp1)).width === readPngInfo(readFileSync(dp2)).width,
      existsSync(dp1) ? `${readPngInfo(readFileSync(dp1)).width}x${readPngInfo(readFileSync(dp1)).height}` : '无 PNG')
  } else {
    ok('P1-3b 各档 PNG 尺寸（非 Windows → 跳过）', true, `platform=${process.platform}`)
    ok('P1-3b 未设 env 走原生路径（非 Windows → 跳过）', true, `platform=${process.platform}`)
  }

  // ---------------- P0-2 视觉升级低并发 + 预算护栏（假视觉桥，0 次真实视觉调用） ----------------
  const sleep = ms => new Promise(r => setTimeout(r, ms))
  /** 假视觉桥：记录在飞峰值/调用页/次数，每页 sleep 120ms 模拟网络往返。 */
  function fakeVision() {
    const state = { inflight: 0, maxInflight: 0, calls: 0, pages: [] }
    return {
      state,
      bridge: {
        output: { render: () => [{ type: 'text', text: '' }] },
        async execute(args) {
          state.inflight++
          state.maxInflight = Math.max(state.maxInflight, state.inflight)
          state.calls++
          const m = /page-(\d+)\.png/.exec(String(args.path))
          state.pages.push(m ? Number(m[1]) : 0)
          try { await sleep(120); return `第${m ? m[1] : '?'}页正文` } finally { state.inflight-- }
        },
      },
    }
  }
  /** 独立 apply() 上下文：tools.get('modlens_read_image') 返回假桥（各自的 captured Map）。 */
  function ctxWithVision(vision) {
    const cap = new Map()
    apply({
      get: n => (n !== 'tools' ? undefined : {
        register: d => { cap.set(d.name, d); return () => { } },
        get: nm => (nm === 'modlens_read_image' ? vision : undefined),
      }),
    })
    return cap
  }
  const vr = (cap, args) => cap.get('office_read').execute(args, exec)

  ok('P0-2 并发度解析：只有 2/3 生效，其余（含未设）一律回落 1', (() => {
    const set = v => { if (v === undefined) delete process.env.DSH_OFFICE_VISION_CONCURRENCY; else process.env.DSH_OFFICE_VISION_CONCURRENCY = v; return visionConcurrency() }
    const got = [set(undefined), set('1'), set('2'), set('3'), set('4'), set('0'), set('abc'), set('')]
    delete process.env.DSH_OFFICE_VISION_CONCURRENCY
    return got.join(',') === '1,1,2,3,1,1,1,1'
  })())
  ok('P0-2 预算解析：默认无上限，显式 0 合法，垃圾值回落无上限', (() => {
    const set = v => { if (v === undefined) delete process.env.DSH_OFFICE_VISION_MAX_CALLS; else process.env.DSH_OFFICE_VISION_MAX_CALLS = v; return visionBudget() }
    const d = set(undefined); const z = set('0'); const t = set('2'); const bad = set('x')
    delete process.env.DSH_OFFICE_VISION_MAX_CALLS
    return d.max === Infinity && d.enabled === false && z.max === 0 && z.enabled === true && t.max === 2 && bad.max === Infinity
  })())

  // 6 页无文本层 PDF（hr + pagebreak）。渲染 6 页 = 6 次 PowerShell 单页冷启，
  // 先预热一次把渲染缓存填满，后面计时的几读就只剩视觉往返。
  const vPdf = join(OUT, 'vision-pages.pdf')
  await call('office_create', {
    path: vPdf,
    document: { blocks: Array.from({ length: 11 }, (_, i) => ({ type: i % 2 ? 'pagebreak' : 'hr' })) },
  })
  const vMeta = await call('office_read', { path: vPdf, as: 'meta' })
  ok('P0-2 夹具为 6 页无文本层 PDF', vMeta.stats.pages === 6 && vMeta.stats.scannedPages.length === 6, `pages=${vMeta.stats.pages}`)

  // ---- 第八轮（自 WB 侧 P1-2 回移）：未指定 pages 时的 OCR 预览必须记账 ----
  // 走 sidecar 缓存命中路径（把 1-3 页的 OCR 文本直接铺好），绝不冷跑引擎；
  // 读完立刻删掉，别污染紧接着的视觉桥段（那段要求 6 页都真跑）。
  {
    const pvSide = writeOcrSidecar(vPdf, `covered: 1-3 | total: 6 | parser: ${PDF_PARSER_VERSION} | src: rapidocr=1-3`,
      ['# OCR 文本（由 dsh-office 生成，可安全删除以重新识别）',
        '', ...[1, 2, 3].flatMap(n => [`## 第 ${n} 页（OCR）`, `Fixture preview page ${n}`, ''])].join('\n'))
    const pv = await call('office_read', { path: vPdf, ocr: 'always', ocrEngine: 'local' })
    rmSync(pvSide, { force: true })
    ok('P1 预览记账：stats.ocrPreview = {preview:3, total:6, skipped:"4-6"}',
      pv.stats.ocrPreview?.preview === 3 && pv.stats.ocrPreview?.total === 6 && pv.stats.ocrPreview?.skipped === '4-6',
      JSON.stringify(pv.stats.ocrPreview))
    ok('P1 预览记账：正文脚注明说"未指定 pages 按预览只 OCR 前 3 页" + 待识别总数',
      /未指定 pages 按预览只 OCR 前 3 页/.test(String(pv.content))
      && /待识别共 6 页，第 4-6 页未做，仍用原文本层/.test(String(pv.content)),
      (String(pv.content).match(/未指定 pages 按预览[^；\n]*/) || ['<无>'])[0].slice(0, 150))
    ok('P1 预览记账：脚注给可复制的续读命令（ocr="always" ocrEngine="local" pages="4-6"）',
      /office_read path="[^"]+" ocr="always" ocrEngine="local" pages="4-6"/.test(String(pv.content)),
      (String(pv.content).match(/续读：[^；\n]*/) || ['<无>'])[0].slice(0, 150))
  }

  // ---- 第八轮（自 WB 侧 P1-2 回移）：OCR_MAX_PAGES 砍页时的**文案通道** ----
  // 历史判断（R9-3 之后的行为已修正，勿再当作现状）：那一段的 21-30 页保留乱码文本层，整篇会被质量门判
  // "不可读"而走 sidecar 降级，降级 notice/content **曾**覆盖这条脚注（记账只剩 stats）——那是旧行为；
  // R9-3 起降级路径会从 stats 重建同一批脚注并追加进 notice/content。夹具选择不变：这里仍换一个
  // **有文本层**的 26 页文档验文案通道（1-20 页手写 sidecar 缓存、零渲染零引擎，21-26 页正常落回文本层，脚注因此可见）；对应修订说明见 test.mjs:1834。
  {
    const capPdf = join(OUT, 'preview-capped.pdf')
    await call('office_create', {
      path: capPdf,
      document: {
        blocks: Array.from({ length: 51 }, (_, i) => (i % 2
          ? { type: 'pagebreak' }
          : { type: 'paragraph', runs: [{ text: `Fixture cap page ${(i >> 1) + 1}` }] })),
      },
    })
    const capMeta = await call('office_read', { path: capPdf, as: 'meta' })
    ok('页数上限文案：夹具为 26 页有文本层 PDF（不触发降级，脚注才看得见）',
      capMeta.stats.pages === 26 && capMeta.stats.scannedPages.length === 0, `pages=${capMeta.stats.pages}`)
    const capSide = writeOcrSidecar(capPdf, `covered: 1-20 | total: 26 | parser: ${PDF_PARSER_VERSION} | src: rapidocr=1-20`,
      ['# OCR 文本（由 dsh-office 生成，可安全删除以重新识别）',
        '', ...Array.from({ length: 20 }, (_, i) => i + 1).flatMap(n => [`## 第 ${n} 页（OCR）`, `Fixture cap page ${n}`, ''])].join('\n'))
    const capRead = await call('office_read', { path: capPdf, pages: '1-26', ocr: 'always', ocrEngine: 'local' })
    rmSync(capSide, { force: true })
    ok('页数上限文案：stats.ocrPagesCapped = {limit:20, requested:26, applied:20, skipped:[21..26]}',
      capRead.stats.ocrPagesCapped?.limit === 20 && capRead.stats.ocrPagesCapped?.requested === 26
      && capRead.stats.ocrPagesCapped?.applied === 20
      && capRead.stats.ocrPagesCapped?.skipped.join(',') === '21,22,23,24,25,26',
      JSON.stringify(capRead.stats.ocrPagesCapped))
    ok('页数上限文案：正文脚注明说"单次读取最多识别 20 页"并给出续读命令（pages="21-26"）',
      /单次读取最多识别 20 页：本次要求 26 页，第 21-26 页未做 OCR（仍用原文本层）/.test(String(capRead.content))
      && /office_read path="[^"]+" ocr="always" ocrEngine="local" pages="21-26"/.test(String(capRead.content)),
      (String(capRead.content).match(/单次读取最多识别[^；\n]*/) || ['<无>'])[0].slice(0, 170))
  }

  const vSide = join(OUT, 'vision-pages.ocr.md')
  await vr(ctxWithVision(fakeVision().bridge), { path: vPdf, pages: '1-6', ocr: 'always', ocrEngine: 'vision' })
  rmSync(vSide, { force: true })

  // a) 默认并发（未设）→ 最大在飞 ≤ 1，行为与旧版逐页串行一致
  const vA = fakeVision()
  const tSerial = Date.now()
  const readA = await vr(ctxWithVision(vA.bridge), { path: vPdf, pages: '1-6', ocr: 'always', ocrEngine: 'vision' })
  const serialMs = Date.now() - tSerial
  ok('P0-2 默认并发 1：假视觉桥最大在飞 ≤1', vA.state.maxInflight === 1, `maxInflight=${vA.state.maxInflight}`)
  ok('P0-2 6 页各出正文、visionCalls=6', vA.state.calls === 6 && readA.stats.ocrVisionCalls === 6
    && (readA.content.match(/第\d+页正文/g) || []).length === 6, `calls=${vA.state.calls} visionCalls=${readA.stats.ocrVisionCalls}`)
  ok('P0-2 未设预算时不外显 ocrVisionSkipped（既有 stats 结构不变）', readA.stats.ocrVisionSkipped === undefined)
  rmSync(vSide, { force: true })

  // b) 并发 2 → 最大在飞 =2，墙钟明显短于串行，页级结果一字不差
  process.env.DSH_OFFICE_VISION_CONCURRENCY = '2'
  const vB = fakeVision()
  const tPar = Date.now()
  const readB = await vr(ctxWithVision(vB.bridge), { path: vPdf, pages: '1-6', ocr: 'always', ocrEngine: 'vision' })
  const parMs = Date.now() - tPar
  delete process.env.DSH_OFFICE_VISION_CONCURRENCY
  rmSync(vSide, { force: true })
  const bodies = s => (String(s).match(/第\d+页正文/g) || []).sort().join(',')
  ok('P0-2 并发 2：最大在飞 =2', vB.state.maxInflight === 2, `maxInflight=${vB.state.maxInflight}`)
  ok('P0-2 并发 2 墙钟明显短于串行', serialMs - parMs >= 200, `serial=${serialMs}ms parallel=${parMs}ms`)
  ok('P0-2 并发不改变页级结果（页码排序后逐字一致）', bodies(readB.content) === bodies(readA.content),
    `${bodies(readB.content)} vs ${bodies(readA.content)}`)

  // c) 预算护栏：MAX_CALLS=2 覆盖 6 页 → 只花 2 次，3-6 页记 skipped 并写进 stats + 脚注
  process.env.DSH_OFFICE_VISION_MAX_CALLS = '2'
  const vC = fakeVision()
  const readC = await vr(ctxWithVision(vC.bridge), { path: vPdf, pages: '1-6', ocr: 'always', ocrEngine: 'vision' })
  delete process.env.DSH_OFFICE_VISION_MAX_CALLS
  ok('P0-2 预算上限 2：真实只调 2 次且 visionCalls=2', vC.state.calls === 2 && readC.stats.ocrVisionCalls === 2,
    `calls=${vC.state.calls} visionCalls=${readC.stats.ocrVisionCalls}`)
  ok('P0-2 超预算页记 skipped 并外显为 stats.ocrVisionSkipped="3-6"', readC.stats.ocrVisionSkipped === '3-6',
    String(readC.stats.ocrVisionSkipped))
  ok('P0-2 脚注写明计划/完成/跳过（绝不静默截断）',
    /第 3-6 页因视觉调用预算上限跳过（计划 6 \/ 完成 2 \/ 跳过 4）/.test(readC.content),
    (readC.content.match(/^> .*预算.*$/m) || ['<无脚注>'])[0].slice(0, 130))
  ok('P0-2 被跳过的页仍逐页给出失败原因', readC.stats.ocrFailed.length === 4
    && /预算已用尽（DSH_OFFICE_VISION_MAX_CALLS=2）/.test(readC.content), JSON.stringify(readC.stats.ocrFailed))
  // DSH 补充（任务三-3 / F3）：批内缺页必须"页码 + 一句话原因"成对给出，且正文脚注自带 sidecar 现状
  ok('批内缺页：stats.ocrFailedPages 与 ocrFailed 同页、每页带原因', readC.stats.ocrFailedPages?.length === 4
    && readC.stats.ocrFailedPages.every(r => Number.isInteger(r.page) && typeof r.reason === 'string' && r.reason.length > 3)
    && JSON.stringify(readC.stats.ocrFailedPages.map(r => r.page)) === JSON.stringify(readC.stats.ocrFailed),
    JSON.stringify(readC.stats.ocrFailedPages?.[0] || {}))
  ok('批内缺页：正文脚注直接列出缺哪几页 + sidecar 路径 + covered（不必再 grep）',
    /批内缺页：/.test(readC.content) && /sidecar=/.test(readC.content) && /covered:/.test(readC.content)
    && /grep 'covered:'/.test(readC.content), (readC.content.match(/批内缺页：[^\n]*/) || ['<无>'])[0].slice(0, 150))
  // 预算读到的页已进 sidecar：二次命中不再消耗任何视觉调用
  const vD = fakeVision()
  const readD = await vr(ctxWithVision(vD.bridge), { path: vPdf, pages: '1-2', ocr: 'always', ocrEngine: 'vision' })
  ok('P0-2 已识别的页进缓存，二次命中 0 调用', vD.state.calls === 0 && readD.stats.ocrVisionCalls === undefined
    && /OCR 缓存/.test(readD.content), `calls=${vD.state.calls}`)
  rmSync(vSide, { force: true })

  // ---------------- P1-4 PDF 全量解析 memo（LRU ≤3） ----------------
  const memoPdf = join(OUT, 'memo.pdf')
  await call('office_create', { path: memoPdf, markdown: '# 记忆化\n\n正文一行\n' })
  const ms0 = pdfMemoStats()
  const memo1 = await call('office_read', { path: memoPdf, as: 'meta' })
  const ms1 = pdfMemoStats()
  const memo2 = await call('office_read', { path: memoPdf, as: 'meta' })
  const ms2 = pdfMemoStats()
  ok('P1-4 首次读 PDF 走解析（misses+1、无命中）', ms1.misses === ms0.misses + 1 && ms1.hits === ms0.hits,
    `misses ${ms0.misses}→${ms1.misses} / hits ${ms0.hits}→${ms1.hits}`)
  ok('P1-4 二次读命中 memo（hits+1、不再解析）', ms2.hits === ms1.hits + 1 && ms2.misses === ms1.misses,
    `misses ${ms1.misses}→${ms2.misses} / hits ${ms1.hits}→${ms2.hits}`)
  ok('P1-4 memo 命中不改变返回内容与 stats', memo1.content === memo2.content
    && memo1.stats.pages === memo2.stats.pages, `pages=${memo2.stats.pages}`)
  utimesSync(memoPdf, new Date(Date.now() - 120000), new Date(Date.now() - 120000))
  const ms3 = pdfMemoStats()
  await call('office_read', { path: memoPdf, as: 'meta' })
  const ms4 = pdfMemoStats()
  ok('P1-4 mtime 一变 key 失效（不会拿旧解析结果）', ms4.misses === ms3.misses + 1,
    `misses ${ms3.misses}→${ms4.misses}`)
  ok('P1-4 LRU 上限 3', pdfMemoStats().size <= 3, `size=${pdfMemoStats().size}`)

  // ---------------- P2-5 渲染缓存清理 ----------------
  const sweepPdf = join(OUT, 'dsh-sweep-probe.pdf')
  await call('office_create', { path: sweepPdf, markdown: '# sweep\n' })
  const ocrRoot = join(tmpdir(), 'dsh-office-ocr')
  mkdirSync(ocrRoot, { recursive: true })
  const probeBase = basename(sweepPdf)
  const curRenderDir = renderDirFor(sweepPdf)   // 任务一后目录名含路径指纹，这里必须与实现同源
  const mkAged = (name, ageDays) => {
    const d = join(ocrRoot, name)
    rmSync(d, { recursive: true, force: true })
    mkdirSync(d, { recursive: true })
    const t = new Date(Date.now() - ageDays * 86400e3)
    utimesSync(d, t, t)
    return d
  }
  const sweepOld = mkAged(`${probeBase}-111`, 8)
  const sweepFresh = mkAged(`${probeBase}-222`, 1)
  const rapidOld = mkAged('rapid-sweeptest-old', 2)
  const rapidFresh = mkAged('rapid-sweeptest-fresh', 0.02)
  const rapidWithCache = mkAged('rapid-sweeptest-withcache', 30)
  writeFileSync(join(rapidWithCache, 'x.ocr.json'), '{}')
  const protFile = join(ocrRoot, 'sweeptest.ocr.md')
  writeFileSync(protFile, '<!-- dsh-office OCR cache -->\n')
  mkdirSync(curRenderDir, { recursive: true })

  const sweepRep = sweepRenderCache(sweepPdf)
  ok('P2-5 同 basename 的旧 mtime 目录被清掉（>7 天）',
    !existsSync(sweepOld) && sweepRep.removed.some(p => p === sweepOld), `removed=${sweepRep.removed.length}`)
  ok('P2-5 未到 7 天的同 basename 目录保留', existsSync(sweepFresh))
  ok('P2-5 当前正在用的渲染目录绝不碰', existsSync(curRenderDir))
  ok('P2-5 rapid-* 超 1 天清理、新的保留', !existsSync(rapidOld) && existsSync(rapidFresh))
  ok('P2-5 .ocr.md/.ocr.json 与其所在目录永不碰',
    existsSync(protFile) && existsSync(rapidWithCache) && existsSync(join(rapidWithCache, 'x.ocr.json')))
  ok('P2-5 文件不存在也不抛（now=0 → 什么都不判为陈旧）', (() => {
    try {
      const r = sweepRenderCache(join(OUT, 'no-such-probe.pdf'), { now: 0 })
      return typeof r.scanned === 'number' && Array.isArray(r.removed) && r.removed.length === 0
    } catch { return false }
  })())
  // 不许在系统 temp 里留测试垃圾
  for (const d of [sweepOld, sweepFresh, rapidOld, rapidFresh, rapidWithCache, curRenderDir]) rmSync(d, { recursive: true, force: true })
  rmSync(protFile, { force: true })

  // ---------------- P2-6 >20 页本地分批护栏 ----------------
  ok('P2-6 分批纯函数：25→20+5、40→20+20、41→20+20+1、0→[]、3→[3]', (() => {
    const seq = n => Array.from({ length: n }, (_, i) => i + 1)
    const shape = a => a.map(b => b.length).join('+')
    const a = chunkLocalBatches(seq(25)); const b = chunkLocalBatches(seq(40))
    const c = chunkLocalBatches(seq(0)); const d = chunkLocalBatches(seq(3)); const e = chunkLocalBatches(seq(41))
    return shape(a) === '20+5' && a[1][0] === 21 && a[1][4] === 25
      && shape(b) === '20+20' && b.length === 2
      && c.length === 0 && shape(d) === '3' && shape(e) === '20+20+1'
  })())
  ok('P2-6 默认批大小 = LOCAL_MAX_IMAGES(20)，非法 size 回落默认', (() => {
    const seq = n => Array.from({ length: n }, (_, i) => i + 1)
    return LOCAL_MAX_IMAGES === 20
      && chunkLocalBatches(seq(25), 0)[0].length === 20
      && chunkLocalBatches(seq(25), -3)[0].length === 20
      && chunkLocalBatches(seq(25), 10).length === 3
      && chunkLocalBatches(seq(25))[0].length === LOCAL_MAX_IMAGES
  })())
  ok('P2-6 单批内行为不变（2 页仍是一批走完）', chunkLocalBatches([1, 2]).length === 1)

  const odt = join(OUT, 'notes.odt')
  r = await call('office_create', { path: odt, markdown: MD })
  ok('创建 odt', existsSync(odt) && r.bytes > 1000)
  txt = await textOf(odt)
  ok('读取 odt', txt.includes('季度经营报告') && txt.includes('华北'))

  const ods = join(OUT, 'sheet.ods')
  r = await call('office_create', { path: ods, workbook: { sheets: [{ name: '库存', rows: [['物料', '数量'], ['螺丝', 120], ['垫片', 340]] }] } })
  ok('创建 ods', existsSync(ods))
  txt = await textOf(ods)
  ok('读取 ods 数值', txt.includes('螺丝') && txt.includes('120'))

  const odp = join(OUT, 'slides.odp')
  r = await call('office_create', { path: odp, slides: { slides: [{ layout: 'title', title: '周会汇报', subtitle: '2026-W12' }, { layout: 'content', title: '进度', bullets: [{ text: '完成 A 模块', level: 0 }, { text: '联调 B 模块', level: 0 }] }] } })
  ok('创建 odp', existsSync(odp))
  txt = await textOf(odp)
  ok('读取 odp', txt.includes('周会汇报') && txt.includes('完成 A 模块'))

  const csv = join(OUT, 'table.csv')
  r = await call('office_create', { path: csv, workbook: { sheets: [{ name: 'S1', rows: [['城市', '人口'], ['唐山', 770], ['保定', 1150]] }] } })
  ok('创建 csv', existsSync(csv))
  txt = await textOf(csv)
  ok('读取 csv', txt.includes('唐山') && txt.includes('1150'))

  const mdOut = join(OUT, 'note.md')
  r = await call('office_create', { path: mdOut, markdown: '# 便签\n\n- 一\n- 二\n' })
  ok('创建 md', existsSync(mdOut))
  txt = await textOf(mdOut)
  ok('读取 md', txt.includes('便签') && txt.includes('二'))

  const txtOut = join(OUT, 'plain.txt')
  r = await call('office_create', { path: txtOut, document: { blocks: [{ type: 'paragraph', runs: [{ text: '第一行' }] }, { type: 'heading', level: 1, text: '标题' }] } })
  ok('创建 txt', existsSync(txtOut))
  txt = await textOf(txtOut)
  ok('读取 txt', txt.includes('第一行') && txt.includes('标题'))

  // ---------------- editing ----------------
  r = await call('office_edit', {
    path: docx,
    operations: [
      { op: 'replace_text', find: '91%', replace: '93%' },
      { op: 'append_markdown', markdown: '## 补充说明\n\n新增内容已写入。' },
      { op: 'set_meta', title: '季度报告(修订版)', author: 'DSH' },
    ],
  })
  ok('docx 编辑返回摘要', r.summary.length === 3, r.summary.join(' | '))
  txt = await textOf(docx)
  ok('docx replace_text 生效', txt.includes('93%') && !txt.includes('91%'))
  ok('docx append_markdown 生效', txt.includes('补充说明') && txt.includes('新增内容已写入'))
  const docxMeta = await call('office_read', { path: docx, as: 'meta' })
  ok('docx set_meta 生效', docxMeta.meta.title === '季度报告(修订版)' && docxMeta.meta.author === 'DSH', JSON.stringify(docxMeta.meta))

  r = await call('office_edit', {
    path: xlsx,
    operations: [
      { op: 'set_cell', sheet: '成绩', cell: 'B5', value: 100 },
      { op: 'append_rows', sheet: '成绩', rows: [['赵六', 66, '物理']] },
      { op: 'replace_value', find: '李四', replace: '李四(改)' },
      { op: 'add_sheet', name: '新增表', rows: [['k', 'v'], ['a', '1']] },
      { op: 'rename_sheet', name: '汇总', newName: '统计' },
    ],
  })
  ok('xlsx 编辑 5 步', r.summary.length === 5, r.summary.join(' | '))
  txt = await textOf(xlsx)
  ok('xlsx set_cell 生效', txt.includes('100'))
  ok('xlsx append_rows 生效', txt.includes('赵六') && txt.includes('物理'))
  ok('xlsx replace_value 生效', txt.includes('李四(改)'))
  ok('xlsx add_sheet/rename_sheet 生效', txt.includes('新增表') && txt.includes('统计') && !txt.includes('汇总'))
  r = await call('office_edit', { path: xlsx, operations: [{ op: 'delete_sheet', name: '新增表' }] })
  txt = await textOf(xlsx)
  ok('xlsx delete_sheet 生效', !txt.includes('新增表'), r.summary[0])

  r = await call('office_edit', {
    path: pptx,
    operations: [
      { op: 'replace_text', find: '团队版', replace: '企业版' },
      { op: 'add_slide', slide: { title: '答疑环节', bullets: ['Q&A', '联系方式'] } },
      { op: 'update_slide', index: 2, title: '本次亮点(修订)', bullets: ['零依赖', '全格式'] },
    ],
  })
  ok('pptx 编辑 3 步', r.summary.length === 3, r.summary.join(' | '))
  txt = await textOf(pptx)
  ok('pptx replace_text 生效', txt.includes('企业版') && !txt.includes('团队版'))
  ok('pptx add_slide 生效', txt.includes('答疑环节') && txt.includes('Q&A'))
  ok('pptx update_slide 生效', txt.includes('本次亮点(修订)') && txt.includes('零依赖'))
  ok('pptx 删除前共 4 页', (await call('office_read', { path: pptx, as: 'meta' })).stats.slides === 4)
  await call('office_edit', { path: pptx, operations: [{ op: 'delete_slide', index: 4 }] })
  ok('pptx delete_slide 生效', (await call('office_read', { path: pptx, as: 'meta' })).stats.slides === 3)

  await call('office_edit', {
    path: odt,
    operations: [
      { op: 'replace_text', find: '季度经营报告', replace: '年度经营报告' },
      { op: 'append_markdown', markdown: '追加段：ODF 编辑测试。' },
    ],
  })
  txt = await textOf(odt)
  ok('odt 编辑生效', txt.includes('年度经营报告') && txt.includes('ODF 编辑测试'))

  await call('office_edit', { path: ods, operations: [{ op: 'replace_text', find: '螺丝', replace: '螺栓' }] })
  txt = await textOf(ods)
  ok('ods replace_text 生效', txt.includes('螺栓') && !txt.includes('螺丝'))

  await call('office_edit', { path: odp, operations: [{ op: 'replace_text', find: '周会汇报', replace: '月度汇报' }] })
  txt = await textOf(odp)
  ok('odp replace_text 生效', txt.includes('月度汇报'))

  await call('office_edit', { path: csv, operations: [{ op: 'replace_value', find: '770', replace: '771' }, { op: 'append_rows', rows: [['石家庄', 1120]] }] })
  txt = await textOf(csv)
  ok('csv 编辑生效', txt.includes('771') && txt.includes('石家庄'))

  await call('office_edit', { path: mdOut, operations: [{ op: 'append_text', text: '- 三' }, { op: 'replace_text', find: '便签', replace: '记事' }] })
  txt = await textOf(mdOut)
  ok('md 编辑生效', txt.includes('记事') && txt.includes('三'))

  // ---------------- conversions ----------------
  const conv = async (src, dst, opts = {}) => {
    const res = await call('office_convert', { source: src, target: dst, ...opts })
    return res
  }
  let c = await conv(xlsx, join(OUT, 'data.csv'))
  ok('xlsx → csv', existsSync(c.target) && c.bytes > 50)
  const csvText = await readFile(join(OUT, 'data.csv'), 'utf8')
  ok('csv 转换内容正确', csvText.includes('张三') && csvText.includes('赵六'))

  c = await conv(join(OUT, 'data.csv'), join(OUT, 'from-csv.xlsx'))
  txt = await textOf(join(OUT, 'from-csv.xlsx'))
  ok('csv → xlsx', txt.includes('张三'))

  c = await conv(docx, join(OUT, 'report-from-docx.pdf'))
  ok('docx → pdf', existsSync(c.target) && c.bytes > 2000)
  ok('docx → pdf 内容可回读', (await textOf(join(OUT, 'report-from-docx.pdf'))).includes('季度经营报告'))

  c = await conv(pptx, join(OUT, 'deck.md'))
  txt = await textOf(join(OUT, 'deck.md'))
  ok('pptx → md', txt.includes('答疑环节') || txt.includes('本次亮点'))

  c = await conv(docx, join(OUT, 'report.odt'))
  ok('docx → odt', (await textOf(join(OUT, 'report.odt'))).includes('季度经营报告'))

  c = await conv(odt, join(OUT, 'notes.docx'))
  ok('odt → docx', (await textOf(join(OUT, 'notes.docx'))).includes('年度经营报告'))

  c = await conv(xlsx, join(OUT, 'data.pptx'))
  ok('xlsx → pptx', (await textOf(join(OUT, 'data.pptx'))).includes('张三'))

  c = await conv(pptx, join(OUT, 'deck.xlsx'))
  ok('pptx → xlsx', (await textOf(join(OUT, 'deck.xlsx'))).length > 20)

  c = await conv(join(OUT, 'note.md'), join(OUT, 'from-md.pdf'))
  ok('md → pdf', existsSync(c.target))

  c = await conv(docx, join(OUT, 'report.txt'))
  ok('docx → txt', (await textOf(join(OUT, 'report.txt'))).includes('季度经营报告'))

  // ---------------- PNG 分带裁剪（应对视觉模型 1024 token 输出上限） ----------------
  {
    const { deflateSync } = await import('node:zlib')
    const { readPngInfo, cropPngBand } = await import('./png.js')
    const t = (() => { const a = new Int32Array(256); for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; a[i] = c } return a })()
    const crc = b => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = t[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0 }
    const chunk = (type, data) => {
      const out = Buffer.alloc(12 + data.length)
      out.writeUInt32BE(data.length, 0)
      out.write(type, 4, 'latin1')
      data.copy(out, 8)
      out.writeUInt32BE(crc(Buffer.concat([Buffer.from(type, 'latin1'), data])), 8 + data.length)
      return out
    }
    const makePng = (w, h) => {
      const raw = Buffer.alloc((w * 4 + 1) * h)
      for (let y = 0; y < h; y++) {
        const off = y * (w * 4 + 1)
        raw[off] = 0
        for (let x = 0; x < w; x++) { raw[off + 1 + x * 4] = x & 0xff; raw[off + 2 + x * 4] = y & 0xff; raw[off + 3 + x * 4] = 255; raw[off + 4 + x * 4] = 0 }
      }
      const ihdr = Buffer.alloc(13)
      ihdr.writeUInt32BE(w, 0)
      ihdr.writeUInt32BE(h, 4)
      ihdr[8] = 8
      ihdr[9] = 6
      return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0)),
      ])
    }
    try {
      const src = makePng(60, 400)
      const info = readPngInfo(src)
      const i1 = readPngInfo(cropPngBand(src, 0, 2))
      const i2 = readPngInfo(cropPngBand(src, 1, 2))
      const i4 = readPngInfo(cropPngBand(src, 3, 4))
      ok('PNG 分带裁剪：宽度不变、每段≈1/N 页高', i1.width === 60 && i2.width === 60
        && Math.abs(i1.height - 200) <= 30 && Math.abs(i2.height - 200) <= 30 && Math.abs(i4.height - 100) <= 30,
      `full=${info.width}x${info.height} 2band=${i1.height},${i2.height} 4band=${i4.height}`)
      ok('PNG 裁剪结果仍是合法 PNG', i1.colorType === 6 && i1.height > 0 && cropPngBand(src, 0, 1).length === src.length, `colorType=${i1.colorType}`)
    } catch (e) {
      ok('PNG 分带裁剪：宽度不变、每段≈1/N 页高', false, e.message)
    }
  }

  // ---------------- 权限加密 PDF（空打开密码）+ 扫描件 ----------------
  const encPdf = 'samples/sample-C.pdf'
  if (existsSync(encPdf)) {
    try {
      const em = await call('office_read', { path: encPdf, as: 'meta' })
      ok('权限加密 PDF 可透明解密', em.stats.pages === 11 && em.stats.scannedPages.length === 11, `pages=${em.stats.pages} scanned=${em.stats.scannedPages.length}`)
      // P0 端到端：只识别 1-2 页时必须报出剩余页（拷进产物目录，不碰用户文件旁的 sidecar）
      const encCopy = join(OUT, 'enc-scan.pdf')
      await copyFile(encPdf, encCopy)
      const part = await call('office_read', { path: encCopy, pages: '1-2', ocr: 'always', ocrEngine: 'local' })
      const encSide = readFileSync(join(OUT, 'enc-scan.ocr.md'), 'utf8')
      ok('P0 部分覆盖时明确报出未识别页', /未识别/.test(part.content) && /全书 11 页/.test(part.content),
        (part.content.match(/^> .*$/m) || ['<无脚注>'])[0].slice(0, 160))
      ok('P0 sidecar manifest 写出 covered/total', /covered: [\d,\-]+ \| total: 11/.test(encSide), encSide.split('\n')[1])
      ok('P0 stats.ocrUncovered 给出剩余页', /^[\d,\-]+$/.test(String(part.stats.ocrUncovered)), String(part.stats.ocrUncovered))
    } catch (e) {
      ok('权限加密 PDF 可透明解密', false, e.message)
    }
  } else {
    ok('权限加密 PDF 可透明解密', true, '样本不存在，跳过')
  }

  // ---------------- real-world files ----------------
  const real = [
    ['samples/sample-shortlist.csv', '.csv'],
    ['samples/sample-positions.xlsx', '.xlsx'],
    ['samples/sample-deck.pptx', '.pptx'],
    ['samples/sample-lecture.pdf', '.pdf'],
    ['samples/sample-lecture.docx', '.docx'],
  ]
  for (const [file, ext] of real) {
    // 环境依赖样本：文件缺失 = 本机没有该文件，不是代码问题 —— 按套件既有惯例记
    // skip（与"权限加密 PDF…样本不存在，跳过"同口径），不再把环境缺失误报成 FAIL。
    if (!existsSync(file)) { ok(`真实文件 ${ext} 存在`, true, `缺失，跳过: ${file}`); continue }
    try {
      const res = await call('office_read', { path: file, as: 'markdown', limit: 5000 })
      const size = (res.content || '').length
      ok(`真实 ${ext} 读取`, size > 40, `${size} 字符, format=${res.format}`)
    } catch (e) {
      ok(`真实 ${ext} 读取`, false, e.message)
    }
  }

  // real xlsx metadata + sheet listing
  const realXlsx = 'samples/sample-roster.xlsx'
  if (existsSync(realXlsx)) {
    const meta = await call('office_read', { path: realXlsx, as: 'meta' })
    ok('真实 xlsx 工作表枚举', Array.isArray(meta.stats.sheets) && meta.stats.sheets.length >= 1, JSON.stringify(meta.stats.sheets))
  }

  // ---------------- independent zip/CRC validation (own parser, not zip.js) ----------------
  const CRC_T = (() => { const t = new Int32Array(256); for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[i] = c } return t })()
  const crc32 = b => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0 }
  function auditZip(file) {
    const buf = readFileSync(file)
    let eocd = -1
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 65536); i--) {
      if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
    }
    if (eocd < 0) throw new Error('no EOCD')
    const count = buf.readUInt16LE(eocd + 10)
    const cdSize = buf.readUInt32LE(eocd + 12)
    const cdOff = buf.readUInt32LE(eocd + 16)
    if (cdOff + cdSize > buf.length) throw new Error('central directory out of range')
    let p = cdOff
    const names = []
    let bad = 0
    for (let n = 0; n < count; n++) {
      if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('bad CD entry')
      const method = buf.readUInt16LE(p + 10)
      const crc = buf.readUInt32LE(p + 16)
      const csize = buf.readUInt32LE(p + 20)
      const usize = buf.readUInt32LE(p + 24)
      const nameLen = buf.readUInt16LE(p + 28)
      const extraLen = buf.readUInt16LE(p + 30)
      const commentLen = buf.readUInt16LE(p + 32)
      const localOff = buf.readUInt32LE(p + 42)
      const name = buf.subarray(p + 46, p + 46 + nameLen).toString('utf8')
      names.push(name)
      if (buf.readUInt32LE(localOff) !== 0x04034b50) throw new Error(`bad local header for ${name}`)
      const lNameLen = buf.readUInt16LE(localOff + 26)
      const lExtraLen = buf.readUInt16LE(localOff + 28)
      const dataStart = localOff + 30 + lNameLen + lExtraLen
      const raw = buf.subarray(dataStart, dataStart + csize)
      const data = method === 8 ? inflateRawSync(raw) : raw
      if (data.length !== usize) bad++
      if (crc32(data) !== crc) bad++
      p += 46 + nameLen + extraLen + commentLen
    }
    return { count, names, bad }
  }
  for (const f of ['report.docx', 'data.xlsx', 'deck.pptx', 'notes.odt', 'sheet.ods', 'slides.odp']) {
    try {
      const a = auditZip(join(OUT, f))
      const required = f.endsWith('.docx') ? '[Content_Types].xml'
        : f.endsWith('.xlsx') ? 'xl/workbook.xml'
          : f.endsWith('.pptx') ? 'ppt/presentation.xml'
            : 'content.xml'
      ok(`zip 完整性与 CRC (${f})`, a.bad === 0 && a.names.includes(required), `${a.count} 条目, CRC 错误 ${a.bad}, 含 ${required}`)
    } catch (e) {
      ok(`zip 完整性与 CRC (${f})`, false, e.message)
    }
  }
  ok('ODF mimetype 存储方式合法', (() => {
    const buf = readFileSync(join(OUT, 'notes.odt'))
    const eocd = buf.length - 22
    const cdOff = buf.readUInt32LE(eocd + 16)
    const method = buf.readUInt16LE(cdOff + 10)
    const name = buf.subarray(cdOff + 46, cdOff + 46 + 8).toString('utf8')
    return method === 0 && name === 'mimetype'
  })(), 'mimetype 必须首个且不压缩')

  // ===========================================================================
  // 任务三（R16）：zip 解压资源限制与结构校验
  // 前缀：R16-zip：
  // ===========================================================================
  {
    /** 逐字段自造 zip（makeZip 造不出越界 / 说谎的结构）。 */
    const mkZipBytes = (o = {}) => {
      const { name = 'a.txt', method = 0, body = Buffer.from('hello dsh-office'), compSize, uncSize,
        localOff = 0, extra = Buffer.alloc(0), cdOffsetDelta = 0, cdSizeOverride, eocdCount,
        eocdComment = Buffer.alloc(0), patchCd } = o
      const stored = method === 0 ? body : deflateRawSync(body, { level: 9 })
      const nameBytes = Buffer.from(name, 'utf8')
      const lh = Buffer.alloc(30)
      lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x800, 6)
      lh.writeUInt16LE(method, 8); lh.writeUInt32LE(crc32(body), 14)
      lh.writeUInt32LE(compSize ?? stored.length, 18); lh.writeUInt32LE(uncSize ?? body.length, 22)
      lh.writeUInt16LE(nameBytes.length, 26); lh.writeUInt16LE(0, 28)
      const cdOff = 30 + nameBytes.length + stored.length
      const cen = Buffer.alloc(46)
      cen.writeUInt32LE(0x02014b50, 0); cen.writeUInt16LE(20, 4); cen.writeUInt16LE(20, 6)
      cen.writeUInt16LE(0x800, 8); cen.writeUInt16LE(method, 10); cen.writeUInt32LE(crc32(body), 16)
      cen.writeUInt32LE(compSize ?? stored.length, 20); cen.writeUInt32LE(uncSize ?? body.length, 24)
      cen.writeUInt16LE(nameBytes.length, 28); cen.writeUInt16LE(extra.length, 30)
      cen.writeUInt32LE(localOff, 42)
      const eocd = Buffer.alloc(22)
      eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(eocdCount ?? 1, 10)
      eocd.writeUInt32LE(cdSizeOverride ?? (46 + nameBytes.length + extra.length), 12)
      eocd.writeUInt32LE(cdOff + cdOffsetDelta, 16)
      eocd.writeUInt16LE(eocdComment.length, 20)
      const buf = Buffer.concat([lh, nameBytes, stored, cen, nameBytes, extra, eocd, eocdComment])
      if (patchCd) patchCd(buf, cdOff)
      return buf
    }
    /** 读一次；RangeError 一律算失败（宿主级异常不算"可控错误"）。 */
    const tryZip = (buf, name = 'a.txt', opts) => {
      try { return { ok: true, data: Buffer.from(openZip(buf, opts).get(name) ?? []) } }
      catch (e) { return { ok: false, msg: String(e && e.message), range: e instanceof RangeError } }
    }

    // ---- A 防回归：正常 Office 包与既有独立校验器逐项一致 ----
    for (const f of ['report.docx', 'data.xlsx', 'deck.pptx', 'notes.odt']) {
      const z = openZip(readFileSync(join(OUT, f)))
      const a = auditZip(join(OUT, f))
      ok(`R16-zip：${f} 条目清单与独立校验器一致`,
        JSON.stringify(z.names) === JSON.stringify(a.names), `${z.names.length} 条目`)
      ok(`R16-zip：${f} 每个条目都能完整读出（新读取器 0 误拒）`,
        z.names.every(n => Buffer.from(z.get(n)).length >= 0), z.names.slice(0, 3).join(', ') + '…')
    }

    // ---- B 合法大条目 / 高压缩比不被误杀 ----
    const bigZip = mkZipBytes({ name: 'big.bin', method: 8, body: Buffer.alloc(4 * 1024 * 1024) })
    const big = tryZip(bigZip, 'big.bin')
    ok('R16-zip：4 MiB 高压缩比条目默认仍可正常解压（不按压缩比误杀）',
      big.ok && big.data.length === 4 * 1024 * 1024, big.ok ? `${big.data.length} 字节` : big.msg.slice(0, 90))
    ok('R16-zip：默认单条目/累计阈值足以容纳合法大文档（打印实际默认值）',
      maxEntryBytes() >= 256 * 1024 * 1024 && maxTotalBytes() >= 256 * 1024 * 1024,
      `entry=${maxEntryBytes() / 1048576}MiB total=${maxTotalBytes() / 1048576}MiB`)
    const capped = tryZip(bigZip, 'big.bin', { maxEntryBytes: 1024 * 1024 })
    ok('R16-zip：单条目上限命中 → 解压前即被拒，文案点名环境变量',
      !capped.ok && !capped.range && /zip 条目解压后过大/.test(capped.msg)
      && /DSH_OFFICE_ZIP_MAX_ENTRY_BYTES/.test(capped.msg), capped.msg.slice(0, 130))
    const uncapped = tryZip(bigZip, 'big.bin', { maxEntryBytes: 0 })
    ok('R16-zip：上限 0 = 关闭该上限', uncapped.ok && uncapped.data.length === 4 * 1024 * 1024,
      uncapped.ok ? `${uncapped.data.length}` : uncapped.msg.slice(0, 80))

    const manyZip = makeZip(Array.from({ length: 8 }, (_, i) => ({ name: `b${i}.bin`, data: Buffer.alloc(1024 * 1024, 65) })))
    const zMany = openZip(manyZip, { maxTotalBytes: 3 * 1024 * 1024 })
    let cumMsg = ''
    try { for (const n of zMany.names) zMany.get(n) } catch (e) { cumMsg = String(e.message) }
    ok('R16-zip：累计解压超限 → 读到第 N 条时被拒，文案点名 DSH_OFFICE_ZIP_MAX_TOTAL_BYTES',
      /zip 累计解压体积超限/.test(cumMsg) && /DSH_OFFICE_ZIP_MAX_TOTAL_BYTES/.test(cumMsg), cumMsg.slice(0, 130))

    // ---- C 声明尺寸不符 / 炸弹 ----
    const lie = tryZip(mkZipBytes({ method: 8, body: Buffer.from('x'.repeat(2000)), uncSize: 1 }))
    ok('R16-zip：uncSize 说谎（声明 1、实际 2000）→ 拒绝，绝不返回部分数据',
      !lie.ok && !lie.range && /zip 炸弹防护/.test(lie.msg), lie.msg.slice(0, 120))
    const lieLen = tryZip(mkZipBytes({ method: 8, body: Buffer.from('x'.repeat(2000)), uncSize: 4000 }))
    ok('R16-zip：声明长度大于实际解压量 → 报"声明 N 字节，实际 M 字节"',
      !lieLen.ok && !lieLen.range && /声明解压后/.test(lieLen.msg), lieLen.msg.slice(0, 120))

    // ---- D 越界 / 损坏：一律可读错误，绝不是 RangeError ----
    const corruptCases = [
      ['cdOffset 越界', mkZipBytes({ cdOffsetDelta: 99000 })],
      ['localOff 越界', mkZipBytes({ localOff: 0x7fffffff })],
      ['compSize 越界', mkZipBytes({ compSize: 4294967280 })],
      ['CD nameLen 说谎', mkZipBytes({ patchCd: (b, o) => b.writeUInt16LE(60000, o + 28) })],
      ['CD 签名破坏', mkZipBytes({ patchCd: (b, o) => b.writeUInt32LE(0xdeadbeef, o) })],
      ['CD 条目数虚报', mkZipBytes({ eocdCount: 0xffff })],
      ['文件被截断', mkZipBytes({}).subarray(0, 5)],
    ]
    for (const [label, buf] of corruptCases) {
      const r = tryZip(buf)
      ok(`R16-zip：${label} → 可读的 zip 结构错误（不是 RangeError）`,
        !r.ok && !r.range && /zip |未找到中央目录/.test(r.msg), r.msg.slice(0, 110))
    }
    const badCrc = mkZipBytes({ patchCd: (b, o) => b.writeUInt32LE(0xdeadbeef, o + 16) })
    const cr = tryZip(badCrc)
    ok('R16-zip：CD 里的 CRC 不符 → 明确拒绝并点名 DSH_OFFICE_ZIP_CRC',
      !cr.ok && /CRC32 不匹配/.test(cr.msg) && /DSH_OFFICE_ZIP_CRC=0/.test(cr.msg), cr.msg.slice(0, 130))
    const crOff = tryZip(badCrc, 'a.txt', { checkCrc: false })
    ok('R16-zip：checkCrc=false（= DSH_OFFICE_ZIP_CRC=0）→ 放行', crOff.ok, crOff.msg || '')

    // ---- E ZIP64 本地头偏移（旧版第 90 行赋值丢失的 bug）----
    const mkExtra = size => {
      const e = Buffer.alloc(4 + size)
      e.writeUInt16LE(0x0001, 0); e.writeUInt16LE(size, 2)
      if (size >= 8) e.writeBigUInt64LE(0n, 4)   // 真实 localOff = 0
      return e
    }
    const z64 = tryZip(mkZipBytes({ localOff: 0xffffffff, extra: mkExtra(8) }))
    ok('R16-zip：ZIP64 本地头偏移的条目 → 现在能正常读出（旧版必 RangeError）',
      z64.ok && z64.data.toString('utf8') === 'hello dsh-office', z64.ok ? z64.data.toString('utf8') : z64.msg.slice(0, 110))
    const z64bad = tryZip(mkZipBytes({ localOff: 0xffffffff, extra: mkExtra(0) }))
    ok('R16-zip：有 ZIP64 哨兵但扩展字段不足 → 明确报错（不静默读垃圾）',
      !z64bad.ok && !z64bad.range && /ZIP64 扩展字段不足/.test(z64bad.msg), z64bad.msg.slice(0, 120))

    // ---- F 注释里的伪 EOCD 不影响定位 ----
    const withComment = (() => {
      const g = mkZipBytes({})
      const comment = Buffer.concat([Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(20)])
      const b = Buffer.concat([g, comment])
      b.writeUInt16LE(comment.length, b.length - comment.length - 2)
      return b
    })()
    const wc = tryZip(withComment)
    ok('R16-zip：注释里含伪 EOCD 签名 → 仍读到真条目（先严格后宽容两轮定位）',
      wc.ok && wc.data.toString('utf8') === 'hello dsh-office', wc.ok ? 'ok' : wc.msg.slice(0, 120))

    // ---- G 写侧（R18 任务 B）：真实 ZIP64 + 不需要 ZIP64 的包逐字节布局不变 ----
    /** 独立走一遍 zip 结构（只用于**写侧布局**断言，故意不复用 openZip 的宽容分支）。 */
    const walkZip = (buf, cdRelOverride) => {
      const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
      const eocdAt = buf.length - 22
      const locals = []
      let p = 0
      while (p + 4 <= buf.length && dv.getUint32(p, true) === 0x04034b50) {
        const nameLen = dv.getUint16(p + 26, true), extraLen = dv.getUint16(p + 28, true)
        const compSize = dv.getUint32(p + 18, true)
        locals.push({ at: p, name: buf.subarray(p + 30, p + 30 + nameLen).toString('utf8'), extraLen, compSize, version: dv.getUint16(p + 4, true), size: 30 + nameLen + extraLen + compSize })
        p += 30 + nameLen + extraLen + compSize
      }
      const cdAt = cdRelOverride ?? dv.getUint32(eocdAt + 16, true)
      const cdRecords = []
      let q = cdAt
      while (q + 4 <= buf.length && dv.getUint32(q, true) === 0x02014b50) {
        const nameLen = dv.getUint16(q + 28, true), extraLen = dv.getUint16(q + 30, true)
        cdRecords.push({
          at: q, name: buf.subarray(q + 46, q + 46 + nameLen).toString('utf8'), extraLen,
          extraRel: extraLen ? q + 46 + nameLen : -1,
          extra: extraLen ? buf.subarray(q + 46 + nameLen, q + 46 + nameLen + extraLen) : null,
          localOff: dv.getUint32(q + 42, true), version: dv.getUint16(q + 6, true),
          size: 46 + nameLen + extraLen + dv.getUint16(q + 32, true),
        })
        q += 46 + nameLen + extraLen + dv.getUint16(q + 32, true)
      }
      const locSig = dv.getUint32(eocdAt - 20, true)
      return {
        dv, eocdAt, locals, cdRecords, hasLocator: locSig === 0x07064b50,
        z64At: locSig === 0x07064b50 ? Number(dv.getBigUint64(eocdAt - 12, true)) : -1,
        cdCount: dv.getUint16(eocdAt + 10, true), cdSize: dv.getUint32(eocdAt + 12, true), cdOffset: dv.getUint32(eocdAt + 16, true),
      }
    }

    for (const [label, entries] of [
      ['纯 store 小包', [{ name: 'a.txt', data: 'hello', store: true }, { name: 'b/c.bin', data: 'world' }]],
      ['ODF 形态（mimetype 首个不压缩）', [{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true }, { name: 'content.xml', data: '<x/>' }]],
      ['大文本（走 deflate）', [{ name: 'big.txt', data: 'A'.repeat(200000) }, { name: 'cn.txt', data: '中文'.repeat(5000) }]],
    ]) {
      const buf = Buffer.from(makeZip(entries))
      const w = walkZip(buf)
      // 旧版布局 = Σ(30+name+extra+compSize) + Σ(46+name+extra) + 22，且**没有任何 ZIP64 结构**
      const expect = w.locals.reduce((s, l) => s + l.size, 0) + w.cdRecords.reduce((s, c) => s + c.size, 0) + 22
      ok(`R18-zip：${label} 仍是旧版字节布局（无 ZIP64 结构、长度公式精确、version needed=20）`,
        w.locals.length === entries.length && w.cdRecords.length === entries.length
        && w.locals.every(l => l.extraLen === 0 && l.version === 20)
        && w.cdRecords.every(c => c.extraLen === 0 && c.version === 20)
        && !w.hasLocator && buf.length === expect && w.cdOffset === expect - 22 - w.cdSize,
        `bytes=${buf.length}/${expect} locator=${w.hasLocator}`)
    }

    /** 用 .NET 的 ZipArchive（`Expand-Archive` 内部就是它）打开 zip：受限沙箱里管道 stdio 会被拒，
     *  所以 spawnSync 一律 `stdio:'ignore'`，结果经文件回读。 */
    const dotNetZipProbe = zipPath => {
      const log = join(OUT, `r18-dotnet-${Date.now()}-${Math.floor(Math.random() * 1e6)}.txt`)
      const ps = 'Add-Type -AssemblyName System.IO.Compression.FileSystem;'
        + `$z=[System.IO.Compression.ZipFile]::OpenRead('${zipPath}');`
        + '$e=$z.Entries[0];$sr=New-Object System.IO.StreamReader($e.Open());$t=$sr.ReadToEnd();'
        + '$r="COUNT=" + $z.Entries.Count + ";FIRST=" + $t;'
        + '$sr.Dispose();$z.Dispose();'
        + `Set-Content -LiteralPath '${log}' -Value $r -Encoding UTF8`
      const res = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps],
        { windowsHide: true, timeout: 300000, stdio: 'ignore' })
      const txt = (existsSync(log) ? readFileSync(log, 'utf8') : '').replace(/^\uFEFF/, '').replace(/\r?\n/g, ' ')
      try { rmSync(log, { force: true }) } catch { /* ignore */ }
      return { status: res.status, txt }
    }

    // ---- G2 >65535 条目：真实 ZIP64 EOCD + 定位记录；本仓 openZip 与 .NET 都要认得 ----
    {
      const many = Array.from({ length: 65536 }, (_, i) => ({ name: `e${i}.bin`, data: '', store: true }))
      const manyBytes = Buffer.from(makeZip(many))
      const w = walkZip(manyBytes)
      const z64Count = w.z64At >= 0 ? Number(w.dv.getBigUint64(w.z64At + 32, true)) : -1
      ok('R18-zip：65536 条目 → 写 ZIP64（EOCD 计数写哨兵 0xffff、定位记录 + EOCD64 计数为真值 44）',
        w.cdCount === 0xffff && w.hasLocator && w.z64At > 0
        && w.dv.getUint32(w.z64At, true) === 0x06064b50 && Number(w.dv.getBigUint64(w.z64At + 4, true)) === 44
        && z64Count === 65536,
        `count=0x${w.cdCount.toString(16)} locator=${w.hasLocator} z64At=${w.z64At} z64count=${z64Count}`)
      const z = openZip(manyBytes)
      ok('R18-zip：65536 条目 → 本仓 openZip 全部回读（旧版会静默丢 1 条以上）',
        z.names.length === 65536 && z.names[0] === 'e0.bin' && z.names[65535] === 'e65535.bin',
        `${z.names.length} 条目 first=${z.names[0]} last=${z.names[z.names.length - 1]}`)
      const z64Path = join(OUT, 'r18-zip64-65536.zip')
      writeFileSync(z64Path, manyBytes)
      const dn = dotNetZipProbe(z64Path)
      ok('R18-zip：.NET ZipArchive（Expand-Archive 的引擎）能打开这个 ZIP64 包并列出 65536 条',
        dn.status === 0 && /COUNT=65536/.test(dn.txt), `status=${dn.status} ${dn.txt.slice(0, 90)}`)
      // 真·Expand-Archive 逐文件解包 65536 条目在本机实测 >10 分钟（PowerShell 每文件开销极大）
      // → 默认不跑，设 DSH_OFFICE_TEST_EXPAND_ARCHIVE=1 才做（会解出 65536 个空文件）。
      if (process.env.DSH_OFFICE_TEST_EXPAND_ARCHIVE === '1') {
        const dest = join(OUT, 'r18-expand-65536')
        rmSync(dest, { recursive: true, force: true })
        const res = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
          `Expand-Archive -LiteralPath '${z64Path}' -DestinationPath '${dest}' -Force`],
        { windowsHide: true, timeout: 3600000, stdio: 'ignore' })
        ok('R18-zip：Expand-Archive 真解包 ZIP64 包（opt-in）→ 65536 个文件',
          res.status === 0 && existsSync(dest) && readdirSync(dest).length === 65536,
          `status=${res.status} files=${existsSync(dest) ? readdirSync(dest).length : 0}`)
      } else {
        ok('R18-zip：Expand-Archive 真解包 ZIP64 包 → 跳过（opt-in：本机 65536 文件解包 >10 分钟；'
          + '.NET ZipArchive 读取已自动验证）', true, '设 DSH_OFFICE_TEST_EXPAND_ARCHIVE=1 打开')
      }
    }

    // ---- G3 >4 GiB 本地头偏移（baseOffset 接缝）：哨兵 + 中央目录 ZIP64 扩展 + EOCD64 ----
    {
      const baseOffset = 0xfffffffe
      const entries = [{ name: 'a', data: 'x'.repeat(10), store: true }, { name: 'b', data: 'y', store: true }]
      const cdRel = 30 + 1 + 10 + 30 + 1 + 1
      const buf = Buffer.from(makeZip(entries, { baseOffset }))
      const w = walkZip(buf, cdRel)
      const second = w.cdRecords[1]
      const ex = second?.extra
      ok('R18-zip：>4 GiB 本地头偏移 → 中央目录字段写哨兵 + ZIP64 扩展带真值（APPNOTE 字段序）',
        !!second && second.localOff === 0xffffffff && second.extraLen === 12 && ex
        && w.dv.getUint16(second.extraRel, true) === 0x0001 && w.dv.getUint16(second.extraRel + 2, true) === 8
        && Number(w.dv.getBigUint64(second.extraRel + 4, true)) === baseOffset + 30 + 1 + 10,
        `localOff=0x${(second?.localOff ?? 0).toString(16)} extra=${ex ? ex.toString('hex') : '(none)'}`)
      // 定位记录里存的是**虚拟**偏移（baseOffset + 真实位置），所以不能拿它当缓冲区下标；
      // EOCD64 的真实位置 = EOCD 之前 56+20 字节处。
      const z64Real = w.eocdAt - 76
      ok('R18-zip：>4 GiB 偏移 → EOCD 偏移写哨兵，定位记录 + EOCD64 带真值',
        w.cdOffset === 0xffffffff && w.hasLocator
        && w.dv.getUint32(z64Real, true) === 0x06064b50
        && Number(w.dv.getBigUint64(z64Real + 48, true)) === baseOffset + cdRel
        && Number(w.dv.getBigUint64(w.eocdAt - 12, true)) === baseOffset + cdRel + w.cdSize,
        `eocdOff=0x${w.cdOffset.toString(16)} z64cd=${Number(w.dv.getBigUint64(z64Real + 48, true))}`
        + ` locatorPtr=${Number(w.dv.getBigUint64(w.eocdAt - 12, true))} cdSize=${w.cdSize}`)
      ok('R18-zip：不传 baseOffset 的同一组条目仍能正常回读（接缝不影响默认行为）',
        openZip(makeZip(entries)).names.join(',') === 'a,b')
    }

    // ---- G4 条目名 > 65535 字节：显式错误（ZIP64 没有名字长度扩展）----
    let longNameErr = ''
    try { makeZip([{ name: 'n'.repeat(70000), data: '' }]) } catch (e) { longNameErr = String(e.message) }
    ok('R18-zip：条目名 > 65535 字节 → 显式错误，并说清"ZIP64 只扩尺寸/偏移、没有名字长度扩展"',
      /条目名过长/.test(longNameErr) && /65535/.test(longNameErr) && /ZIP64/.test(longNameErr) && /APPNOTE 4\.5\.3/.test(longNameErr),
      longNameErr.slice(0, 150))

    // ---- G5 zip64ExtraField 记录布局 ----
    {
      const ex = zip64ExtraField(['uncSize', 'localOffset'], { uncSize: 4294967300n, compSize: 1n, localOffset: 4294967296n })
      const edv = new DataView(ex.buffer, ex.byteOffset, ex.byteLength)
      ok('R18-zip：zip64ExtraField = ID 0x0001 + 长度 + 按 APPNOTE 顺序的 8 字节小端值',
        ex.length === 20 && edv.getUint16(0, true) === 0x0001 && edv.getUint16(2, true) === 16
        && edv.getBigUint64(4, true) === 4294967300n && edv.getBigUint64(12, true) === 4294967296n,
        ex.toString('hex'))
    }

    // ---- G6 CRC 误拒面（R18 任务 B.2）：**保持严格拒绝**，但点名条目 + 明确逃生口 ----
    {
      const good = Buffer.from(makeZip([{ name: 'payload.bin', data: 'AAAA' }, { name: 'other.txt', data: 'ok' }]))
      const g = walkZip(good)
      // 样本①：数据流"被等长替换"（内容 BBBB 替掉 AAAA），中央目录 CRC 与流不符
      const swapped = Buffer.from(good)
      swapped.writeUInt32LE(crc32(Buffer.from('BBBB')), g.cdRecords[0].at + 16)
      const rSwap = tryZip(swapped, 'payload.bin')
      ok('R18-zip：CD CRC 与数据流不符（数据流被等长替换）→ 严格拒绝 + 点名条目 + 逃生口',
        !rSwap.ok && /CRC32 不匹配/.test(rSwap.msg) && /"payload.bin"/.test(rSwap.msg) && /DSH_OFFICE_ZIP_CRC=0/.test(rSwap.msg),
        rSwap.msg.slice(0, 130))
      process.env.DSH_OFFICE_ZIP_CRC = '0'
      const rEnv = tryZip(swapped, 'payload.bin')
      delete process.env.DSH_OFFICE_ZIP_CRC
      ok('R18-zip：DSH_OFFICE_ZIP_CRC=0 逃生口对"等长替换"样本同样生效（无需改调用代码）',
        rEnv.ok && rEnv.data.toString('utf8') === 'AAAA', rEnv.msg || '')
      // 样本②：CD 里 CRC = 0（旧式流式写包器不填）→ 默认放行，不算误拒
      const zero = Buffer.from(good)
      zero.writeUInt32LE(0, g.cdRecords[0].at + 16)
      ok('R18-zip：CD CRC=0（旧式流式写包）→ 默认放行，不算误拒', tryZip(zero, 'payload.bin').ok)
    }

    // ---- R19-A：ZIP64「声明 ≥4 GiB 数据/压缩尺寸」的**外部实现行为边界**（本轮把风险钉成硬证据）----
    // R18 交付时 >4 GiB **数据尺寸**分支只有布局级证据：`makeZip()` 要真的收 ≥4 GiB 的 `entry.data`
    // （`crc32()` 遍历它 + `concatBytes()` 再复制一份 ⇒ 峰值 ≈8.6 GB），本机（15.9 GB 总内存、
    // 测试时空闲 3.6 GB）跑不了端到端。这里补上**能补的那一半**：手工拼"声明 ≥4 GiB、实际数据 5 字节"
    // 的 ZIP64 包（extra 直接调 `zip64ExtraField()`，字段序与 `makeZip` 的 unc64/comp64 分支一致），
    // 把"我们的 ZIP64 尺寸声明会不会被主流实现接受"钉死；真实 >4 GiB 合法条目另做 opt-in（③）。
    {
      const BIG = 4294967396            // = 0x100000064，越过 0xffffffff 哨兵线
      /** 按 `makeZip()` 的 ZIP64 分支规则手工拼包；**声明值**可以与实际数据不一致。 */
      const declZip64 = ({ name, body, declaredUnc, declaredComp, sentinelUnc = true, sentinelComp = false, method = 0, crc: crcOverride }) => {
        const nameB = Buffer.from(name, 'utf8')
        const crc = crcOverride ?? crc32(body)
        const lhKeys = (sentinelUnc || sentinelComp) ? ['uncSize', 'compSize'] : []
        const lhExtra = lhKeys.length ? zip64ExtraField(lhKeys, { uncSize: declaredUnc, compSize: declaredComp }) : Buffer.alloc(0)
        const lh = Buffer.alloc(30)
        lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(45, 4); lh.writeUInt16LE(0x800, 6)
        lh.writeUInt16LE(method, 8); lh.writeUInt32LE(crc, 14)
        lh.writeUInt32LE(sentinelComp ? 0xffffffff : declaredComp, 18)
        lh.writeUInt32LE(sentinelUnc ? 0xffffffff : declaredUnc, 22)
        lh.writeUInt16LE(nameB.length, 26); lh.writeUInt16LE(lhExtra.length, 28)
        const localPart = Buffer.concat([lh, nameB, lhExtra, body])
        const cdKeys = []
        if (sentinelUnc) cdKeys.push('uncSize')
        if (sentinelComp) cdKeys.push('compSize')
        const cdExtra = cdKeys.length ? zip64ExtraField(cdKeys, { uncSize: declaredUnc, compSize: declaredComp }) : Buffer.alloc(0)
        const cen = Buffer.alloc(46)
        cen.writeUInt32LE(0x02014b50, 0); cen.writeUInt16LE(45, 4); cen.writeUInt16LE(45, 6)
        cen.writeUInt16LE(0x800, 8); cen.writeUInt16LE(method, 10); cen.writeUInt32LE(crc, 16)
        cen.writeUInt32LE(sentinelComp ? 0xffffffff : declaredComp, 20)
        cen.writeUInt32LE(sentinelUnc ? 0xffffffff : declaredUnc, 24)
        cen.writeUInt16LE(nameB.length, 28); cen.writeUInt16LE(cdExtra.length, 30)
        cen.writeUInt32LE(0, 42)
        const centralPart = Buffer.concat([cen, nameB, cdExtra])
        const cdOffset = localPart.length, cdSize = centralPart.length
        const z = Buffer.alloc(56)
        z.writeUInt32LE(0x06064b50, 0); z.writeBigUInt64LE(44n, 4); z.writeUInt16LE(45, 12); z.writeUInt16LE(45, 14)
        z.writeBigUInt64LE(1n, 24); z.writeBigUInt64LE(1n, 32)
        z.writeBigUInt64LE(BigInt(cdSize), 40); z.writeBigUInt64LE(BigInt(cdOffset), 48)
        const loc = Buffer.alloc(20)
        loc.writeUInt32LE(0x07064b50, 0); loc.writeBigUInt64LE(BigInt(cdOffset + cdSize), 8); loc.writeUInt32LE(1, 16)
        const eocd = Buffer.alloc(22)
        eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(0xffff, 8); eocd.writeUInt16LE(0xffff, 10)
        eocd.writeUInt32LE(0xffffffff, 12); eocd.writeUInt32LE(0xffffffff, 16)
        return Buffer.concat([localPart, centralPart, z, loc, eocd])
      }
      /** `.NET ZipArchive`（`Expand-Archive` 的引擎）：打开 → 列 Length/CompressedLength → 试着读 64 KiB。 */
      const dotNetZipInfo = zipPath => {
        const log = join(OUT, `r19-dotnet-info-${Date.now()}-${Math.floor(Math.random() * 1e6)}.txt`)
        const ps = 'Add-Type -AssemblyName System.IO.Compression.FileSystem;'
          + `$z=[System.IO.Compression.ZipFile]::OpenRead('${zipPath}');`
          + '$e=$z.Entries[0];'
          + '$r="OPEN=ok COUNT=" + $z.Entries.Count + " LENGTH=" + $e.Length + " COMPRESSED=" + $e.CompressedLength;'
          + 'try { $s=$e.Open(); $b=New-Object byte[] 65536; $n=$s.Read($b,0,65536); $r += " READ=ok read=" + $n; $s.Dispose() }'
          + 'catch { $r += " READ=err " + $_.Exception.GetType().Name + " :: " + $_.Exception.Message };'
          + '$z.Dispose();'
          + `Set-Content -LiteralPath '${log}' -Value $r -Encoding UTF8`
        const res = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps],
          { windowsHide: true, timeout: 120000, stdio: 'ignore' })
        const txt = (existsSync(log) ? readFileSync(log, 'utf8') : '').replace(/^\uFEFF/, '').replace(/\r?\n/g, ' ').trim()
        try { rmSync(log, { force: true }) } catch { /* ignore */ }
        return { status: res.status, txt }
      }

      // ① 只有"未压缩尺寸"越线（等价于 deflate 高压缩比 + >4 GiB 原始数据）
      const uncOnly = declZip64({ name: 'big.txt', body: Buffer.from('hello'), declaredUnc: BIG, declaredComp: 5 })
      const r1 = tryZip(uncOnly, 'big.txt')
      ok('R19-zip：CD 声明未压缩尺寸 ≥4 GiB（数据只有 5 字节）→ 本仓按**声明**拦在单条目上限（绝不先解压）',
        !r1.ok && !r1.range && /zip 条目解压后过大/.test(r1.msg) && /4096/.test(r1.msg), r1.msg.slice(0, 130))
      const p1 = join(OUT, 'r19-declared-unc-4gib.zip')
      writeFileSync(p1, uncOnly)
      const dn1 = dotNetZipInfo(p1)
      ok('R19-zip：同一包 .NET ZipArchive **接受该声明**（Entry.Length 读成 4 GiB+）⇒ 我们的 ZIP64 尺寸声明被主流实现认可',
        dn1.status === 0 && /OPEN=ok/.test(dn1.txt) && /LENGTH=4294967396/.test(dn1.txt), dn1.txt.slice(0, 160))
      ok('R19-zip：.NET 对该包**不做"未压缩尺寸一致性"校验**（按 compSize 只读 5 字节就返回）—— 本仓比它更严',
        /READ=ok read=5/.test(dn1.txt), dn1.txt.slice(0, 160))

      // ② 未压缩 + 压缩尺寸**同时**越线（store 4 GiB 条目的真实形态）
      const both = declZip64({ name: 'big.bin', body: Buffer.from('hello'), declaredUnc: BIG, declaredComp: BIG, sentinelComp: true })
      const r2 = tryZip(both, 'big.bin')
      ok('R19-zip：CD 同时声明压缩/未压缩 ≥4 GiB → 本仓报"声明压缩长度超过文件长度"（结构自洽性校验）',
        !r2.ok && !r2.range && /声明压缩长度/.test(r2.msg) && /超过文件长度/.test(r2.msg), r2.msg.slice(0, 130))
      const p2 = join(OUT, 'r19-declared-both-4gib.zip')
      writeFileSync(p2, both)
      const dn2 = dotNetZipInfo(p2)
      ok('R19-zip：同一包 .NET 能打开、Length/CompressedLength 都报 4 GiB+，但真读流时报"本地文件头已损坏"',
        dn2.status === 0 && /LENGTH=4294967396/.test(dn2.txt) && /COMPRESSED=4294967396/.test(dn2.txt)
        && /READ=err/.test(dn2.txt), dn2.txt.slice(0, 170))

      // ③ opt-in：**真实**的 >4 GiB 合法条目（流式 deflate，内存 O(1)）—— 声明与真实内容都被接受
      if (process.env.DSH_OFFICE_TEST_REAL_4GIB === '1') {
        const N = 0x100000000 + 4096
        const chunk = Buffer.alloc(1 << 20, 0x41)
        let written = 0, crcAcc = 0
        const parts = []
        const d = createDeflateRaw({ level: 9 })
        d.on('data', c => parts.push(c))
        const ended = new Promise((res, rej) => { d.on('end', res); d.on('error', rej) })
        const t0 = Date.now()
        while (written < N) {
          const n = Math.min(chunk.length, N - written)
          const part = n === chunk.length ? chunk : chunk.subarray(0, n)
          crcAcc = zlibCrc32(part, crcAcc) >>> 0
          if (!d.write(part)) await once(d, 'drain')
          written += n
        }
        d.end()
        await ended
        const deflated = Buffer.concat(parts)
        const realMs = Date.now() - t0
        const realZip = declZip64({ name: 'big.bin', body: deflated, method: 8, declaredUnc: N, declaredComp: deflated.length, crc: crcAcc })
        const p3 = join(OUT, 'r19-real-4gib.zip')
        writeFileSync(p3, realZip)
        const r3 = tryZip(realZip, 'big.bin')
        ok('R19-zip：真实 >4 GiB 条目（流式 deflate）→ 本仓仍按声明拦在 256 MiB 上限（不先解压）',
          !r3.ok && !r3.range && /zip 条目解压后过大/.test(r3.msg), r3.msg.slice(0, 120))
        const log3 = join(OUT, 'r19-dotnet-real4gib.txt')
        const ps3 = 'Add-Type -AssemblyName System.IO.Compression.FileSystem;'
          + `$z=[System.IO.Compression.ZipFile]::OpenRead('${p3}'); $e=$z.Entries[0];`
          + '$s=$e.Open(); $b=New-Object byte[] 1048576; $n=$s.Read($b,0,$b.Length);'
          + '$bad=0; for($i=0;$i -lt $n;$i++){ if($b[$i] -ne 65){$bad++} };'
          + `Set-Content -LiteralPath '${log3}' -Value ("LENGTH=" + $e.Length + " COMPRESSED=" + $e.CompressedLength + " READ=" + $n + " NONA=" + $bad) -Encoding UTF8;`
          + '$s.Dispose(); $z.Dispose()'
        const res3 = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps3],
          { windowsHide: true, timeout: 300000, stdio: 'ignore' })
        const txt3 = (existsSync(log3) ? readFileSync(log3, 'utf8') : '').replace(/^\uFEFF/, '').replace(/\r?\n/g, ' ').trim()
        ok('R19-zip：同一真实包 .NET 能读到 Length=4 GiB+ 并**正确解出**前 1 MiB（全为 A）',
          res3.status === 0 && txt3.includes(`LENGTH=${N}`) && /READ=1048576 NONA=0/.test(txt3),
          `${txt3 || '(no log)'}（deflate ${realMs}ms → ${(deflated.length / 1048576).toFixed(2)}MiB）`)
      } else {
        ok('R19-zip：真实 >4 GiB 合法条目端到端（opt-in）→ 跳过', true,
          '设 DSH_OFFICE_TEST_REAL_4GIB=1 打开（本机实测：4 GiB deflate 13.0 s → 4.17 MB；.NET 正确解出前 1 MiB）')
      }
    }

    // ---- H 端到端：炸弹经 office_read 被上限拦下且错误可读、带下一步 ----
    process.env.DSH_OFFICE_ZIP_MAX_ENTRY_BYTES = '1048576'
    try {
      const bombDocx = join(OUT, 'r16-bomb.docx')
      writeFileSync(bombDocx, makeZip([
        { name: '[Content_Types].xml', data: '<?xml version="1.0"?><Types/>' },
        { name: 'word/document.xml', data: `<?xml version="1.0"?><w:document>${'A'.repeat(8 * 1024 * 1024)}</w:document>` },
      ]))
      let bombMsg = ''
      try { await call('office_read', { path: bombDocx, as: 'meta' }) } catch (e) { bombMsg = String(e.message) }
      ok('R16-zip：端到端炸弹（docx 形态）→ 可读错误 + 四要素下一步（不是宿主级异常）',
        /zip 条目解压后过大/.test(bombMsg) && /DSH_OFFICE_ZIP_MAX_ENTRY_BYTES/.test(bombMsg)
        && /下一步=/.test(bombMsg), bombMsg.slice(0, 150))
    } finally { delete process.env.DSH_OFFICE_ZIP_MAX_ENTRY_BYTES }
  }

  // ===========================================================================
  // 任务二（R16）：创建 / 转换 / 编辑 / 缓存的落盘一律"临时件 → 成功后替换"
  // 前缀：R16-原子写：
  // ===========================================================================
  {
    const target = join(OUT, 'r16-atomic.txt')
    writeFileSync(target, '原始内容-必须完整保留', 'utf8')
    const before = readFileSync(target, 'utf8')
    const strayTmp = () => readdirSync(OUT).filter(n => n.startsWith('.dsh-tmp-'))

    // ① 临时写入失败 → 原目标不动
    process.env.DSH_OFFICE_ATOMIC_FAULT = 'temp-write'
    let m1 = ''
    try { await writeFileAtomic(target, Buffer.from('新内容')) } catch (e) { m1 = String(e.message) }
    delete process.env.DSH_OFFICE_ATOMIC_FAULT
    ok('R16-原子写：临时写入失败 → 原目标逐字节不变 + 可读四要素错误',
      readFileSync(target, 'utf8') === before && /【写盘失败｜四要素】/.test(m1), m1.slice(0, 100))
    ok('R16-原子写：临时写入失败 → 不留临时文件', strayTmp().length === 0, strayTmp().join(', '))

    // ② 最终替换失败 → 原目标不动 + 清理
    process.env.DSH_OFFICE_ATOMIC_FAULT = 'publish'
    let m2 = ''
    try { await writeFileAtomic(target, Buffer.from('新内容')) } catch (e) { m2 = String(e.message) }
    delete process.env.DSH_OFFICE_ATOMIC_FAULT
    ok('R16-原子写：替换失败 → 原目标不变 + 说明"已保留原文件" + 清理临时文件',
      readFileSync(target, 'utf8') === before && /已保留原文件/.test(m2) && strayTmp().length === 0, m2.slice(0, 100))

    // ③ 成功路径
    await writeFileAtomic(target, Buffer.from('新内容'))
    ok('R16-原子写：成功路径内容正确且无残留',
      readFileSync(target, 'utf8') === '新内容' && strayTmp().length === 0)

    // ④ 并发：临时名由 O_EXCL('wx') 保证唯一 → 16 路互不覆盖，终值恰为某个完整值
    const conc = join(OUT, 'r16-atomic-conc.txt')
    const vals = Array.from({ length: 16 }, (_, i) => `并发内容-${i}-${'x'.repeat(50)}`)
    let concErr = ''
    try { await Promise.all(vals.map(v => writeFileAtomic(conc, Buffer.from(v)))) } catch (e) { concErr = String(e.message) }
    const fin = existsSync(conc) ? readFileSync(conc, 'utf8') : ''
    ok('R16-原子写：16 路并发写同一目标 → 无错、终值恰为某一个完整值（无交错/半截）',
      !concErr && vals.includes(fin) && strayTmp().length === 0, concErr || fin.slice(0, 50))

    // ⑤ 工具级：office_edit 发布失败不得改坏原件（旧版是 writeFile 直接截断）
    const editTarget = join(OUT, 'r16-atomic-edit.docx')
    await call('office_create', { path: editTarget, markdown: '# 原子写\n\n原件正文\n' })
    const editBefore = readFileSync(editTarget)
    process.env.DSH_OFFICE_ATOMIC_FAULT = 'publish'
    let em = ''
    try { await call('office_edit', { path: editTarget, operations: [{ op: 'replace_text', find: '原件正文', replace: '改后正文' }] }) } catch (e) { em = String(e.message) }
    delete process.env.DSH_OFFICE_ATOMIC_FAULT
    ok('R16-原子写：office_edit 发布失败 → 原件逐字节不变（不再"改坏原件"）',
      Buffer.compare(readFileSync(editTarget), editBefore) === 0 && /写盘失败/.test(em), em.slice(0, 110))
    const reread = await textOf(editTarget)
    ok('R16-原子写：失败后原件仍可正常读取且内容未变',
      reread.includes('原件正文') && !reread.includes('改后正文'), reread.replace(/\n/g, ' ').slice(0, 70))

    // ⑥ office_convert 的 source === target：源先完整读入，再安全替换
    const selfConv = join(OUT, 'r16-atomic-self.md')
    writeFileSync(selfConv, '# 自转\n\n正文甲\n', 'utf8')
    await call('office_convert', { source: selfConv, target: selfConv, encoding: 'utf-8' })
    ok('R16-原子写：convert source===target（同路径自转）→ 源先读完再替换，内容正确',
      readFileSync(selfConv, 'utf8').includes('正文甲'), readFileSync(selfConv, 'utf8').slice(0, 40))

    // ⑦ sidecar（同步版）：注入失败不得留半截缓存
    const junkFor = join(OUT, 'r16-atomic-junk.md')
    writeFileSync(junkFor, `# 乱码\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(400)}`, 'utf8')
    process.env.DSH_OFFICE_ATOMIC_FAULT = 'temp-write'
    let rb = null
    try { rb = await call('office_read', { path: junkFor, as: 'markdown' }) } catch (e) { rb = { error: String(e.message) } }
    delete process.env.DSH_OFFICE_ATOMIC_FAULT
    const sideP = readSidecarPath(junkFor)
    ok('R16-原子写：sidecar 写入失败 → 不留半截 .read.md，且返回里说明失败',
      !existsSync(sideP) && String(rb?.stats?.sidecarNote || rb?.error || '').length > 0,
      JSON.stringify(rb?.stats?.sidecarNote || rb?.error || '').slice(0, 90))
    await call('office_read', { path: junkFor, as: 'markdown' })
    ok('R16-原子写：解除故障后 sidecar 完整写出（含身份行）',
      existsSync(sideP) && readFileSync(sideP, 'utf8').includes('srcsha256'))

    // ---- ⑧ R18 任务 C：**真实文件系统错误**（不是注入故障）----
    // C1 只读目标：`chmodSync(0o444)` 在 Windows 上就是设"只读属性"（用 chmod 目录无效，故不用）。
    {
      const ro = join(OUT, 'r18-readonly.txt')
      writeFileSync(ro, '只读目标-必须原样保留', 'utf8')
      chmodSync(ro, 0o444)
      let roErr = ''
      try { await writeFileAtomic(ro, Buffer.from('新内容')) } catch (e) { roErr = String(e.message) }
      const survived = existsSync(ro)
      const afterRo = survived ? readFileSync(ro, 'utf8') : '(目标已被删除!)'
      chmodSync(ro, 0o666)
      ok('R18-原子写：只读目标 → 可读四要素错误，且**绝不 unlink 只读目标**',
        survived && /【写盘失败｜四要素】/.test(roErr) && /已保留原文件/.test(roErr), roErr.slice(0, 120))
      ok('R18-原子写：只读目标失败后内容逐字节不变', afterRo === '只读目标-必须原样保留', afterRo.slice(0, 40))
      ok('R18-原子写：只读目标失败后不留临时件', strayTmp().length === 0, strayTmp().join(', '))
      ok('R18-原子写：只读目标的根因是真实 Fs 错误码（不是注入故障的文案）',
        /EPERM|EACCES|EBUSY|EPERM: operation not permitted|access/i.test(roErr) && !/注入故障/.test(roErr),
        roErr.replace(/\s+/g, ' ').slice(0, 150))
    }
    // C1' 同步版（sidecar 路径 writeFileAtomicSync）同等断言
    {
      const roS = join(OUT, 'r18-readonly-sync.md')
      writeFileSync(roS, '同步侧-只读', 'utf8')
      chmodSync(roS, 0o444)
      let sErr = ''
      try { writeFileAtomicSync(roS, '新内容') } catch (e) { sErr = String(e.message) }
      const sSurvived = existsSync(roS)
      const sAfter = sSurvived ? readFileSync(roS, 'utf8') : '(已丢失!)'
      chmodSync(roS, 0o666)
      ok('R18-原子写：writeFileAtomicSync（sidecar 路径）只读目标 → 四要素错误 + 目标未被删',
        sSurvived && /【写盘失败｜四要素】/.test(sErr), sErr.slice(0, 110))
      ok('R18-原子写：writeFileAtomicSync 只读失败后原内容逐字节不变 + 不留临时件',
        sAfter === '同步侧-只读' && strayTmp().length === 0, `${sAfter} | ${strayTmp().join(',')}`)
    }
    // C2 目标被占用：① 同进程句柄（实测口径，断言"不产生半截文件"这个不变式）；
    //                  ② 跨进程 `FileShare.None` 真锁（Windows 语义上一定挡住 rename）。
    {
      const busy = join(OUT, 'r18-busy.txt')
      writeFileSync(busy, '被占用-原样', 'utf8')
      const beforeBusy = readFileSync(busy, 'utf8')
      const fh = openSync(busy, 'r+')
      let busyErr = ''
      let busyOk = false
      try { await writeFileAtomic(busy, Buffer.from('占用期写入')); busyOk = true } catch (e) { busyErr = String(e.message) }
      closeSync(fh)
      const nowBusy = readFileSync(busy, 'utf8')
      ok('R18-原子写：同进程持有句柄 → 要么原子成功、要么原目标逐字节不变（两者都不得留半个文件）',
        (busyOk && nowBusy === '占用期写入') || (!busyOk && nowBusy === beforeBusy),
        busyOk ? '同进程句柄没挡住 rename（Node 的 open 带 FILE_SHARE_DELETE）—— 如实记录，不假装测到了锁'
          : busyErr.replace(/\s+/g, ' ').slice(0, 120))
      ok('R18-原子写：同进程持有句柄场景不留临时件', strayTmp().length === 0, strayTmp().join(', '))

      // 跨进程真锁：子 PowerShell 以 FileShare.None 打开目标并保持，直到看到 release 标志。
      const lockFlag = join(OUT, 'r18-lock-ready.flag')
      const relFlag = join(OUT, 'r18-lock-release.flag')
      rmSync(lockFlag, { force: true })
      rmSync(relFlag, { force: true })
      const lockPs = `$fs=[System.IO.File]::Open('${busy}','Open','ReadWrite','None');`
        + `Set-Content -LiteralPath '${lockFlag}' -Value '1' -Encoding UTF8;`
        + `while(-not (Test-Path -LiteralPath '${relFlag}')){Start-Sleep -Milliseconds 100};`
        + '$fs.Close()'
      const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', lockPs],
        { stdio: 'ignore', windowsHide: true })
      const childDone = new Promise(r => child.on('exit', () => r(0)))
      const lockWait0 = Date.now()
      while (!existsSync(lockFlag) && Date.now() - lockWait0 < 20000) await sleep(100)
      const locked = existsSync(lockFlag)
      let lockErr = ''
      let lockOk = false
      let lockMs = 0
      if (locked) {
        const t0 = Date.now()
        try { await writeFileAtomic(busy, Buffer.from('锁下写入')); lockOk = true } catch (e) { lockErr = String(e.message) }
        lockMs = Date.now() - t0
      }
      writeFileSync(relFlag, '1', 'utf8')
      if (locked) await Promise.race([childDone, sleep(10000)])
      else { try { child.kill() } catch { /* ignore */ } }
      if (!locked) {
        ok('R18-原子写：跨进程 FileShare.None 真锁 → 跳过（20 s 内未建立锁）', true, 'PowerShell 子进程未就绪')
      } else {
        ok('R18-原子写：跨进程真锁下 → 写盘失败（四要素），原目标仍在且逐字节不变',
          !lockOk && /【写盘失败｜四要素】/.test(lockErr) && existsSync(busy) && readFileSync(busy, 'utf8') === beforeBusy,
          lockOk ? '真锁下竟然成功了（需要复核）' : lockErr.replace(/\s+/g, ' ').slice(0, 120))
        ok('R18-原子写：跨进程真锁的根因是 EPERM/EBUSY（Windows 锁语义），且不留临时件',
          /EPERM|EBUSY|EACCES/.test(lockErr) && strayTmp().length === 0,
          `${lockErr.replace(/\s+/g, ' ').slice(0, 100)} | tmp=${strayTmp().join(',')}`)
        ok('R18-原子写：真锁场景**重试窗口真的被走到**（≥3 次退避，60+120+180ms）',
          lockMs >= 300, `耗时 ${lockMs}ms（阈值 300ms = 三次退避之和的上界）`)
      }
    }

    // ---- R19-C1：**瞬时锁**（子进程持锁 ~250 ms 后释放）→ 退避窗口内重试并成功 ----
    // R18 只证明了"锁一直不放 → 失败 + 退避窗口被走到"；这里补另一半：
    // 锁在窗口内释放 → **重试后成功**，且终态是新内容、原文件未被破坏、无临时件残留。
    // 用"反向握手"保证首次 rename 一定落在锁持有期内（否则测不出重试）：
    //   子进程拿锁 → 写 lock 标志 → 等 go 标志 → 再 Sleep 250 ms → 释放；
    //   主进程看到 lock 标志 → 写 go 标志 → 立刻 writeFileAtomic。
    {
      const tf = join(OUT, 'r19-transient-lock.txt')
      writeFileSync(tf, '瞬时锁-旧内容', 'utf8')
      const lockFlag = join(OUT, 'r19-transient-lock.flag')
      const relFlag = join(OUT, 'r19-transient-release.flag')
      rmSync(lockFlag, { force: true })
      rmSync(relFlag, { force: true })
      // 反向握手（R19 实测修正）：子进程拿锁 → 写 lock 标志 → 等 **release 标志** → 关闭；
      // 主进程看到 lock 标志后，**先挂一个 70 ms 后写 release 标志的定时器、再立刻开始写** ——
      // "首次 rename 一定落在锁持有期内"（否则测不出重试）且"释放时刻"由主进程控制，
      // 不依赖 PowerShell 的 Start-Sleep 精度（第一版让子进程自己 Sleep 250 ms，实测释放晚于
      // 360 ms 退避窗口上界 ⇒ 用例红；那是夹具时序问题，**不是实现问题**）。
      const lockPs = `$fs=[System.IO.File]::Open('${tf}','Open','ReadWrite','None');`
        + `Set-Content -LiteralPath '${lockFlag}' -Value '1' -Encoding UTF8;`
        + `while(-not (Test-Path -LiteralPath '${relFlag}')){Start-Sleep -Milliseconds 10};`
        + '$fs.Close()'
      const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', lockPs],
        { stdio: 'ignore', windowsHide: true })
      const childDone = new Promise(r => child.on('exit', () => r(0)))
      const w0 = Date.now()
      while (!existsSync(lockFlag) && Date.now() - w0 < 20000) await sleep(50)
      const locked = existsSync(lockFlag)
      let okFlag = false, err = '', ms = 0
      if (locked) {
        const releaseTimer = setTimeout(() => { try { writeFileSync(relFlag, '1', 'utf8') } catch { /* ignore */ } }, 70)
        const t0 = Date.now()
        try { await writeFileAtomic(tf, Buffer.from('瞬时锁-新内容')); okFlag = true } catch (e) { err = String(e.message) }
        ms = Date.now() - t0
        clearTimeout(releaseTimer)
      }
      await Promise.race([childDone, sleep(15000)])
      if (!locked) {
        ok('R19-原子写：瞬时锁 → 跳过（20 s 内未建立锁）', true, 'PowerShell 子进程未就绪')
      } else {
        const finalTxt = existsSync(tf) ? readFileSync(tf, 'utf8') : '(目标已丢失!)'
        ok('R19-原子写：瞬时锁在退避窗口内释放 → writeFileAtomic **重试后成功**（R18 只证明了"窗口被走到"）',
          okFlag, okFlag ? `耗时 ${ms}ms` : err.replace(/\s+/g, ' ').slice(0, 150))
        ok('R19-原子写：瞬时锁重试成功 → 终态是新内容、原文件未被破坏（不是半截、也不是旧内容）',
          finalTxt === '瞬时锁-新内容', finalTxt.slice(0, 60))
        ok('R19-原子写：瞬时锁重试成功 → 不留临时件，且真重试过（耗时 ≥ 1 次退避 60 ms）',
          strayTmp().length === 0 && ms >= 60,
          `耗时 ${ms}ms（writeFileAtomic 开始后 70 ms 释放锁；退避表 60/120/180 ms）tmp=${strayTmp().join(',') || 'none'}`)
      }
    }

    // ---- R19-C2：跨卷 rename → EXDEV（"临时件必须放目标同目录"这条约束的直接证据）----
    // 网络盘 = 另一个卷：若临时件放 `tmpdir()`（C:）、发布目标是网络盘，就会 EXDEV。
    // 本机没有网络盘（见 DEVELOPMENT.md R19 §C2），所以这里用"同机不同卷"给出等价的行为证据。
    {
      const volOf = p => resolve(p).slice(0, 2).toUpperCase()
      const tmpVol = volOf(tmpdir())
      const cands = [process.env.DSH_OFFICE_TEST_OTHER_VOL, resolve(process.cwd()), resolve(HERE), 'D:\\', 'E:\\'].filter(Boolean)
      let other = null
      for (const c of cands) {
        if (volOf(c) === tmpVol) continue
        try {
          const probe = join(c, `.r19-xdev-probe-${Date.now()}`)
          writeFileSync(probe, 'x')
          rmSync(probe, { force: true })
          other = c
          break
        } catch { /* 试下一个候选 */ }
      }
      if (!other) {
        ok('R19-原子写：跨卷 rename → EXDEV（找不到与 %TEMP% 不同的可写卷 → 跳过）', true,
          `tmp 卷=${tmpVol}；候选=${cands.join(', ')}`)
      } else {
        const src = join(tmpdir(), `r19-xdev-${Date.now()}.txt`)
        writeFileSync(src, 'x')
        const dst = join(other, `r19-xdev-${Date.now()}.txt`)
        let code = ''
        try { renameSync(src, dst) } catch (e) { code = String(e.code || e.message) }
        try { rmSync(src, { force: true }) } catch { /* ignore */ }
        try { rmSync(dst, { force: true }) } catch { /* ignore */ }
        ok('R19-原子写：跨卷 rename → EXDEV（证明"临时件必须与目标同目录"是硬约束，不是风格偏好）',
          code === 'EXDEV', `vol ${tmpVol} → ${volOf(dst)}；code=${code || '(竟然成功了!)'}`)
      }
    }

    // ---- R19-C2'：fsync 开关在本地卷上不改语义（网络盘上的差异本轮无法验证）----
    {
      const on = join(OUT, 'r19-fsync-on.txt')
      const off = join(OUT, 'r19-fsync-off.txt')
      const t0 = Date.now()
      await writeFileAtomic(on, Buffer.from('fsync-on'))
      const msOn = Date.now() - t0
      process.env.DSH_OFFICE_ATOMIC_FSYNC = '0'
      const t1 = Date.now()
      await writeFileAtomic(off, Buffer.from('fsync-off'))
      const msOff = Date.now() - t1
      delete process.env.DSH_OFFICE_ATOMIC_FSYNC
      ok('R19-原子写：DSH_OFFICE_ATOMIC_FSYNC=0 与默认在本地卷上语义一致（开关只关持久性，不改"临时件→替换"）',
        readFileSync(on, 'utf8') === 'fsync-on' && readFileSync(off, 'utf8') === 'fsync-off' && strayTmp().length === 0,
        `fsync=on ${msOn}ms / fsync=off ${msOff}ms`)
    }

    // ---- R19-C2''：网络盘 / OneDrive（opt-in；本机没有，明确跳过而不是拿本地目录冒充）----
    if (process.env.DSH_OFFICE_TEST_NET_DIR) {
      const nd = process.env.DSH_OFFICE_TEST_NET_DIR
      const nf = join(nd, `r19-netdrive-${Date.now()}.txt`)
      writeFileSync(nf, '网络盘-旧内容', 'utf8')
      const t0 = Date.now()
      let netErr = ''
      let netOk = false
      try { await writeFileAtomic(nf, Buffer.from('网络盘-新内容')); netOk = true } catch (e) { netErr = String(e.message) }
      const netMs = Date.now() - t0
      const txt = existsSync(nf) ? readFileSync(nf, 'utf8') : '(丢失!)'
      try { rmSync(nf, { force: true }) } catch { /* ignore */ }
      ok('R19-原子写：网络盘 / OneDrive 上"写临时件 → rename 覆盖"成功且内容正确（opt-in）',
        netOk && txt === '网络盘-新内容', netOk ? `${netMs}ms` : netErr.replace(/\s+/g, ' ').slice(0, 150))
      ok('R19-原子写：网络盘上不留临时件（网络语义下同样不该留）', strayTmp().length === 0, strayTmp().join(','))
    } else {
      ok('R19-原子写：网络盘 / OneDrive 上的 rename+fsync 语义（opt-in）→ 跳过：本机无可用同步盘/网络盘',
        true, '设 DSH_OFFICE_TEST_NET_DIR=<网络盘目录> 打开（本机无映射网络驱动器；OneDrive 目录未登录/未同步）')
    }
  }

  // ===========================================================================
  // 任务四（R16）：工具参数 schema 与运行时行为一致
  // 前缀：R16-schema：
  // ===========================================================================
  {
    const gridSpec = captured.get('office_create').parameters.properties.grid
    ok('R16-schema：grid 用 oneOf 声明 string|number（host 子集支持 oneOf，禁用 type 数组）',
      Array.isArray(gridSpec.oneOf) && gridSpec.oneOf.length === 2
      && gridSpec.oneOf.some(s => s.type === 'string') && gridSpec.oneOf.some(s => s.type === 'number')
      && !Array.isArray(gridSpec.type), JSON.stringify(gridSpec.oneOf))
    const gNum = join(OUT, 'r16-grid-num.docx')
    const gStr = join(OUT, 'r16-grid-str.docx')
    await call('office_create', { path: gNum, markdown: '# 申论\n\n第一段。\n', grid: 20 })
    await call('office_create', { path: gStr, markdown: '# 申论\n\n第一段。\n', grid: '20' })
    const docX = openZip(readFileSync(gNum)).getText('word/document.xml')
    ok('R16-schema：grid=20（数字）端到端成功并真的写出 docGrid',
      /<w:docGrid w:type="linesAndChars"/.test(docX), (docX.match(/<w:docGrid[^>]*\/>/) || ['<无>'])[0])
    ok('R16-schema：grid=20 与 grid="20" 的 document.xml 逐字等价（字符串形式未被破坏）',
      docX === openZip(readFileSync(gStr)).getText('word/document.xml'))
    let gErr = ''
    try { await call('office_create', { path: join(OUT, 'r16-grid-bad.docx'), markdown: '# T', grid: true }) } catch (e) { gErr = String(e.message) }
    ok('R16-schema：grid 类型错误 → 文案列出两种接受类型', /参数 "grid" 应为字符串 或 数字/.test(gErr), gErr)
    let gErr2 = ''
    try { await call('office_create', { path: join(OUT, 'r16-grid-bad.odt'), markdown: '# T', grid: 20 }) } catch (e) { gErr2 = String(e.message) }
    ok('R16-schema：grid 数字对非 docx 目标仍显式报错（不静默忽略）', /【稿纸网格】/.test(gErr2), gErr2.slice(0, 100))

    ok('R16-schema：integer 参数在 schema 里不再被降级成 number（与 validator 一致）',
      captured.get('office_read').parameters.properties.offset.type === 'integer'
      && captured.get('office_read').parameters.properties.limit.type === 'integer')
    let offErr = ''
    try { await call('office_read', { path: gNum, offset: 1.5 }) } catch (e) { offErr = String(e.message) }
    ok('R16-schema：offset=1.5 报"应为整数"（schema 与实际校验一致）', /参数 "offset" 应为整数/.test(offErr), offErr)

    const editDef = captured.get('office_edit')
    const opsSpec = editDef.parameters.properties.operations.items
    ok('R16-schema：operations.items 声明 op 枚举（16 个）与 required=["op"]',
      Array.isArray(opsSpec.properties?.op?.enum) && opsSpec.properties.op.enum.length === 16
      && Array.isArray(opsSpec.required) && opsSpec.required.includes('op'), `${opsSpec.properties?.op?.enum?.length} 值`)
    ok('R16-schema：只用 host 子集允许的关键字（无 anyOf/$ref/pattern）',
      !/anyOf|\$ref|pattern|prefixItems/.test(JSON.stringify(opsSpec)))
    const badOps = [
      [[{ op: 'nope' }], /operations\[0\]\.op 未识别/],
      [['replace_text'], /operations\[0\] 应为对象/],
      [[{}], /operations\[0\] 缺少必填字段 "op"/],
      [[{ op: 'replace_text' }], /operations\[0\] 的 "replace_text" 需要 "find" 或 "regex"/],
      [[{ op: 'prepend_text' }], /缺少必填字段 "text"/],
    ]
    for (const [ops, re] of badOps) {
      let m = ''
      try { await call('office_edit', { path: gNum, operations: ops }) } catch (e) { m = String(e.message) }
      ok(`R16-schema：非法 operations ${JSON.stringify(ops)} → 带下标定位的参数错误`, re.test(m), m.slice(0, 110))
    }
    const mdOk = join(OUT, 'r16-ops.md')
    writeFileSync(mdOk, '甲\n乙\n', 'utf8')
    let okMsg = ''
    try { await call('office_edit', { path: mdOk, operations: [{ op: 'replace_text', find: '甲', replace: '丙' }] }) } catch (e) { okMsg = String(e.message) }
    ok('R16-schema：合法 replace_text 不被元素级校验误伤', okMsg === '' && readFileSync(mdOk, 'utf8').includes('丙'), okMsg)
    const xlsxOk = join(OUT, 'r16-ops.xlsx')
    await call('office_create', { path: xlsxOk, workbook: { sheets: [{ name: 'S1', columns: ['a'], rows: [['1']] }] } })
    let xMsg = ''
    try { await call('office_edit', { path: xlsxOk, operations: [{ op: 'set_cell', cell: 'A1', value: '2' }, { op: 'rename_sheet', name: 'S1', newName: 'S2' }] }) } catch (e) { xMsg = String(e.message) }
    ok('R16-schema：合法 xlsx 操作（set_cell + rename_sheet 全字段）不被误伤', xMsg === '', xMsg)
    let rnErr = ''
    try { await call('office_edit', { path: xlsxOk, operations: [{ op: 'rename_sheet', name: 'S2' }] }) } catch (e) { rnErr = String(e.message) }
    ok('R16-schema：rename_sheet 缺 newName → 参数错误（不再落到执行期写出 undefined）',
      /newName/.test(rnErr) && /operations\[0\]/.test(rnErr), rnErr)

    ok('R16-schema：插图 width 描述与实现一致（无"默认 450"，写清 72/96 与 451.3pt 上限）',
      !/默认 450/.test(editDef.description) && /72\/96/.test(editDef.description) && /451\.3/.test(editDef.description)
      && /72\/96/.test(String(opsSpec.properties.width.description)))
    const asDesc = String(captured.get('office_read').parameters.properties.as.description)
    ok('R16-schema：as="json" 的描述写明 content 是 JSON 文本字符串',
      /content 是 JSON 文本字符串/.test(asDesc) && /JSON\.parse/.test(asDesc))
    const jsonRead = await call('office_read', { path: gNum, as: 'json' })
    ok('R16-schema：as="json" 实际返回 string 且可 JSON.parse（结构未变）',
      typeof jsonRead.content === 'string'
      && (() => { try { return typeof JSON.parse(jsonRead.content).kind === 'string' } catch { return false } })()
      && jsonRead.nextOffset === undefined, `type=${typeof jsonRead.content}`)
    ok('R16-schema：as="json" 的 render 原样吐出 JSON 文本（不二次包装）',
      captured.get('office_read').output.render({ as: 'json' }, jsonRead)[0].text === jsonRead.content)

    // ---- R18 任务 D：输出边界与 schema 加固 ----
    ok('R18-schema：output.schema 声明 properties.content.type="string"，且**没有** required',
      ['office_read', 'office_create', 'office_edit', 'office_convert'].every(n =>
        captured.get(n)?.output?.schema?.properties?.content?.type === 'string'
        && captured.get(n).output.schema.required === undefined),
      JSON.stringify(captured.get('office_read').output?.schema))
    // 动手前的核实：五条返回形态里"若有 content，必为字符串"（schema 不会误伤）
    const shapeChecks = []
    for (const args of [{ path: gNum, as: 'json' }, { path: gNum, as: 'meta' }, { path: gNum, as: 'markdown' }, { path: gNum, as: 'text' }, { paths: [gNum] }]) {
      const r = await call('office_read', args)
      shapeChecks.push(r.content === undefined || typeof r.content === 'string')
    }
    ok('R18-schema：as=json/meta/markdown/text + 批量盘点的 content 全是字符串（一条也没被新 schema 误伤）',
      shapeChecks.every(Boolean), JSON.stringify(shapeChecks))

    const badShapes = [
      [[null], /operations\[0\] 不能为空（null）/],
      [['replace_text'], /operations\[0\] 应为对象（形如 .*），实际收到字符串/],
      [[42], /operations\[0\] 应为对象（形如 .*），实际收到数字/],
      [[[{}]], /operations\[0\] 应为对象（形如 .*），实际收到数组/],
      [[{ op: 7 }], /operations\[0\]\.op 未识别：7/],
      [[{ op: { a: 1 } }], /operations\[0\]\.op 未识别：\{"a":1\}/],
      [[{ op: 'replace_text', find: 123 }], /operations\[0\]\.find 应为字符串/],
      [[{ op: 'replace_value', find: { a: 1 }, replace: 'x' }], /operations\[0\]\.find 应为字符串/],
      [[{ op: 'replace_text', find: 'a', replace: 5 }], /operations\[0\]\.replace 应为字符串/],
      [[{ op: 'append_text', text: 5 }], /operations\[0\]\.text 应为字符串/],
      [[{ op: 'append_markdown', markdown: 5 }], /operations\[0\]\.markdown 应为字符串/],
      [[{ op: 'append_rows', rows: 'a,b' }], /operations\[0\]\.rows 应为数组/],
      [[{ op: 'delete_slide', index: 1.5 }], /operations\[0\]\.index 应为整数/],
      [[{ op: 'update_slide', index: 1.5 }], /operations\[0\]\.index 应为整数/],
      [[{ op: 'rename_sheet', name: 'S1', newName: 5 }], /operations\[0\]\.newName 应为字符串/],
      [[{ op: 'set_cell', cell: 5 }], /operations\[0\]\.cell 应为字符串/],
      [[{ op: 'add_sheet', name: 5 }], /operations\[0\]\.name 应为字符串/],
    ]
    for (const [ops, re] of badShapes) {
      let m = ''
      try { await call('office_edit', { path: gNum, operations: ops }) } catch (e) { m = String(e.message) }
      ok(`R18-schema：非法元素 ${JSON.stringify(ops)} → 带 operations[N] 下标的清晰文案`, re.test(m), m.slice(0, 130))
    }
    {
      // 稀疏数组的空洞：forEach 会静默跳过 → 必须也点名（R18 起用下标循环）
      const sparse = new Array(2)
      sparse[1] = { op: 'replace_text', find: 'x' }
      let m = ''
      try { await call('office_edit', { path: gNum, operations: sparse }) } catch (e) { m = String(e.message) }
      ok('R18-schema：operations 稀疏数组的空洞 → 一样被点名（旧版 forEach 静默跳过）',
        /operations\[0\] 不能为空（undefined）/.test(m), m.slice(0, 130))
    }
    // 不误伤：本轮**没有**收窄那些"数字也合法"的字段
    let numValErr = ''
    try { await call('office_edit', { path: xlsxOk, operations: [{ op: 'set_cell', cell: 'A2', value: 5 }] }) } catch (e) { numValErr = String(e.message) }
    ok('R18-schema：set_cell 的 value 传数字仍被接受（本轮未收窄 value 类型 → 不误伤）', numValErr === '', numValErr)
    let reOkErr = ''
    try { await call('office_edit', { path: mdOk, operations: [{ op: 'replace_text', find: '乙', replace: '丁', regex: false, whole: true }] }) } catch (e) { reOkErr = String(e.message) }
    ok('R18-schema：replace_text 带齐 replace/regex/whole 不被误伤', reOkErr === '', reOkErr)
  }

  // ---------------- 返回边界：码点消毒 + 无损 JSON（G1 防回归） ----------------
  {
    const lone = sanitizeTextForReturn('A\uD800B')
    ok('消毒：游离代理（高位）→ U+FFFD', lone.fixedCount === 1 && lone.value === 'A\uFFFDB', JSON.stringify(lone.value))
    ok('消毒：游离代理（低位孤儿）→ U+FFFD', (() => { const r = sanitizeTextForReturn('A\uDFFFB'); return r.fixedCount === 1 && r.value === 'A\uFFFDB' })())
    ok('消毒：合法配对代理不动（emoji）', (() => { const r = sanitizeTextForReturn('👍文'); return r.fixedCount === 0 && r.value === '👍文' })())
    ok('消毒：合法配对代理不动（CJK 扩展 B）', (() => { const r = sanitizeTextForReturn('\u{20BB7}'); return r.fixedCount === 0 && r.value === '\u{20BB7}' })())
    ok('消毒：C0 控制符 → U+FFFD，\\t\\n\\r 保留', (() => { const r = sanitizeTextForReturn('a\u0000b\tc\nd\re'); return r.fixedCount === 1 && r.value === 'a\uFFFDb\tc\nd\re' })())
    ok('消毒：C1 控制符 → U+FFFD', (() => { const r = sanitizeTextForReturn('x\u0085y\u009Fz'); return r.fixedCount === 2 && !/[\u0085\u009F]/.test(r.value) })())
    ok('消毒：U+FFFE / U+FFFF → U+FFFD', (() => { const r = sanitizeTextForReturn('a\uFFFEb\uFFFFc'); return r.fixedCount === 2 && r.value === 'a\uFFFDb\uFFFDc' })())
    ok('消毒：NFC 规范化（e + U+0301 → é）', (() => { const r = sanitizeTextForReturn('e\u0301'); return r.value === '\u00e9' && r.nfcChanged })())
    ok('消毒：U+2028/2029 保留（合法行分隔符，宁留噪不删正文）', (() => { const r = sanitizeTextForReturn('a\u2028b'); return r.fixedCount === 0 && r.value === 'a\u2028b' })())
    ok('消毒：只替换不删除（长度守恒、不断句）', (() => { const s = '第一句。\u0000第二句！'; const r = sanitizeTextForReturn(s); return r.value.length === s.length && r.value.includes('第二句') && r.fixed[0].at === 4 })())
    ok('消毒：坏码点可精确定位（at + codepoint + kind）', (() => { const r = sanitizeTextForReturn('\uDC01\u0007'); return r.fixed.length === 2 && r.fixed[0].codepoint === 0xdc01 && r.fixed[0].kind === 'lone-surrogate' && r.fixed[1].kind === 'control' })())
    ok('消毒：结果恒为 NFC 且可 UTF-8 往返', (() => { const r = sanitizeTextForReturn('e\u0301\uD800\u2028'); return isNfc(r.value) && utf8RoundTrips(r.value) && scanBadCodePoints(r.value).bad === 0 })())

    // 判定器：语义与 host 侧的 snapshotJsonValue 对齐
    ok('判定器：认出 undefined 属性（本次事故形态）', losslessJsonProblem({ stats: { ocrCovered: undefined } }) === '$.stats.ocrCovered 是 undefined（host 会整条拒收）')
    ok('判定器：干净对象通过', losslessJsonProblem({ a: 1, b: 'x', c: [1, 2], d: null, e: true }) === null)
    ok('判定器：NaN / Infinity / -0 都不合法', losslessJsonProblem({ a: NaN }) !== null && losslessJsonProblem({ a: Infinity }) !== null && losslessJsonProblem({ a: -0 }) !== null)
    ok('判定器：稀疏数组不合法', losslessJsonProblem({ a: [, 1] }) !== null)
    ok('判定器：Date / Map / 函数 等非纯对象不合法', losslessJsonProblem({ a: new Date() }) !== null && losslessJsonProblem({ a: new Map() }) !== null && losslessJsonProblem({ a: () => 1 }) !== null)
    ok('判定器：循环引用不合法', (() => { const c = { a: 1 }; c.self = c; return losslessJsonProblem({ c }) !== null })())

    // 收口：finalizeToolValue 是唯一出站终点
    const fixed1 = finalizeToolValue({ path: 'p', stats: { ocrCovered: undefined, ok: 1 } }, {})
    ok('收口：undefined 属性被剔除（回归本次整条拒收的元凶）', losslessJsonProblem(fixed1) === null && !('ocrCovered' in fixed1.stats) && fixed1.stats.ok === 1)
    ok('收口：数组里的 undefined → null（不留洞）', (() => { const v = finalizeToolValue({ a: [1, undefined, 3] }, {}); return losslessJsonProblem(v) === null && v.a[1] === null })())
    ok('收口：Date / Map / Set / NaN 全部压平为纯 JSON', (() => {
      const v = finalizeToolValue({ d: new Date(0), m: new Map([['k', 1]]), s: new Set([1]), n: NaN }, {})
      return losslessJsonProblem(v) === null && v.d === new Date(0).toISOString() && v.m.k === 1 && Array.isArray(v.s) && v.n === null
    })())
    ok('收口：循环引用不炸、不丢整条', (() => { const c = { a: 1 }; c.self = c; const v = finalizeToolValue({ c }, {}); return losslessJsonProblem(v) === null && v.c.a === 1 })())
    ok('收口：循环引用/丢弃都进 sanitizeNotes（绝不静默）', (() => { const c = { a: 1 }; c.self = c; const v = finalizeToolValue({ content: 'x', stats: {}, c }, {}); return (v.stats.sanitizeNotes || []).some(n => n.includes('循环引用')) })())
    ok('收口：BigInt → 字符串', losslessJsonProblem(finalizeToolValue({ n: 10n }, {})) === null)
    ok('收口：Buffer 等二进制 → 描述对象', (() => { const v = finalizeToolValue({ b: Buffer.from([1, 2, 3]) }, {}); return v.b.type === 'binary' && v.b.bytes === 3 })())

    // 正文脚注 + 记账
    const foot = finalizeToolValue({ content: '正文\u0000尾巴', stats: {} }, { as: 'markdown' })
    ok('收口：正文脚注 + stats.sanitized 记账（绝不静默）', foot.content.includes('1 个非法码点已替换为 U+FFFD') && foot.stats.sanitized === 1)
    ok('收口：无坏码点时脚注不出现（干净文件逐字不变）', (() => { const v = finalizeToolValue({ content: '干净正文', stats: {} }, { as: 'markdown' }); return v.content === '干净正文' && v.stats.sanitized === 0 })())
    ok('收口：as=json/meta 的脚注走 notice（不破坏可解析性）', (() => {
      const rawJson = `{"a":1,"t":"x${'\u0000'}y"}`
      const v = finalizeToolValue({ content: rawJson, stats: {} }, { as: 'meta' })
      const p = JSON.parse(v.content)
      return p.a === 1 && typeof p.notice === 'string' && p.notice.includes('1 个非法码点已替换为 U+FFFD')
    })())

    // 内联体积护栏（任务五）
    const big = '行\n'.repeat(90000)
    const saved = process.env.DSH_OFFICE_MAX_INLINE_CHARS
    process.env.DSH_OFFICE_MAX_INLINE_CHARS = '5000'
    const g = finalizeToolValue({ content: big, stats: {} }, { as: 'markdown' })
    ok('护栏：超阈值改为首段 + nextOffset 续读协议', g.truncated === true && typeof g.nextOffset === 'number' && g.nextOffset > 3000 && g.content.length < big.length)
    // 任务五：截断说明从 content 移到 notice + stats.truncateNote（content 只放纯前缀）
    ok('护栏：截断说明写在 notice 与 stats.truncateNote，不掺进 content',
      !g.content.includes('内联上限') && String(g.notice).includes('内联上限 5000') && String(g.notice).includes('offset:') && g.stats.truncateNote === g.notice,
      String(g.notice).slice(0, 60))
    ok('护栏：硬不变式 offset + content.length === nextOffset', (() => {
      const v = finalizeToolValue({ content: 'x'.repeat(40000), stats: {} }, { as: 'markdown', offset: 100 })
      return v.truncated === true && 100 + v.content.length === v.nextOffset
    })(), '带 offset 时也必须对得上账')
    ok('护栏：切在段落边界（不劈开一行）', g.content.slice(0, g.nextOffset).endsWith('行') || !g.content.slice(0, g.nextOffset).endsWith('\n'))
    ok('护栏：显式给了 limit 就不拦（opt-out 语义不变）', (() => { const v = finalizeToolValue({ content: big, stats: {} }, { as: 'markdown', limit: 200000 }); return v.content === big && v.truncated === undefined })())
    process.env.DSH_OFFICE_MAX_INLINE_CHARS = '0'
    ok('护栏：DSH_OFFICE_MAX_INLINE_CHARS=0 关闭', finalizeToolValue({ content: big, stats: {} }, { as: 'markdown' }).content === big)
    if (saved === undefined) delete process.env.DSH_OFFICE_MAX_INLINE_CHARS
    else process.env.DSH_OFFICE_MAX_INLINE_CHARS = saved

    // repro.mjs 的四查：真实样本上必须干净
    const md = await call('office_read', { path: docx, as: 'markdown' })
    ok('四查：docx 正文无游离代理 / 控制符 / 非 NFC / UTF-8 可往返', describeStringProblems(md.content, 'docx 正文') === null)
    ok('四查：工具返回值本身无损（office_read docx）', losslessJsonProblem(md) === null)
    for (const [name, value] of Object.entries({
      'office_create': await call('office_read', { path: xlsx, as: 'json' }),
      'office_edit': await call('office_read', { path: join(OUT, 'notes.odt'), as: 'meta' }),
      'office_convert': await call('office_read', { path: join(OUT, 'table.csv'), as: 'text' }),
    })) ok(`四查：${name} 对应读取路径返回无损`, losslessJsonProblem(value) === null)
  }

  // ---------------- 正文质量门 + 读取降级链（G2 防回归） ----------------
  {
    ok('质量门：正常中文不误判', textQuality('这是一段正常的中文正文，用于验证质量门不会误报。' + '需求是指消费者愿意购买的数量。'.repeat(5)).garbled === false)
    ok('质量门：正常英文不误判（不拿 CJK 覆盖率当闸门）', textQuality('This is a perfectly normal English paragraph used to verify that the quality gate does not misfire on non-CJK documents. '.repeat(4)).garbled === false)
    ok('质量门：私用区码点成片 → 判定乱码', (() => { const q = textQuality('\uE0A1\uE0A2\uE0A3'.repeat(200)); return q.garbled && q.privateUseRatio > 0.9 && q.reasons.some(r => r.includes('私用区')) })())
    ok('质量门：替换字符成片 → 判定乱码', (() => { const q = textQuality('\uFFFD'.repeat(200)); return q.garbled && q.reasons.some(r => r.includes('替换字符')) })())
    ok('质量门：控制符成片 → 判定乱码', textQuality('\u0001\u0002\u0003'.repeat(200)).garbled === true)
    ok('质量门：文本过短不触发（<40 可见字符不判）', textQuality('\uE0A1\uE0A2').garbled === false)
    ok('质量门：相对判据 —— 某页 CJK 覆盖率远低于全书且非英文页', (() => {
      const baseline = textQuality('需求是指消费者愿意购买的数量。'.repeat(20))
      const page = textQuality('¡¢£¤¥¦§¨©ª«¬®¯°±²³´µ¶·¸¹º»¼½¾¿'.repeat(20), baseline)
      return baseline.cjkRatio > 0.9 && page.garbled && page.reasons.some(r => r.includes('CJK 覆盖率'))
    })())
    ok('质量门：相对判据不误伤正常页', (() => { const q = textQuality('需求是指消费者愿意购买的数量，这是一段正常的页。'.repeat(6), { cjkRatio: 0.9 }); return q.garbled === false })())
    ok('质量门：中文书里的英文/代码页不被误判（相对判据带 ASCII 字母护栏）', (() => {
      const baseline = textQuality('需求是指消费者愿意购买的数量。'.repeat(20))
      const en = textQuality(('const total = rows.reduce((a, b) => a + b.value, 0); // 汇总\n').repeat(6), baseline)
      return baseline.cjkRatio > 0.9 && en.cjkRatio < 0.05 && en.garbled === false
    })())

    // sidecar 兜底：正文不可读时落 .read.md，返回值本身仍可无损序列化
    const garbled = join(OUT, 'garbled.md')
    const junk = `# 乱码样本\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(400)}`
    writeFileSync(garbled, junk, 'utf8')
    const rb = await call('office_read', { path: garbled, as: 'markdown' })
    ok('降级：正文不可读 → stats.fallback = sidecar', rb.stats.fallback === 'sidecar', JSON.stringify(rb.stats.fallback))
    ok('降级：sidecar 返回值本身无损（host 一定收）', losslessJsonProblem(rb) === null)
    ok('降级：返回值带可续读的文件路径', typeof rb.sidecar === 'string' && existsSync(rb.sidecar) && rb.sidecar === readSidecarPath(garbled))
    ok('降级：返回值只带首部摘录（不灌整篇乱码）', rb.content.length < junk.length && rb.truncated === true)
    ok('降级：notice 给出可复制的下一步参数串', rb.notice.includes('ocr="always"') && rb.notice.includes('read'))
    ok('降级：sidecar 正文与原文逐字一致（同一套消毒规则）', (() => {
      const t = readFileSync(rb.sidecar, 'utf8')
      return t.includes(junk.split('\n\n')[1]) && scanBadCodePoints(t).bad === 0
    })())
    ok('降级：stats.quality 记下判定依据（绝不静默）', rb.stats.quality?.privateUseRatio > 0.9 && rb.stats.quality.reasons.length > 0)

    // 干净文件不触发任何降级
    const clean = await call('office_read', { path: docx, as: 'markdown' })
    ok('降级：干净文件不触发（无 quality 字段、fallback=none）', clean.stats.fallback === 'none' && clean.stats.quality === undefined && clean.sidecar === undefined)
    ok('降级：干净文本文件也不触发', (await call('office_read', { path: join(OUT, 'plain.txt'), as: 'markdown' })).stats.fallback === 'none')

    // 无格式差异：.read.md 属于受保护缓存名，清扫不许碰
    ok('降级：.read.md 与 .ocr.md 同列受保护，渲染清理永不删成果', (() => {
      const root = join(tmpdir(), 'dsh-office-ocr')
      mkdirSync(root, { recursive: true })
      const probe = join(root, 'sweep-probe.read.md')
      writeFileSync(probe, 'keep me', 'utf8')
      utimesSync(probe, new Date(Date.now() - 60 * 86400000), new Date(Date.now() - 60 * 86400000))
      const rep = sweepRenderCache(join(OUT, 'plain.txt'), { now: Date.now() })
      const noCachedRemoved = rep.removed.every(p => !/\.(ocr|read)\.(md|json)$/i.test(p))
      return existsSync(probe) && noCachedRemoved
    })())

    // 错误分支四要素
    let hint = ''
    try { await call('office_read', { path: join(OUT, 'definitely-missing.docx') }) } catch (e) { hint = e.message }
    ok('错误四要素：页码 / 格式 / 根因 / 下一步 齐备', ['页码=', '格式=', '根因=', '下一步='].every(k => hint.includes(k)), hint.split('\n')[1]?.slice(0, 80))
    ok('错误四要素：不再出现裸 invalid output', !/invalid output/i.test(hint))
  }

  // ---------------- 全格式回归矩阵：4 形态 × 不重不漏（任务四） ----------------
  {
    const matrix = ['report.docx', 'data.xlsx', 'deck.pptx', 'notes.odt', 'sheet.ods', 'slides.odp',
      'table.csv', 'data.csv', 'plain.txt', 'report.txt', 'note.md', 'deck.md', 'report.pdf', 'report.odt']
    const present = matrix.filter(f => existsSync(join(OUT, f)))
    ok('矩阵：至少覆盖 10 个格式夹具', present.length >= 10, `${present.length}/${matrix.length}：${present.join(', ')}`)
    for (const f of present) {
      const p = join(OUT, f)
      let fourOk = true
      const detail = []
      for (const as of ['meta', 'markdown', 'text', 'json']) {
        try {
          const r = await call('office_read', { path: p, as })
          if (losslessJsonProblem(r) !== null) { fourOk = false; detail.push(`${as}:无损失败`) }
          else if (typeof r.content !== 'string' || !r.content.length) { fourOk = false; detail.push(`${as}:空正文`) }
        } catch (e) { fourOk = false; detail.push(`${as}:${e.message.slice(0, 40)}`) }
      }
      ok(`矩阵：${f} 的 meta/markdown/text/json 四形态全部成功且无损`, fourOk, detail.join('；'))
    }

    // 分段拼接不重不漏（读长文本的首选协议：用 truncated + nextOffset 驱动）
    const whole = (await call('office_read', { path: join(OUT, 'report.txt'), as: 'text' })).content
    let stitched = ''
    let off = 0
    let pieces = 0
    for (let guard = 0; guard < 200; guard++) {
      const r = await call('office_read', { path: join(OUT, 'report.txt'), as: 'text', offset: off, limit: 200 })
      // content 尾部可能带"已达上限"的说明后缀：真正属于这一段的正文是 [0, nextOffset-offset)
      const piece = r.nextOffset !== undefined ? r.content.slice(0, r.nextOffset - off) : r.content
      stitched += piece
      pieces++
      if (!r.truncated || r.nextOffset === undefined || r.nextOffset <= off) break
      off = r.nextOffset
    }
    ok('矩阵：report.txt 分段拼接不重不漏', stitched === whole, `${stitched.length} / ${whole.length}，${pieces} 段`)
    const firstChunk = await call('office_read', { path: join(OUT, 'report.txt'), as: 'text', offset: 0, limit: 200 })
    ok('矩阵：限长读取给出 truncated=true 与精确 nextOffset（旧版非 PDF 的 truncated 恒为 false）', firstChunk.truncated === true && firstChunk.nextOffset === 200, `truncated=${firstChunk.truncated} nextOffset=${firstChunk.nextOffset}`)

    // 真实 PDF（本样本）正文完整性：不得整段缺失
    const cidSample = process.env.DSH_OFFICE_TEST_CID_PDF || ''
    if (existsSync(cidSample)) {
      const local = join(OUT, 'cid-sample.pdf')
      await copyFile(cidSample, local)
      const raw = readPdfText(readFileSync(local))
      const r1 = await call('office_read', { path: local, pages: '1' })
      const p1 = String(raw.sections[0].text)
      ok('样本 PDF：pages="1" 正文完整（无整段缺失）', r1.content.includes(p1.slice(0, 200)) && r1.content.includes(p1.slice(-200)), `${r1.content.length} 字符`)
      ok('样本 PDF：pages="1" 返回无损（旧版此处整条拒收）', losslessJsonProblem(r1) === null)
      ok('样本 PDF：pages="1" 不误触降级', r1.stats.fallback === 'none' && r1.sidecar === undefined)
      const all = await call('office_read', { path: local, pages: '1-43' })
      ok('样本 PDF：pages="1-43" 全本可读', all.content.includes(p1.slice(0, 120)) && all.stats.pages === 43 && losslessJsonProblem(all) === null, `${all.content.length} 字符`)
      const meta = await call('office_read', { path: local, as: 'meta' })
      ok('样本 PDF：as=meta 给出分批建议与单批字符估计', meta.stats.suggestion?.plan?.length === 3 && meta.stats.suggestion.estCharsPerBatch > 0, JSON.stringify(meta.stats.suggestion?.plan))
      const js = await call('office_read', { path: local, pages: '3', as: 'json' })
      ok('样本 PDF：as=json pages="3" 可解析', losslessJsonProblem(js) === null && (() => { JSON.parse(js.content); return true })())
      const ocr = await call('office_read', { path: local, pages: '1', ocr: 'always', ocrEngine: 'local' })
      ok('样本 PDF：ocr="always" ocrEngine="local" 可读且无损', losslessJsonProblem(ocr) === null && ocr.stats.fallback === 'ocr', `fallback=${ocr.stats.fallback}`)
      // 消毒规则在 read 路径与 convert 落盘之间一致
      const convTxt = join(OUT, 'cid-sample.txt')
      await call('office_convert', { source: local, target: convTxt })
      const disk = readFileSync(convTxt, 'utf8')
      ok('样本 PDF：convert 落盘与 read 正文同为零坏码点（同一套消毒规则）', scanBadCodePoints(disk).bad === 0 && scanBadCodePoints(all.content).bad === 0 && utf8RoundTrips(disk) && isNfc(disk))
      // `p1` 是原始抽取（`readPdfText`），`disk` 是 `as="markdown"` 的产物：现在断行与标题都修好了，
      // 落盘正文会多出 `# ` 标题标记、页脚数字也归了位 —— 前缀不同是**修复的必然结果**，不是回归。
      // 断言改成"忽略 markdown 标记与空白后的前 60 个字符一致"。
      const norm = s => String(s).replace(/<!--[^>]*-->/g, '').replace(/^#+\s*/gm, '').replace(/\s/g, '')
      ok('样本 PDF：convert 落盘正文与页 1 文本一致（忽略空白与标题标记）',
        norm(disk).includes(norm(p1).slice(0, 60)),
        `normHead="${norm(disk).slice(0, 60)}" want="${norm(p1).slice(0, 60)}"`)
      // 验收 §3.5：干净对照样本的既有行为必须"逐字不变"——新机制一点痕迹都不该留
      ok('对照样本：干净文件四条路径无降级痕迹（fallback=none、无 stats.quality、建议仍是分批计划）',
        [r1, all, meta].every(x => x.stats?.fallback === 'none' && x.stats?.quality === undefined)
        && meta.stats.textLayerUsable === true && Array.isArray(meta.stats.garbledPages) && meta.stats.garbledPages.length === 0
        && Array.isArray(meta.stats.suggestion?.plan) && meta.stats.suggestion.plan.length > 0
        && existsSync(join(OUT, 'cid-sample.read.md')) === false,
        `usable=${meta.stats.textLayerUsable} garbled=${JSON.stringify(meta.stats.garbledPages)}`)
      ok('对照样本：ocrFailedPages 常驻为空数组（干净读取不该缺页）',
        Array.isArray(all.stats.ocrFailedPages) && all.stats.ocrFailedPages.length === 0, JSON.stringify(all.stats.ocrFailedPages))
    } else {
      ok('样本 PDF：夹具缺失 —— 未执行（用 DSH_OFFICE_TEST_CID_PDF 指定）', false)
    }
  }

  // ======================================================================
  // 第二轮补丁：截断协议 / 缓存位置 / meta 质量画像 / 出站质量门 / 换倍率重试
  // ======================================================================

  // ---------------- 任务五：截断协议收敛 ----------------
  {
    const cap = capWithOffset('abcdefghij'.repeat(5), 10, 7)
    ok('护栏：capWithOffset 的 content 是纯前缀（不带截断说明）', cap.content === 'abcdefghij' && cap.truncated === true, JSON.stringify(cap.content))
    ok('护栏：capWithOffset 满足 offset + content.length === nextOffset', 7 + cap.content.length === cap.nextOffset, `7+${cap.content.length} vs ${cap.nextOffset}`)
    ok('护栏：capWithOffset 的截断说明走 note 字段', typeof cap.note === 'string' && cap.note.includes('offset:17'), cap.note.slice(0, 50))
    ok('护栏：未截断时 note 为空、nextOffset 为 undefined', (() => { const c = capWithOffset('短文本', 100, 0); return c.truncated === false && c.nextOffset === undefined && c.note === '' })())

    const bigTxt = join(OUT, 'guardrail-big.txt')
    writeFileSync(bigTxt, '第X段正文内容。'.repeat(3000), 'utf8')          // 24000 字，无换行
    const savedInline = process.env.DSH_OFFICE_MAX_INLINE_CHARS
    process.env.DSH_OFFICE_MAX_INLINE_CHARS = '20000'
    try {
      const g2 = await call('office_read', { path: bigTxt, as: 'markdown' })
      ok('护栏：自动截断时 content 只放纯前缀（无 "正文超过" 后缀）',
        g2.content.length === 20000 && !g2.content.includes('内联上限') && !g2.content.includes('>'), `${g2.content.length} 字符`)
      ok('护栏：自动截断满足 offset + content.length === nextOffset', 0 + g2.content.length === g2.nextOffset, `nextOffset=${g2.nextOffset}`)
      ok('护栏：截断说明进 notice 与 stats.truncateNote',
        String(g2.notice).includes('内联上限') && String(g2.notice).includes('offset:') && String(g2.stats.truncateNote).includes('内联上限'),
        String(g2.notice).slice(0, 60))
    } finally {
      if (savedInline === undefined) delete process.env.DSH_OFFICE_MAX_INLINE_CHARS
      else process.env.DSH_OFFICE_MAX_INLINE_CHARS = savedInline
    }

    // 分段拼接：只用 offset + content.length 驱动（不再需要 slice 掉后缀），必须不重不漏
    const wholeTxt = (await call('office_read', { path: join(OUT, 'report.txt'), as: 'text' })).content
    let stitch2 = ''
    let off2 = 0
    let pieces2 = 0
    let invariantHeld = true
    for (let guard = 0; guard < 200; guard++) {
      const r2 = await call('office_read', { path: join(OUT, 'report.txt'), as: 'text', offset: off2, limit: 200 })
      stitch2 += r2.content
      pieces2++
      if (!r2.truncated || r2.nextOffset === undefined || r2.nextOffset <= off2) break
      if (r2.nextOffset !== off2 + r2.content.length) { invariantHeld = false; break }
      off2 = r2.nextOffset
    }
    ok('护栏：仅用 offset + content.length 驱动分段拼接，不重不漏',
      invariantHeld && stitch2 === wholeTxt, `${stitch2.length} / ${wholeTxt.length}，${pieces2} 段`)
  }

  // ---------------- 任务三：换倍率重试的可调表与目录隔离（纯函数面） ----------------
  {
    ok('重试倍率：默认表 2,1.5,1；空串关闭；垃圾值被滤掉且保序去重', (() => {
      const savedRetry = process.env.DSH_OFFICE_OCR_RETRY_SCALES
      delete process.env.DSH_OFFICE_OCR_RETRY_SCALES
      const def = ocrRetryScales()
      if (savedRetry !== undefined) process.env.DSH_OFFICE_OCR_RETRY_SCALES = savedRetry
      return JSON.stringify(def) === '[2,1.5,1]'
        && JSON.stringify(ocrRetryScales('')) === '[]'
        && JSON.stringify(ocrRetryScales('1.5, 3 ,abc,,1.5')) === '[1.5,3]'
        && JSON.stringify(ocrRetryScales('9')) === '[]'
    })(), `${JSON.stringify(ocrRetryScales('1.5, 3 ,abc,,1.5'))}`)
    ok('重试倍率：当前生效倍率的接受窗口与 pdf-render.ps1 一致（0.5–4）', (() => {
      const savedScale = process.env.DSH_OFFICE_RENDER_SCALE
      delete process.env.DSH_OFFICE_RENDER_SCALE
      const none = currentRenderScale()
      if (savedScale !== undefined) process.env.DSH_OFFICE_RENDER_SCALE = savedScale
      return none === null && currentRenderScale('') === null && currentRenderScale('2') === 2 && currentRenderScale('1.5') === 1.5
        && currentRenderScale('5') === null && currentRenderScale('abc') === null
    })())
    ok('重试倍率：重试渲染目录与默认目录分离（否则 PNG 命中缓存，重试变空转）',
      renderDirFor(join(OUT, 'report.pdf'), '2') !== renderDirFor(join(OUT, 'report.pdf'))
      && renderDirFor(join(OUT, 'report.pdf'), '2').endsWith('-s2'), renderDirFor(join(OUT, 'report.pdf'), '2'))
  }

  // ---------------- 任务四：sidecar 落盘位置可配（DSH_OFFICE_CACHE_DIR） ----------------
  {
    const probe = join(OUT, 'cache-probe.pdf')
    await copyFile(join(OUT, 'report.pdf'), probe)
    ok('缓存目录：未设 DSH_OFFICE_CACHE_DIR → 与旧行为逐字一致（源文件同目录）',
      ocrCachePath(probe) === `${probe.replace(/\.pdf$/i, '')}.ocr.md` && readSidecarPath(probe) === `${probe}.read.md`,
      ocrCachePath(probe))

    const cacheRoot = join(OUT, 'cache-root')
    process.env.DSH_OFFICE_CACHE_DIR = cacheRoot
    try {
      ok('缓存目录：设置后 .ocr.md / .read.md 都落到该目录，文件名带路径指纹（任务一：防同名文件撞车）',
        ocrCachePath(probe) === join(cacheRoot, `cache-probe-${pathKeyOf(probe)}.ocr.md`)
        && readSidecarPath(probe) === join(cacheRoot, `cache-probe.pdf-${pathKeyOf(probe)}.read.md`),
        ocrCachePath(probe))
      // 真跑一次降级，确认 sidecar 不是"只改了纯函数"
      const junk2 = join(OUT, 'cache-junk.txt')
      writeFileSync(junk2, `# 乱码样本\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(400)}`, 'utf8')
      const rb2 = await call('office_read', { path: junk2, as: 'markdown' })
      ok('缓存目录：降级 sidecar 真的写到该目录（带路径指纹）',
        rb2.sidecar === join(cacheRoot, `cache-junk.txt-${pathKeyOf(junk2)}.read.md`) && existsSync(rb2.sidecar), String(rb2.sidecar))
      ok('缓存目录：.read.md 仍在受保护名单内（清扫不碰）', /\.(ocr|read)\.(md|json)$/i.test(basename(rb2.sidecar || '')))
    } finally { delete process.env.DSH_OFFICE_CACHE_DIR }

    // 不可写 → 回退默认位置，并把原因写进 stats（绝不静默）
    process.env.DSH_OFFICE_CACHE_DIR = join(OUT, 'cache-junk.txt', 'not-a-dir')   // 拿文件当目录 → 必失败
    try {
      const st = cacheDirState()
      ok('缓存目录：不可写时回退默认位置并给出原因', st.dir === '' && /不可用/.test(st.note), st.note)
      ok('缓存目录：回退后路径与未设置时一致', ocrCachePath(probe) === `${probe.replace(/\.pdf$/i, '')}.ocr.md`)
      const junk3 = join(OUT, 'cache-junk2.txt')
      writeFileSync(junk3, `# 乱码样本\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(400)}`, 'utf8')
      const rb3 = await call('office_read', { path: junk3, as: 'markdown' })
      ok('缓存目录：回退原因进 stats.cacheDirNote（不静默）',
        rb3.stats.fallback === 'sidecar' && /DSH_OFFICE_CACHE_DIR 不可用/.test(String(rb3.stats.cacheDirNote)),
        String(rb3.stats.cacheDirNote || '（无）').slice(0, 70))
    } finally { delete process.env.DSH_OFFICE_CACHE_DIR }
  }

  // ===========================================================================
  // 任务一（R16）：OCR sidecar 的**缓存身份** —— 源路径 + 源内容双绑定
  // 前缀：缓存身份：
  // ===========================================================================
  {
    // ---- 1) 同一路径、内容被换掉：旧 OCR 缓存必须失效并说清原因 ----
    const idPdf = join(OUT, 'identity-swap.pdf')
    await call('office_create', { path: idPdf, markdown: '# 第一版\n\n甲内容\n' })
    writeOcrSidecar(idPdf, `covered: 1 | total: 1 | parser: ${PDF_PARSER_VERSION}`,
      '# OCR\n\n## 第 1 页（OCR）\n第一版识别文本\n')
    const idMeta1 = await call('office_read', { path: idPdf, as: 'meta' })
    ok('缓存身份：同一路径同一内容 → 缓存照常命中（身份门不是"一律作废"）',
      idMeta1.stats.ocrCacheStale === undefined && /covered: 1\b/.test(sidecarCoverage(idPdf)),
      sidecarCoverage(idPdf).slice(0, 110) || '<空>')

    // 同路径覆盖成另一份内容（另存为 / 重新扫描 / 覆盖下载的等价场景）
    await call('office_create', { path: idPdf, markdown: '# 第二版\n\n乙内容，比第一版长一些。\n' })
    const swapped = await call('office_read', { path: idPdf, as: 'meta' })
    ok('缓存身份：同一路径内容改变 → 旧缓存判 content-mismatch 作废（绝不静默复用旧 OCR 文本）',
      swapped.stats.ocrCacheStale?.reason === 'content-mismatch'
      && /已作废旧 OCR 缓存/.test(String(swapped.stats.ocrCacheNote))
      && /srcsha256/.test(String(swapped.stats.ocrCacheNote)),
      JSON.stringify(swapped.stats.ocrCacheStale || null))
    ok('缓存身份：内容变化后 sidecarCoverage 不再报 covered',
      !/covered: 1\b/.test(sidecarCoverage(idPdf)) && /已作废/.test(sidecarCoverage(idPdf)),
      sidecarCoverage(idPdf).slice(0, 130))

    // ---- 2) 同 basename、不同目录：不串用（同目录 sidecar 模式）----
    const dirA = join(OUT, 'id-dir-a')
    const dirB = join(OUT, 'id-dir-b')
    mkdirSync(dirA, { recursive: true })
    mkdirSync(dirB, { recursive: true })
    const pdfA = join(dirA, 'same-name.pdf')
    const pdfB = join(dirB, 'same-name.pdf')
    await call('office_create', { path: pdfA, markdown: '# 甲文件\n\n甲正文\n' })
    await call('office_create', { path: pdfB, markdown: '# 乙文件\n\n乙正文\n' })
    const sideA = writeOcrSidecar(pdfA, `covered: 1 | total: 1 | parser: ${PDF_PARSER_VERSION}`,
      '# OCR\n\n## 第 1 页（OCR）\n甲文件识别文本\n')
    ok('缓存身份：同 basename 不同目录 → 缓存路径互不相同（同目录 sidecar 也不共用）',
      sideA !== ocrCachePath(pdfB) && !existsSync(ocrCachePath(pdfB)),
      `${basename(sideA)} vs ${basename(ocrCachePath(pdfB))}`)
    const readB = await call('office_read', { path: pdfB, as: 'meta' })
    ok('缓存身份：没有自己缓存的同名文件不会被当成"有缓存"',
      readB.stats.ocrCacheStale === undefined && sidecarCoverage(pdfB) === '',
      sidecarCoverage(pdfB).slice(0, 80) || '<空>')

    // ---- 3) 集中缓存目录下也不碰撞：同名文件各读各的 ----
    const sharedRoot = join(OUT, 'id-shared-cache')
    process.env.DSH_OFFICE_CACHE_DIR = sharedRoot
    try {
      const cA = writeOcrSidecar(pdfA, `covered: 1 | total: 1 | parser: ${PDF_PARSER_VERSION}`,
        '# OCR\n\n## 第 1 页（OCR）\nAAA识别文本\n')
      const cB = writeOcrSidecar(pdfB, `covered: 1 | total: 1 | parser: ${PDF_PARSER_VERSION}`,
        '# OCR\n\n## 第 1 页（OCR）\nBBB识别文本\n')
      ok('缓存身份：集中缓存目录下同名文件各占一个文件（路径指纹去重）',
        cA !== cB && basename(cA) !== basename(cB) && existsSync(cA) && existsSync(cB),
        `${basename(cA)} / ${basename(cB)}`)
      const ra = await call('office_read', { path: pdfA, pages: '1', ocr: 'always', ocrEngine: 'local' })
      const rb = await call('office_read', { path: pdfB, pages: '1', ocr: 'always', ocrEngine: 'local' })
      ok('缓存身份：集中缓存目录下两个同名文件读到各自的内容（不串用、不覆盖）',
        String(ra.content).includes('AAA识别文本') && !String(ra.content).includes('BBB识别文本')
        && String(rb.content).includes('BBB识别文本') && !String(rb.content).includes('AAA识别文本'),
        `A=${(/AAA|BBB/.exec(String(ra.content)) || ['?'])[0]} B=${(/AAA|BBB/.exec(String(rb.content)) || ['?'])[0]}`)
    } finally { delete process.env.DSH_OFFICE_CACHE_DIR }

    // ---- 4) 旧命名（无路径指纹、无身份）的历史缓存：明确作废，绝不静默复用 ----
    const legacyRoot = join(OUT, 'id-legacy-cache')
    mkdirSync(legacyRoot, { recursive: true })
    process.env.DSH_OFFICE_CACHE_DIR = legacyRoot
    try {
      const legacyPdf = join(dirA, 'legacy-name.pdf')
      await call('office_create', { path: legacyPdf, markdown: '# 旧命名\n\n内容\n' })
      const legacyPath = join(legacyRoot, 'legacy-name.ocr.md')      // 旧规则：纯 basename
      writeFileSync(legacyPath, `<!-- dsh-office OCR cache -->\n`
        + `<!-- covered: 1 | total: 1 | parser: ${PDF_PARSER_VERSION} -->\n# OCR\n\n`
        + '## 第 1 页（OCR）\n旧命名缓存文本\n', 'utf8')
      const lMeta = await call('office_read', { path: legacyPdf, as: 'meta' })
      ok('缓存身份：集中目录里旧命名缓存被找到并明确作废（报真实路径 + identity-missing）',
        lMeta.stats.ocrCacheStale?.path === legacyPath
        && lMeta.stats.ocrCacheStale?.reason === 'identity-missing'
        && /已作废旧 OCR 缓存/.test(String(lMeta.stats.ocrCacheNote)),
        JSON.stringify(lMeta.stats.ocrCacheStale || null))
      ok('缓存身份：旧命名缓存不再被当作 covered（绝不静默复用）',
        !/covered: 1\b/.test(sidecarCoverage(legacyPdf)) && /已作废/.test(sidecarCoverage(legacyPdf)),
        sidecarCoverage(legacyPdf).slice(0, 130))
    } finally { delete process.env.DSH_OFFICE_CACHE_DIR }

    // ---- 5) 身份门通过后 covered/src/retry 仍跨批累积（新字段没挤掉旧账）----
    if (eng) {
      const accPdf = join(dirA, 'identity-accum.pdf')
      await call('office_create', {
        path: accPdf,
        document: {
          blocks: [
            { type: 'paragraph', runs: [{ text: '累加测试第一页' }] },
            { type: 'pagebreak' },
            { type: 'paragraph', runs: [{ text: '累加测试第二页' }] },
          ],
        },
      })
      writeOcrSidecar(accPdf, `covered: 1 | total: 2 | parser: ${PDF_PARSER_VERSION} | src: vision=1 | retry: 1=2`,
        '# OCR\n\n## 第 1 页（OCR）\n历史识别文本\n')
      await call('office_read', { path: accPdf, pages: '2', ocr: 'always', ocrEngine: 'local', limit: 200 })
      const accHead = readFileSync(ocrCachePath(accPdf), 'utf8').split('\n').slice(0, 2).join('\n')
      ok('缓存身份：身份门通过后 covered/src/retry 仍跨批累积（新字段没挤掉旧账）',
        /covered: 1-2/.test(accHead) && /vision=1/.test(accHead) && /rapidocr=2/.test(accHead) && /retry: 1=2/.test(accHead),
        accHead.replace(/\n/g, ' ').slice(0, 170))
      rmSync(ocrCachePath(accPdf), { force: true })
    } else {
      ok('缓存身份：跨批累积（本地引擎缺失 → 跳过）', true, '跳过：无 RapidOCR 引擎')
    }
  }

  // ---------------- 任务二：as=meta 质量画像（纯 CPU） ----------------
  {
    const cleanMeta = await call('office_read', { path: join(OUT, 'report.pdf'), as: 'meta' })
    ok('meta 画像：干净 PDF → textLayerUsable=true / garbledPages=[]',
      cleanMeta.stats.textLayerUsable === true && Array.isArray(cleanMeta.stats.garbledPages)
      && cleanMeta.stats.garbledPages.length === 0 && typeof cleanMeta.stats.qualityGate?.testedPages === 'number'
      && Array.isArray(cleanMeta.stats.qualityGate.reasons),
      `usable=${cleanMeta.stats.textLayerUsable} garbled=${JSON.stringify(cleanMeta.stats.garbledPages)} tested=${cleanMeta.stats.qualityGate?.testedPages}`)
    ok('meta 画像：既有字段未被改名改值',
      typeof cleanMeta.stats.pages === 'number' && typeof cleanMeta.stats.pagesWithText === 'number'
      && Array.isArray(cleanMeta.stats.scannedPages) && cleanMeta.stats.fallback === 'none')
    const dxMeta = await call('office_read', { path: docx, as: 'meta' })
    ok('meta 画像：非 PDF 也有 textLayerUsable', dxMeta.stats.textLayerUsable === true)
    const badMetaRaw = join(OUT, 'meta-garbled.md')
    writeFileSync(badMetaRaw, `# 乱码样本\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(400)}`, 'utf8')
    const badMeta = await call('office_read', { path: badMetaRaw, as: 'meta' })
    ok('meta 画像：乱码文件 → textLayerUsable=false', badMeta.stats.textLayerUsable === false && badMeta.stats.qualityGate?.reasons?.length > 0,
      JSON.stringify(badMeta.stats.qualityGate?.reasons?.[0] || ''))
  }

  // ===========================================================================
  // 第三轮：pdf.js 解析层修复（ObjStm off-by-First / mul 矩阵 / writePdf 死循环）
  // 前缀：解析器： / ObjStm： / 矩阵： / 死循环：
  // ===========================================================================
  {
    // —— 任务 A：`expandObjStms` 少加 `First` ——
    const fx = minimalObjStmPdf()
    const fxPdf = new PdfFile(fx.buf)
    ok('ObjStm：夹具确实踩在 off-by-First 路径上（First>0，且旧公式读不到对象本体）',
      fx.first > 0 && !fx.objStmText.slice(0, fx.fontBody.length).startsWith('<</Type/Font'),
      `First=${fx.first}`)
    ok('ObjStm：修后能从 ObjStm 里取出对象本体（逐字节等于原文）',
      (fxPdf.objs.get(4)?.body ?? '') === fx.fontBody, JSON.stringify((fxPdf.objs.get(4)?.body ?? '').slice(0, 40)))
    ok('ObjStm：修后页 /Resources 能解析出字体表（旧版这里是空表 → CID 被拆成单字节 → 整本乱码）',
      (() => {
        const pg = fxPdf.pages()[0]
        if (!pg) return false
        const fonts = fxPdf.fontsOf(pg.attrs.Resources)
        return fonts.size === 1 && fonts.get('F1')?.subtype === 'Type1'
      })())
    ok('解析器：ObjStm 夹具能读出正文（端到端）', /ObjStm works/.test(readPdfText(fx.buf).sections[0]?.text ?? ''))

    // 健壮性红线：畸形文件不得抛
    const fxBad = minimalObjStmPdf({ corruptOffset: true })
    let objStmThrew = null
    let badPages = -1
    try { badPages = new PdfFile(fxBad.buf).pages().length } catch (e) { objStmThrew = e.message }
    ok('ObjStm：`first + off` 越界（畸形）时不抛，只跳过该对象', objStmThrew === null && badPages === 1,
      objStmThrew || `pages=${badPages}`)
    const fxNoFirst = minimalObjStmPdf({ omitFirst: true })
    let noFirstThrew = null
    let noFirstObjs = -1
    try { const p = new PdfFile(fxNoFirst.buf); noFirstObjs = p.objs.has(4) ? 1 : 0 } catch (e) { noFirstThrew = e.message }
    ok('ObjStm：/First 缺失（=0）时行为与旧版一致（读不出该对象，且不抛）',
      noFirstThrew === null && noFirstObjs === 0, noFirstThrew || `obj4=${noFirstObjs}`)

    // —— 任务 B：`mul()` 的 e/f ——
    const ident = [1, 0, 0, 1, 0, 0]
    ok('矩阵：mul(ident, 平移) 的 e/f 不再是 NaN（旧版连最普通的纯平移都炸）',
      JSON.stringify(matrixMul(ident, [1, 0, 0, 1, 5, 7])) === JSON.stringify([1, 0, 0, 1, 5, 7]),
      JSON.stringify(matrixMul(ident, [1, 0, 0, 1, 5, 7])))
    ok('矩阵：mul(ident, ident) 恒等（旧版 e/f = NaN，`Math.round(NaN/2.5)` 于是成了合法 Map key）',
      JSON.stringify(matrixMul(ident, ident)) === JSON.stringify(ident))
    ok('矩阵：mul 不再读越界的 b[6]（b[6] 为 undefined 也不产生 NaN）',
      matrixMul([1, 0, 0, -1, 0, 841.89], [1, 0, 0, 1, 3, 4]).every(Number.isFinite),
      JSON.stringify(matrixMul([1, 0, 0, -1, 0, 841.89], [1, 0, 0, 1, 3, 4])))
    ok('矩阵：`mul(A,B)`＝"先 b 后 a"（scale(2,3) ∘ translate(5,7) 把 (1,1) 送到 (12,24)）',
      (() => {
        const m = matrixMul([2, 0, 0, 3, 0, 0], [1, 0, 0, 1, 5, 7])
        return (2 * 1 + m[4]) === 12 && (3 * 1 + m[5]) === 24
      })(), JSON.stringify(matrixMul([2, 0, 0, 3, 0, 0], [1, 0, 0, 1, 5, 7])))
    // 合成顺序 Trm = Tfs × Tlm × CTM：先字号、再文本矩阵、最后 CTM
    ok('矩阵：合成顺序 Trm = Tfs × Tlm × CTM（旋转 90° 时位移必须跟着转到设备轴）',
      (() => {
        const ctm = [0, -1, 1, 0, 0, 841.89]              // 90° 旋转
        const tlm = [1, 0, 0, 1, 100, 200]
        const tfs = [10.45, 0, 0, 10.45, 0, 0]
        const trm = matrixMul(ctm, matrixMul(tlm, tfs))
        const [x, y] = [trm[4], trm[5]]
        return Math.abs(x - 200) < 1e-9 && Math.abs(y - 741.89) < 1e-9
      })())
    ok('矩阵：`cm` 也是"先 cm 后旧 CTM"（旧版写反：带缩放的嵌套 cm 会算错）',
      (() => {
        const scaled = matrixMul([2, 0, 0, 3, 0, 0], [1, 0, 0, 1, 5, 7])   // CTM=scale, cm=translate
        return scaled[4] === 10 && scaled[5] === 21                        // 位移被 CTM 缩放
      })())

    // —— 任务 B：文本状态算子（Tr / Tz / Ts）真的生效 ——
    {
      const mk = content => {
        const c = content
        const objs = {}
        objs[1] = '<</Type/Catalog/Pages 2 0 R>>'
        objs[2] = '<</Type/Pages/Kids[3 0 R]/Count 1>>'
        objs[3] = '<</Type/Page/Parent 2 0 R/MediaBox[0 0 600 800]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>'
        objs[4] = '<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>'
        objs[5] = `<</Length ${c.length}>>\nstream\n${c}\nendstream`
        let out = '%PDF-1.4\n'
        for (let i = 1; i <= 5; i++) out += `${i} 0 obj\n${objs[i]}\nendobj\n`
        out += 'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n0\n%%EOF\n'
        return Buffer.from(out, 'latin1')
      }
      const visible = readPdfText(mk('BT /F1 12 Tf 1 0 0 1 20 700 Tm (VISIBLE) Tj ET')).sections[0].text
      const invisible = readPdfText(mk('BT /F1 12 Tf 1 0 0 1 20 700 Tm 3 Tr (HIDDEN) Tj ET')).sections[0].text
      ok('解析器：`3 Tr` 的不可见文字不算正文（水印/OCR 隐藏层不再混进来）',
        /VISIBLE/.test(visible) && !/HIDDEN/.test(invisible), `visible="${visible}" hidden="${invisible}"`)
      // 文本状态 `Ts`（以及 `Tz/Tc/Tw/TL/Tf/Tr`）按规范**不随 BT/ET 重置**，
      // 所以第二条必须显式 `0 Ts` 复位，否则两条 y 相同、会被合理地合并成 "RS"。
      const riseSrc = 'BT /F1 12 Tf 100 Ts 1 0 0 1 20 700 Tm (R) Tj ET BT /F1 12 Tf 0 Ts 1 0 0 1 20 700 Tm (S) Tj ET'
      const withRise = readPdfText(mk(riseSrc)).sections[0].text
      ok('解析器：`Ts` 升起量参与 y 坐标（两条同 x 的文字不再被当成同一行）',
        (() => {
          const t = readPdfText(mk(riseSrc))
          return t.sections[0].text.split('\n').length >= 2
        })(), JSON.stringify(withRise))
    }

    // —— 任务 C：`writePdf` 的同步死循环 ——
    // 旧版 `runPara` 的换行循环把 `fit` 夹在 ≥1：一旦"本行剩余宽度装不下一个字符"，
    // 判定条件恒为真 → **同步死转**，事件循环被占死（宿主整个冻住）。触发条件与字符种类
    // 无关：私用区最先撞上，汉字 / 扩展 B / `_` / ASCII 长串一样会挂。
    const hangProbe = runWithTimeout(20, `
      import { writePdf } from ${JSON.stringify(new URL('./pdf.js', import.meta.url).href)}
      for (const ch of ['\\uE0A1', '\\u4E00', '\\u{20000}', '_', 'x']) {
        for (const n of [1, 40, 46, 47, 50, 100, 200]) {
          const b = writePdf({ kind: 'document', meta: {}, blocks: [{ type: 'paragraph', runs: [{ text: ch.repeat(n) }] }] }, {})
          if (!b.length) throw new Error('empty')
        }
      }
    `)
    ok('死循环：私用区/汉字/扩展 B/`_`/ASCII 长度 1..200 扫描全部在超时前返回',
      !hangProbe.hung, hangProbe.hung ? '★ 同步死循环复现（进程被 SIGTERM）' : '未挂死')
    ok('死循环：产物可回读（写出来的 PDF 能被自己的解析器读出正文）',
      (() => {
        const buf = writePdf({ kind: 'document', meta: {}, blocks: [{ type: 'paragraph', runs: [{ text: '\uE0A1'.repeat(120) }] }] }, {})
        const r = readPdfText(buf)
        return r.pages >= 1 && (r.sections[0].text.match(/\uE0A1/g) || []).length === 120
      })())
    const longUnderscore = writePdf({ kind: 'document', meta: {}, blocks: [{ type: 'paragraph', runs: [{ text: '_'.repeat(200) }] }] }, {})
    ok('死循环：`_`.repeat(200) 的产物能被读回（§1-E 的最小复现）',
      longUnderscore.length > 500 && /_{3,}/.test(readPdfText(longUnderscore).sections[0].text))
  }

  // ---------------- 乱码文字层（真·缺 ToUnicode）端到端 ----------------
  // 第三轮起：这段的夹具是**自造**的 `buildGarbledFixture()`，不再依赖用户盘上的样本。
  // 原因：样本①的"整本乱码"是 `expandObjStms` off-by-First 造成的**假乱码**，修好之后
  // 它的文字层完全可用 —— 拿它当"乱码样本"的前提已经消失（保留它只会让这段失去意义）。
  // 真·缺 ToUnicode 的文档仍然存在，质量门 / 降级链 / convert 拒绝必须继续有回归保护。
  {
    const broken = join(OUT, 'garbled-layer.pdf')
    writeFileSync(broken, buildGarbledFixture())
    const side = ocrCachePath(broken)
    rmSync(side, { force: true })

    const bm = await call('office_read', { path: broken, as: 'meta' })
    ok('乱码夹具：meta 一眼看出文字层不可用（真·缺 ToUnicode）',
      bm.stats.textLayerUsable === false && String(bm.stats.garbledPages).length > 0,
      `usable=${bm.stats.textLayerUsable} garbled=${JSON.stringify(bm.stats.garbledPages)} pagesWithText=${bm.stats.pagesWithText}`)
    ok('乱码夹具：meta 的 qualityGate 给页码摘要（≤3 条）',
      Array.isArray(bm.stats.qualityGate?.reasons) && bm.stats.qualityGate.reasons.length > 0 && bm.stats.qualityGate.reasons.length <= 3,
      String(bm.stats.qualityGate?.reasons?.[0] || '').slice(0, 80))
    ok('乱码夹具：suggestion 给可复制的 OCR 参数串',
      /ocr="?always"?/.test(String(bm.stats.suggestion?.copy)) && /ocrEngine="?local"?/.test(String(bm.stats.suggestion?.copy)),
      String(bm.stats.suggestion?.copy || '').slice(0, 100))
    ok('乱码夹具：meta 纯 CPU（不触发渲染/OCR，无 ocrPages）', bm.stats.ocrPages === undefined)

    // ① 无 .ocr.md → convert 必须拒绝，且绝不产出目标文件
    const refused = join(OUT, 'broken-refused.txt')
    rmSync(refused, { force: true })
    let refusal = ''
    try { await call('office_convert', { source: broken, target: refused }) } catch (e) { refusal = e.message }
    ok('乱码夹具：convert 无缓存 → 拒绝且四要素齐备',
      ['页码=', '格式=', '根因=', '下一步='].every(k => refusal.includes(k)), refusal.slice(0, 120))
    ok('乱码夹具：拒绝时绝不产出目标文件', !existsSync(refused), existsSync(refused) ? '目标文件被写出来了' : '未产出')
    // DSH 补充（任务一-2 / F2 F3）：拒绝文本必须同时给出 sidecar 路径与 covered 现状
    ok('sidecar提示：convert 零缓存拒绝文本给出 sidecar 路径', refusal.includes('sidecar='),
      (refusal.match(/sidecar=[^(（\n]*/) || ['<无>'])[0].slice(0, 110))
    ok('sidecar提示：convert 零缓存拒绝文本明说 covered: 无', /covered: 无/.test(refusal),
      (refusal.match(/covered: [^)）]*/) || ['<无>'])[0])
    ok('sidecar提示：convert 拒绝文本给出可照抄的 grep 提示', /grep 'covered:'/.test(refusal))

    if (!eng) {
      ok('乱码夹具：本地引擎缺失 —— 缓存/重试段未执行', true, '引擎未安装')
    } else if (process.env.DSH_OFFICE_TEST_SKIP_BROKEN_OCR) {
      ok('乱码夹具：DSH_OFFICE_TEST_SKIP_BROKEN_OCR 已设 —— OCR 慢用例跳过', true)
    } else {
      // ② 分批跑满 OCR（单次最多 20 页）→ 有完整 .ocr.md 时 convert 才能转出
      await call('office_read', { path: broken, pages: '1-20', ocr: 'always', ocrEngine: 'local' })
      const cold = await call('office_read', { path: broken, pages: '21-40', ocr: 'always', ocrEngine: 'local' })
      const covered = new Set([...readFileSync(side, 'utf8').matchAll(/^## 第 (\d+) 页/gm)].map(m => Number(m[1])))
      const garbledList = String(bm.stats.garbledPages).split(',').flatMap(part => {
        const m = /^(\d+)\s*[-~]\s*(\d+)$/.exec(part.trim())
        if (m) return Array.from({ length: Number(m[2]) - Number(m[1]) + 1 }, (_, i) => Number(m[1]) + i)
        return part.trim() ? [Number(part)] : []
      })
      ok('乱码夹具：OCR 缓存已覆盖全部乱码页', garbledList.length > 0 && garbledList.every(p => covered.has(p)),
        `缓存 ${covered.size} 页 / 乱码页 ${garbledList.join(',')}`)
      // 验收 §3.4：必须能"正面看到无缺页"，而不是靠没报错推断
      ok('乱码夹具：冷启动分批跑完无缺页（stats.ocrFailedPages 为空数组）',
        Array.isArray(cold.stats.ocrFailedPages) && cold.stats.ocrFailedPages.length === 0,
        JSON.stringify(cold.stats.ocrFailedPages))
      ok('乱码夹具：正文无「批内缺页」脚注（无缺页就不该出现）', !/批内缺页/.test(String(cold.content)))

      // ---- 第八轮（自 WB 侧 P1-2 回移）：单次读取的 OCR 页数上限必须记账 ----
      // 走上面已经铺满的 sidecar 缓存（1-20 命中缓存 → 零渲染零引擎）。
      // ⚠ 历史判断（**第八轮记录；第九轮 R9-3 已修**）：21-30 页保留乱码文本层，整篇会被质量门
      //   判"不可读"而走 sidecar 降级；降级 notice/content **曾**覆盖 OCR 记账脚注 —— 所以这里只验
      //   "stats 记账不丢 + 降级 notice 仍给出下一步命令"；脚注文案通道由上面 26 页有文本层夹具单独覆盖（见 test.mjs:806）。
      //   （第九轮 R9-3 起，降级路径会从 stats 重建同一批脚注并追加进 notice/content，
      //    见紧邻本段的 `R9-3:` 两条断言。原文保留修订轨迹。）

      const cappedRead = await call('office_read', { path: broken, pages: '1-30', ocr: 'always', ocrEngine: 'local' })
      ok('页数上限：stats.ocrPagesCapped 记录 limit/requested/applied',
        cappedRead.stats.ocrPagesCapped?.limit === 20 && cappedRead.stats.ocrPagesCapped?.requested === 30
        && cappedRead.stats.ocrPagesCapped?.applied === 20,
        JSON.stringify(cappedRead.stats.ocrPagesCapped))
      ok('页数上限：被砍掉的页按页码数组逐页列出（21..30，不是区间串）',
        JSON.stringify(cappedRead.stats.ocrPagesCapped?.skipped) === JSON.stringify(Array.from({ length: 10 }, (_, i) => 21 + i)),
        JSON.stringify(cappedRead.stats.ocrPagesCapped?.skipped))
      ok('页数上限：正文降级到 sidecar 时记账不丢（stats 仍给上限、notice 仍给下一步）',
        cappedRead.stats.fallback === 'sidecar' && cappedRead.stats.ocrPagesCapped?.applied === 20
        && /正文质量门判定为不可读/.test(String(cappedRead.notice))
        && /ocr="always" ocrEngine="local"/.test(String(cappedRead.notice)),
        `fallback=${cappedRead.stats.fallback} · ${(String(cappedRead.notice).match(/正文质量门判定为不可读[^；]*/) || ['<无>'])[0].slice(0, 90)}`)

      ok('页数上限：stats.ocrPreview 不因显式 pages 出现（两种记账互斥）',
        cappedRead.stats.ocrPreview === undefined, JSON.stringify(cappedRead.stats.ocrPreview))

      // ---- R9-3（第九轮新增，splice 自 stage\w3-asserts.mjs）----
      // 病根：finishRead 降级分支只保留**头部 800 字符**摘录（cutAtBoundary(out.content, 800)），
      // 而「单次读取最多识别 N 页」脚注挂在正文**尾部**（notes 汇聚点）→ 旧实现下脚注整条随截断消失，
      // 调用方在 content/notice 两个通道都看不到"哪几页被上限砍掉"（第八轮 P1 第一版断言就是被它咬到才降级的）。
      // 本两条钉 content/notice 两通道；stats 字段与降级记账已由上面 L1839-1850 既有断言覆盖，不复读。
      ok('R9-3: 降级不丢脚注·content 通道',
        cappedRead.fallback === 'sidecar'
        && /单次读取最多识别 20 页：本次要求 30 页，第 21-30 页未做 OCR（仍用原文本层）/.test(String(cappedRead.content))
        && /续读：office_read path="[^"]+" ocr="always" ocrEngine="local" pages="21-30"/.test(String(cappedRead.content)),
        `fallback=${cappedRead.fallback} · content ${String(cappedRead.content).length} 字符 · 脚注出现 ${(String(cappedRead.content).match(/单次读取最多识别 20 页/g) || []).length} 次`)

      ok('R9-3: 降级不丢脚注·两类说明共存',
        /正文质量门判定为不可读/.test(String(cappedRead.notice))
        && /单次读取最多识别 20 页/.test(String(cappedRead.notice))
        && String(cappedRead.notice).indexOf('正文质量门判定为不可读') < String(cappedRead.notice).indexOf('单次读取最多识别'),
        `notice ${String(cappedRead.notice).length} 字符 · sidecar说明@${String(cappedRead.notice).indexOf('正文质量门判定为不可读')} < 上限脚注@${String(cappedRead.notice).indexOf('单次读取最多识别')}`)

      const convOut = join(OUT, 'broken-converted.txt')
      rmSync(convOut, { force: true })
      const cv = await call('office_convert', { source: broken, target: convOut })
      const diskTxt = readFileSync(convOut, 'utf8')
      ok('乱码夹具：有完整 .ocr.md → 转出成功且 fallback=ocr',
        existsSync(convOut) && cv.stats?.fallback === 'ocr' && String(cv.stats.garbledPages).length > 0,
        `fallback=${cv.stats?.fallback} garbledPages=${JSON.stringify(cv.stats?.garbledPages)}`)
      ok('乱码夹具：转出内容确实来自 OCR（无 NUL / 控制符垃圾）',
        !diskTxt.includes('\u0000') && scanBadCodePoints(diskTxt).bad === 0 && /Fixture page 17/.test(diskTxt),
        `${statSync(convOut).size} 字节`)
      ok('乱码夹具：notice 说明哪几页被替代', /用本地 OCR 文本替代/.test(String(cv.notice)), String(cv.notice || '').slice(0, 90))
      // —— 任务 D：缓存迁移的版本戳 ——
      ok('缓存迁移：sidecar manifest 写入当前 parser 版本戳',
        new RegExp(`parser:\\s*${PDF_PARSER_VERSION}\\b`).test(readFileSync(side, 'utf8').split('\n')[1] || ''),
        readFileSync(side, 'utf8').split('\n')[1]?.slice(0, 150) || '')

      // ③ 删掉 .ocr.md → 冷启动重识别：夹具本身干净，必须既无缺页也无"救回来的页"
      rmSync(side, { force: true })
      const re = await call('office_read', { path: broken, pages: '1-20', ocr: 'always', ocrEngine: 'local' })
      ok('乱码夹具：删缓存后冷启动重识别，本批无缺页',
        Array.isArray(re.stats.ocrFailedPages) && re.stats.ocrFailedPages.length === 0,
        JSON.stringify(re.stats.ocrFailedPages))
      ok('乱码夹具：重识别出的页有真实正文（不是空壳）',
        (ocrPageBodies(re.content).get(17) || '').includes('Fixture page 17'),
        (ocrPageBodies(re.content).get(17) || '').slice(0, 60))
      ok('乱码夹具：重识别只覆盖本批 pages（1-20）',
        Array.isArray(re.stats.ocrPages) && re.stats.ocrPages.length === 20 && re.stats.ocrPages[0] === 1,
        JSON.stringify(re.stats.ocrPages))

      // ---- DSH 补充（任务一-2）：部分覆盖时的拒绝文本：已覆盖到哪、还缺哪几页、下一步只给缺页 ----
      let partial = ''
      const partialOut = join(OUT, 'broken-partial.txt')
      rmSync(partialOut, { force: true })
      try { await call('office_convert', { source: broken, target: partialOut }) } catch (e) { partial = String(e.message) }
      const rng = s => String(s).split(',').flatMap(part => {
        const m = /^(\d+)\s*[-~]\s*(\d+)$/.exec(part.trim())
        return m ? Array.from({ length: Number(m[2]) - Number(m[1]) + 1 }, (_, i) => Number(m[1]) + i)
          : (part.trim() ? [Number(part)] : [])
      }).filter(Number.isFinite)
      const covM = /covered: ([\d,\-~]+)\s*\//.exec(partial)
      const nextM = /ocrEngine="local" pages="([\d,\-~]+)"/.exec(partial)
      const coveredSet = new Set(covM ? rng(covM[1]) : [])
      const wantNext = nextM ? rng(nextM[1]) : []
      ok('sidecar提示：部分缓存 → 拒绝文本报出 covered 区间', !!covM && coveredSet.size > 0,
        covM ? `covered: ${covM[1]}` : partial.slice(0, 130))
      ok('sidecar提示：部分缓存 → 拒绝文本点出仍缺的页', /仍缺第 [\d,\-~]+ 页/.test(partial),
        (partial.match(/仍缺第 [\d,\-~]+ 页/) || ['<无>'])[0])
      ok('sidecar提示：下一步的 pages 只列缺页（不叫已覆盖的页重跑）',
        wantNext.length > 0 && wantNext.every(p => !coveredSet.has(p)),
        nextM ? `pages="${nextM[1]}" 而 covered="${covM ? covM[1] : '?'}"` : '<无下一步串>')
      ok('sidecar提示：部分缓存下 convert 仍然绝不产出目标文件', !existsSync(partialOut))
    }
  }

  // ===========================================================================
  // 第三轮：①（ObjStm 病灶样本）修好之后的**行为翻转** + 页序/缓存迁移 + 断行/填空
  // 前缀：解析器： / 断行： / 填空： / 缓存迁移：
  // ===========================================================================
  {
    const s1Src = process.env.DSH_OFFICE_TEST_BROKEN_PDF
      || '<samples>'
    if (!existsSync(s1Src)) {
      ok('解析器：①夹具缺失 —— 未执行（用 DSH_OFFICE_TEST_BROKEN_PDF 指定）', true, s1Src)
    } else {
      const broken = join(OUT, 'broken-layer.pdf')
      await copyFile(s1Src, broken)
      rmSync(ocrCachePath(broken), { force: true })

      const bm = await call('office_read', { path: broken, as: 'meta' })
      // ⚠ 本轮最重要的行为翻转：修 off-by-First 之前 textLayerUsable===false
      ok('解析器：① 修后 meta → textLayerUsable=true（本轮最重要的行为翻转）',
        bm.stats.textLayerUsable === true, `usable=${bm.stats.textLayerUsable}`)
      ok('解析器：① 修后 meta → garbledPages=[]（整本不再命中质量门）',
        JSON.stringify(bm.stats.garbledPages) === '[]', JSON.stringify(bm.stats.garbledPages))
      ok('解析器：① 修后 qualityGate.reasons 为空（不再有"私用区/控制符"指控）',
        Array.isArray(bm.stats.qualityGate?.reasons) && bm.stats.qualityGate.reasons.length === 0,
        JSON.stringify(bm.stats.qualityGate?.reasons))
      ok('解析器：① 无文字页仍是 5 页（页序修正后为 1,2,3,5,35）',
        JSON.stringify(bm.stats.scannedPages || []) === '[1,2,3,5,35]', JSON.stringify(bm.stats.scannedPages))

      // ① 修后正文：字符层与填空层
      const md = await call('office_read', { path: broken, as: 'markdown', ocr: 'never', limit: 200000 })
      const body = String(md.content)
      const cjk = (body.match(/[\u4e00-\u9fff]/g) || []).length
      const ctrl = (body.match(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g) || []).length
      const blanks = (body.match(/_{3,}/g) || []).length
      ok('解析器：① 修后 CJK > 16,000（旧版为 0）', cjk > 16000, `${cjk} 个汉字`)
      ok('解析器：① 修后控制符 = 0（旧版 19,000+）', ctrl === 0, `${ctrl} 个控制符`)
      ok('解析器：① 修后正文含成片填空下划线', /________/.test(body) && blanks === 305, `${blanks} 段填空`)
      ok('填空：① 填空段数 = 305（与第二节实测一致，改动导致变化必须解释）', blanks === 305, `${blanks}`)
      ok('填空：① 含填空页 = 29 页', (body.match(/第 \d+ 页/g) || []).length >= 0
        && new Set([...body.matchAll(/(?:^|\n)(?:<!-- 第 (\d+) 页 -->)\n([\s\S]*?)(?=\n<!-- 第 |\n\n> |\s*$)/g)]
          .filter(m => /_{3,}/.test(m[2])).map(m => m[1])).size === 29,
      '含填空的页数')
      ok('填空：① 第 6 页能直接读出"习近平新时代中国特色社会主义思想的________"',
        /习近平新时代中国特色社会主义思想的_+/.test(body.replace(/\n/g, '')),
        (body.match(/[^\n]*习近平新时代[^\n]*/) || ['<无>'])[0].slice(0, 70))

      // ② 干净样本的"字符不变 + 结构变好"
      const s2Src = process.env.DSH_OFFICE_TEST_CID_PDF || ''
      if (!existsSync(s2Src)) {
        ok('断行：②夹具缺失 —— 未执行（用 DSH_OFFICE_TEST_CID_PDF 指定）', true, s2Src)
      } else {
        const clean = join(OUT, 'clean-cid.pdf')
        await copyFile(s2Src, clean)
        const raw = readPdfText(readFileSync(clean))
        const all = raw.sections.map(s => s.text).join('\n')
        const pageLines = raw.sections.map(s => s.text.split('\n').length)
        ok('断行：② 每页换行数 > 0（旧版 43 页全是"每页一行"）',
          pageLines.every(n => n > 1), `最少 ${Math.min(...pageLines)} 行`)
        ok('断行：② 换行总数从 0 变成正常量级（>1000）',
          (all.match(/\n/g) || []).length > 1000, `${(all.match(/\n/g) || []).length} 个换行`)
        ok('断行：② CJK 逐字不变（19,879）', (all.match(/[\u4e00-\u9fff]/g) || []).length === 19879,
          `${(all.match(/[\u4e00-\u9fff]/g) || []).length}`)
        ok('断行：② 填空段数不变（700 段 / 43 页）',
          (all.match(/_{3,}/g) || []).length === 700 && raw.sections.filter(s => /_{3,}/.test(s.text)).length === 43,
          `${(all.match(/_{3,}/g) || []).length} 段`)
        // 旧版 ② 的基线：本轮用**备份的旧 `pdf.js`** 重新实测得出，不再用魔法常量 ——
        //   旧：total 29,442 / 换行 42 → **非换行 29,400**
        //   新：total 31,200 / 换行 1,780 → 非换行 29,420
        // 差额 **+20** 全是**行分隔 + markdown 标记**（`# ` 行 7 条）；扣掉标记与换行后只多 6 个字符
        // （`# ` 之后的空格与 run 间确有空隙时补的空格）。留 8 的容差是因为这类空格数随排版微调。
        const OLD_NO_NL = 29400
        const VISIBLE_SLACK = 8
        const noNl = all.replace(/\n/g, '').length
        const heads = (all.match(/^# /gm) || []).length
        ok('断行：② 可见字符数只差在换行与标题标记（正文汉字一个不多一个不少）',
          noNl - heads * 2 <= OLD_NO_NL + VISIBLE_SLACK,
          `非换行 ${noNl}（去 ${heads} 个 "# " 标记后 ${noNl - heads * 2}）vs 旧实测 ${OLD_NO_NL}（容差 ${VISIBLE_SLACK}）`)
      }

      // —— 页序与缓存迁移（任务 D）——
      // 修 off-by-First 之前，① 的 /Pages 主体读成了 ObjStm 的头部索引表 → `/Kids` 解析不出
      // → `pages()` 落到 line 202 的兜底"扫描全部 /Type/Page"，页序变成**对象号升序**：
      // 封面（obj 432）从第 1 页跑到第 34 页。而 OCR 走 WinRT 的物理页序 → 两者错位。
      if (eng && !process.env.DSH_OFFICE_TEST_SKIP_BROKEN_OCR) {
        rmSync(ocrCachePath(broken), { force: true })
        const re = await call('office_read', { path: broken, pages: '15-20', ocr: 'always', ocrEngine: 'local' })
        ok('缓存迁移：第 17 页原生倍率不过质量门，已自动换倍率救回（不再丢页）',
          /(^|,)17(,|$|-)/.test(String(re.stats.ocrRetried || '')), `ocrRetried=${re.stats.ocrRetried}`)
        ok('缓存迁移：stats.ocrRetryScale 给出第 17 页的倍率',
          Number(re.stats.ocrRetryScale?.['17']) >= 1, `scale=${JSON.stringify(re.stats.ocrRetryScale || {})}`)
        ok('缓存迁移：救回后本批无缺页（stats.ocrFailedPages 空数组）',
          Array.isArray(re.stats.ocrFailedPages) && re.stats.ocrFailedPages.length === 0,
          JSON.stringify(re.stats.ocrFailedPages))
        ok('缓存迁移：第 17 页有真实正文（不是空壳）',
          ((ocrPageBodies(re.content).get(17) || '').match(/[一-龥]/g) || []).length >= 50,
          `${((ocrPageBodies(re.content).get(17) || '').match(/[一-龥]/g) || []).length} 个汉字`)
        const head = readFileSync(ocrCachePath(broken), 'utf8').split('\n').slice(0, 2).join('\n')
        ok('缓存迁移：sidecar manifest 同时记 retry: 17=… 与当前 parser 版本戳',
          /retry: [\d=,.]*17=\d/.test(head) && /src: rapidocr=/.test(head)
          && new RegExp(`parser:\\s*${PDF_PARSER_VERSION}\\b`).test(head),
          head.split('\n')[1]?.slice(0, 170) || '')
      } else if (eng) {
        ok('缓存迁移：DSH_OFFICE_TEST_SKIP_BROKEN_OCR 已设 —— 页序/重试段跳过', true)
      } else {
        ok('缓存迁移：本地引擎缺失 —— 页序/重试段跳过', true)
      }

      // —— 迁移协议本身：旧戳/无戳 sidecar 一律视为"未覆盖" ——
      {
        const stale = join(OUT, 'stale-probe.pdf')
        await copyFile(s1Src, stale)
        const staleSide = ocrCachePath(stale)
        // 模拟"上一代解析器写下的 sidecar"：有页正文、但没有 parser 戳
        writeFileSync(staleSide, '<!-- dsh-office OCR cache -->\n<!-- covered: 1-35 | total: 35 -->\n# OCR\n\n'
          + '## 第 1 页（OCR）\n上一代解析器写下的第 1 页文本\n', 'utf8')
        const staleMeta = await call('office_read', { path: stale, as: 'meta' })
        ok('缓存迁移：meta 报出旧 sidecar 已作废（ocrCacheStale + 原因 + 当前版本）',
          staleMeta.stats.ocrCacheStale?.path === staleSide
          && Number(staleMeta.stats.ocrCacheStale?.current) === PDF_PARSER_VERSION
          && /已作废旧 OCR 缓存/.test(String(staleMeta.stats.ocrCacheNote)),
          JSON.stringify(staleMeta.stats.ocrCacheStale || null))
        ok('缓存迁移：作废的 sidecar 不再被当作 covered（绝不静默复用错页文本）',
          !/covered: 1-35/.test(sidecarCoverage(stale)) && /已作废/.test(sidecarCoverage(stale)),
          sidecarCoverage(stale).slice(0, 120))
        // 版本戳正确**且身份相符**时照常命中（任务一：身份是 parser 之外的第二个条件）
        writeOcrSidecar(stale, `covered: 1 | total: 35 | parser: ${PDF_PARSER_VERSION}`,
          '# OCR\n\n## 第 1 页（OCR）\n当前代解析器写下的第 1 页文本\n')
        const freshMeta = await call('office_read', { path: stale, as: 'meta' })
        ok('缓存迁移：parser 版本相符的 sidecar 正常命中（不误伤）',
          freshMeta.stats.ocrCacheStale === undefined && /covered: 1\b/.test(sidecarCoverage(stale)),
          sidecarCoverage(stale).slice(0, 120))
        // 有乱码页 + 只有旧戳缓存 → convert 拒绝时要把"为什么缓存像是空的"说清楚。
        // ⚠ 这里必须换成**自造乱码夹具**：样本①修好之后文字层完全可用，convert 根本不会拒绝，
        // 拿它当"乱码源"这个前提已经消失（第三轮：① 从"必须 OCR"降级为可读）。
        const staleRefuse = join(OUT, 'stale-refuse.pdf')
        writeFileSync(staleRefuse, buildGarbledFixture({ pages: 3 }))
        const staleRefuseSide = ocrCachePath(staleRefuse)
        writeFileSync(staleRefuseSide, '<!-- dsh-office OCR cache -->\n<!-- covered: 1-3 | total: 3 -->\n# OCR\n\n'
          + '## 第 1 页（OCR）\n上一代解析器写下的第 1 页文本\n', 'utf8')
        let staleRefusal = ''
        try { await call('office_convert', { source: staleRefuse, target: join(OUT, 'stale-out.txt') }) } catch (e) { staleRefusal = String(e.message) }
        ok('缓存迁移：convert 拒绝文本解释"缓存为何不算数"（不是静默当没缓存）',
          /已作废旧 OCR 缓存/.test(staleRefusal) && new RegExp(`parser:\\s*${PDF_PARSER_VERSION}`).test(staleRefusal),
          staleRefusal.slice(0, 160))
        rmSync(staleRefuseSide, { force: true })
        rmSync(join(OUT, 'stale-out.txt'), { force: true })
        rmSync(staleSide, { force: true })
      }
    }
  }

  // ---- 任务 H：填空呈现（opt-in，**只改呈现**，不 gate 任何正确性修复）----
  {
    const blankPdfPath = join(OUT, 'blank-view.pdf')
    writeFileSync(blankPdfPath, writePdf({
      kind: 'document',
      meta: { title: 'blanks' },
      blocks: [{ type: 'paragraph', runs: [{ text: '填空：________ 与 __________。' }] }],
    }, {}))
    const off = await call('office_read', { path: blankPdfPath, as: 'markdown' })
    const prevEnv = process.env.DSH_OFFICE_BLANKS
    process.env.DSH_OFFICE_BLANKS = '1'
    const on = await call('office_read', { path: blankPdfPath, as: 'markdown' })
    const onJson = await call('office_read', { path: blankPdfPath, as: 'json' })
    const onMeta = await call('office_read', { path: blankPdfPath, as: 'meta' })
    if (prevEnv === undefined) delete process.env.DSH_OFFICE_BLANKS
    else process.env.DSH_OFFICE_BLANKS = prevEnv
    const offAgain = await call('office_read', { path: blankPdfPath, as: 'markdown' })
    ok('填空：DSH_OFFICE_BLANKS 未设时逐字不变（呈现层绝不 gate 正确性）',
      off.content === offAgain.content && /_{3,}/.test(off.content) && !/class="blank"/.test(off.content))
    ok('填空：DSH_OFFICE_BLANKS=1 → markdown 渲成 <span class="blank">', /<span class="blank">＿+<\/span>/.test(on.content),
      (on.content.match(/<span class="blank">[^<]*<\/span>/) || ['<无>'])[0])
    ok('填空：DSH_OFFICE_BLANKS=1 → json 分支同样生效且仍是合法 JSON',
      /class=\\"blank\\"/.test(onJson.content) && (() => { try { JSON.parse(onJson.content); return true } catch { return false } })())
    ok('填空：DSH_OFFICE_BLANKS=1 → as="meta" 不受影响（不经过呈现层）',
      onMeta.stats && !/class="blank"/.test(String(onMeta.stats.textLayerUsable)))
  }

  // ===========================================================================
  // DSH 专属补丁（第二轮）的账：P0 接线 / F1 矩阵 / 失败通道可观测性 / sidecar 全文
  // 前缀：接线： / F1矩阵： / sidecar提示：
  // ===========================================================================

  // ---- 接线：出站边界的三个调用点（历史教训：上一轮栽在"插了代码没接线"） ----
  ok('接线：defineToolLite 的 execute / render / catch 三个出口都挂上了边界', (() => {
    const src = readFileSync(join(HERE, 'index.js'), 'utf8')
    const from = src.indexOf('function defineToolLite(')
    const to = src.indexOf('// path / io helpers')
    if (from < 0 || to < 0 || to < from) return false
    const body = src.slice(from, to)
    return /finalizeToolValue\(await options\.execute/.test(body)
      && /throw sanitizeThrown\(/.test(body)
      && /finalizeToolValue\(options\.render/.test(body)
  })())
  ok('接线：finalizeToolValue 与 sanitizeThrown 除定义处外确有调用点', (() => {
    const src = readFileSync(join(HERE, 'index.js'), 'utf8')
    const callSites = n => src.split('\n')
      .filter(l => l.includes(`${n}(`) && !new RegExp(`(export )?function ${n}\\(`).test(l)).length
    // finalizeToolValue：execute 出口 + render 投影（≥2）；sanitizeThrown：catch 分支（≥1）
    return callSites('finalizeToolValue') >= 2 && callSites('sanitizeThrown') >= 1
  })(), `finalizeToolValue=${(() => {
    const src = readFileSync(join(HERE, 'index.js'), 'utf8')
    return src.split('\n').filter(l => l.includes('finalizeToolValue(') && !/^\s*(export )?function finalizeToolValue\(/.test(l)).length
  })()} sanitizeThrown=${(() => {
    const src = readFileSync(join(HERE, 'index.js'), 'utf8')
    return src.split('\n').filter(l => l.includes('sanitizeThrown(') && !/^\s*(export )?function sanitizeThrown\(/.test(l)).length
  })()}`)
  {
    const w1 = await call('office_read', { path: join(OUT, 'report.pdf'), pages: '1' })
    ok('接线：真实调用确实过了边界（stats.sanitized 只有 finalizeToolValue 会注入）',
      typeof w1.stats?.sanitized === 'number' && losslessJsonProblem(w1) === null, `sanitized=${w1.stats?.sanitized}`)
  }
  {
    // 错误文本必须经 sanitizeThrown：文件名里的游离代理不得原样出现在 message 里
    let msg = ''
    try { await call('office_read', { path: join(OUT, '坏名字-\ud800.pdf') }) } catch (e) { msg = String(e && e.message) }
    ok('接线：catch 分支过了 sanitizeThrown（错误文本无游离代理 / 无控制符）',
      msg.length > 0 && scanBadCodePoints(msg).bad === 0, msg ? `len=${msg.length} bad=${scanBadCodePoints(msg).bad}` : '<未抛错>')
  }

  // ---- F1 矩阵：修复后 a–f 六条必须全部能返回（DSH 侧最硬的回归清单） ----
  {
    const f1Pdf = [join(OUT, 'broken-layer.pdf'), join(OUT, 'scan-sample.pdf'), join(OUT, 'vision-pages.pdf'), join(OUT, 'report.pdf')]
      .find(p => existsSync(p))
    if (!f1Pdf) {
      ok('F1矩阵：夹具缺失 —— 未执行', true, '产物目录里没有可用的 PDF')
    } else {
      const f1meta = await call('office_read', { path: f1Pdf, as: 'meta' })
      const pg = f1meta.stats.pages || 1
      const one = String(Math.min(3, pg))
      // §3.3 的 f 行 = 病灶样本里"这一页没有文字层、但整本有"的那种页：应当**返回**一段干净的说明
      const mixScanned = (f1meta.stats.scannedPages || []).find(p => p >= 1)
      const rows = [
        ['a1 as=markdown 整本 limit=200000', { path: f1Pdf, as: 'markdown', limit: 200000 }],
        ['a2 as=markdown 整本 limit=20000', { path: f1Pdf, as: 'markdown', limit: 20000 }],
        ['a3 as=markdown 整本 limit=5000', { path: f1Pdf, as: 'markdown', limit: 5000 }],
        ['b  as=text pages 单页 ocr=never', { path: f1Pdf, as: 'text', pages: one, ocr: 'never' }],
        ['c  as=markdown pages 单页 ocr=never', { path: f1Pdf, as: 'markdown', pages: one, ocr: 'never' }],
        ['d  as=text pages 单页 ocr=always local', { path: f1Pdf, as: 'text', pages: one, ocr: 'always', ocrEngine: 'local' }],
        ['e  as=text pages 单页 ocr=always vision', { path: f1Pdf, as: 'text', pages: one, ocr: 'always', ocrEngine: 'vision' }],
      ]
      if (mixScanned && f1meta.stats.pagesWithText > 0) {
        rows.push(['f1 混合书中纯扫描页 ocr=never', { path: f1Pdf, as: 'text', pages: String(mixScanned), ocr: 'never' }])
        rows.push(['f2 混合书中纯扫描页 ocr=always local', { path: f1Pdf, as: 'text', pages: String(mixScanned), ocr: 'always', ocrEngine: 'local' }])
        rows.push(['f3 混合书中纯扫描页 ocr=always vision', { path: f1Pdf, as: 'text', pages: String(mixScanned), ocr: 'always', ocrEngine: 'vision' }])
      }
      for (const [label, args] of rows) {
        if (!existsSync(args.path)) { ok(`F1矩阵：${label} 夹具缺失未执行`, true, args.path); continue }
        let res = null, err = ''
        try { res = await call('office_read', args) } catch (e) { err = String(e && e.message).slice(0, 90) }
        ok(`F1矩阵：${label} 正常返回且无损`, !!res && !err && losslessJsonProblem(res) === null,
          err || (res ? `${String(res.content || '').length} 字符 / fallback=${res.stats?.fallback}` : '<无返回>'))
      }
      if (!mixScanned || !f1meta.stats.pagesWithText) {
        ok('F1矩阵：f 行未执行（该夹具整本都是扫描页，改由下一条断言覆盖）', true, String(f1Pdf))
      } else {
        const f1none = await call('office_read', { path: f1Pdf, as: 'text', pages: String(mixScanned), ocr: 'never' })
        ok('F1矩阵：纯扫描页 ocr=never 返回干净的"无文本层"说明（不是错误、不是乱码）',
          /未提取到文本层/.test(f1none.content) && !/批内缺页|sidecar=/.test(f1none.content) && !/乱码/.test(f1none.content),
          f1none.content.slice(0, 70))
      }
      // 整本都没有文字层 + ocr=never：按设计**抛**四要素（不是裸 invalid output），这条把这个契约钉住
      {
        let allScanErr = ''
        try { await call('office_read', { path: join(OUT, 'vision-pages.pdf'), ocr: 'never' }) } catch (e) { allScanErr = String(e.message) }
        ok('F1矩阵：整本无文字层 + ocr=never → 抛的是带四要素的说明（非裸 invalid output）',
          allScanErr.includes('【读取失败｜四要素】') && /下一步=/.test(allScanErr) && scanBadCodePoints(allScanErr).bad === 0,
          allScanErr.slice(0, 130))
      }
      rmSync(join(OUT, 'vision-pages.ocr.md'), { force: true })
    }
  }

  // ---- sidecar提示：失败通道的可观测性（任务六-1 / F2 F3） ----
  ok('sidecar提示：非 PDF / 无 sidecar 时探测返回空串（绝不误报）',
    sidecarCoverage(join(OUT, 'report.docx')) === '' && sidecarCoverage(join(OUT, 'no-such-file.pdf')) === '')
  {
    // 造一个"无文本层 PDF + 已有部分 OCR 缓存"的组合，触发抛错分支看它是否报出盘上成果
    const hintPdf = join(OUT, 'hint-probe.pdf')
    await call('office_create', {
      path: hintPdf,
      document: { blocks: Array.from({ length: 5 }, (_, i) => ({ type: i % 2 ? 'pagebreak' : 'hr' })) },
    })
    const hintSide = writeOcrSidecar(hintPdf, `covered: 1-2 | total: 3 | parser: ${PDF_PARSER_VERSION}`,
      '# OCR\n\n## 第 1 页（OCR）\n第一页识别文本\n\n## 第 2 页（OCR）\n第二页识别文本\n')
    let hintMsg = ''
    try { await call('office_read', { path: hintPdf, ocr: 'never' }) } catch (e) { hintMsg = String(e && e.message) }
    ok('sidecar提示：抛错分支带出 sidecar 路径 + covered + 缺页（成果在盘上看得见）',
      hintMsg.includes('sidecar=') && /covered: 1-2/.test(hintMsg) && /缺页=3/.test(hintMsg) && /grep 'covered:'/.test(hintMsg),
      hintMsg.split('\n').slice(-1)[0]?.slice(0, 170) || '<未抛错>')
    ok('sidecar提示：四要素仍然齐备（新增信息不挤掉旧契约）',
      ['页码=', '格式=', '根因=', '下一步='].every(k => hintMsg.includes(k)), hintMsg.slice(0, 100))
    rmSync(hintSide, { force: true })
  }

  // ---- sidecar提示：正文不可读时，sidecar 落的是**全文**而不是被护栏截过的先缀 ----
  {
    const bigJunk = join(OUT, 'big-junk.md')
    const full = `# 乱码样本\n\n${'\uE0A1\uE0A2\uE0A3'.repeat(500)}`
    writeFileSync(bigJunk, full, 'utf8')
    const cut = await call('office_read', { path: bigJunk, limit: 300 })
    const sidePath = cut.stats?.sidecar || readSidecarPath(bigJunk)
    const sideText = existsSync(sidePath) ? readFileSync(sidePath, 'utf8') : ''
    ok('sidecar提示：内联截断时 sidecar 仍拿到全文（sidecarChars > content 长度）',
      cut.truncated === true && Number(cut.stats?.sidecarChars) > 300 && sideText.length > cut.content.length,
      `sidecarChars=${cut.stats?.sidecarChars} content=${cut.content.length} sidecar文件=${sideText.length} 字符`)
    ok('sidecar提示：sidecar 正文含护栏截掉的那部分（不重不漏，末尾完整）',
      sideText.includes('\uE0A1\uE0A2\uE0A3'.repeat(3)) && !sideText.includes('__fullBody'),
      `尾部 ${JSON.stringify(sideText.slice(-12))}`)
    rmSync(sidePath, { force: true })
  }
  // ---- R10-1（第十轮新增）：PDF 分支降级 + 内联截断 → sidecar 必须落**整篇** ----
  // 病灶：readPdf 的 out 从不挂 `__fullBody`（非 PDF 分支挂了，见 index.js:3901），
  // 于是 finishRead 只能拿**被内联护栏截断的** content 写 sidecar，而 notice 承诺"整篇正文"。
  // 既有 big-junk 断言走非 PDF 分支，锁不住这条缺口 —— 这里用 PDF 夹具走 PDF 分支。
  // 双向自证（第十轮批写前实测，tc-bidir.mjs）：未修复态 sidecarChars=600=limit、sidecar 文件 782 < content 983；
  // 修复后 sidecarChars=8545 > limit、sidecar 文件 8713 > content 969 ⇒ 未修复必 FAIL、修复后必 PASS。
  {
    const R10_PDF_PAGES = 35            // 夹具页数
    const R10_READ_LIMIT = 600          // office_read limit（PDF 分支下界 100）
    const R10_READ_OCR = 'never'        // 只读文本层，不依赖本地引擎
    const R10_TAIL_MARK = 'R10TAILMARK' // 只出现在末页，护栏截断必然切掉它
    const r10Pdf = join(OUT, 'r10-pdf-sidecar.pdf')
    const r10Blocks = []
    for (let p = 1; p <= R10_PDF_PAGES; p++) {
      if (p > 1) r10Blocks.push({ type: 'pagebreak' })
      const last = p === R10_PDF_PAGES
      r10Blocks.push({ type: 'heading', level: 2, text: last ? `Fixture page ${p} ${R10_TAIL_MARK}` : `Fixture page ${p}` })
      r10Blocks.push({ type: 'paragraph', runs: [{ text: last
        ? `This is fixture page ${p} of the text layer gate. tail=${R10_TAIL_MARK}`
        : `This is fixture page ${p} of the text layer gate.` }] })
      r10Blocks.push({ type: 'paragraph', runs: [{ text: '\uE0A1\uE0A2\uE0A3'.repeat(30) }] })
    }
    writeFileSync(r10Pdf, writePdf({ kind: 'document', meta: { title: 'R10 pdf sidecar full-body fixture' }, blocks: r10Blocks }, {}))
    rmSync(readSidecarPath(r10Pdf), { force: true })
    rmSync(ocrCachePath(r10Pdf), { force: true })

    const r10 = await call('office_read', { path: r10Pdf, limit: R10_READ_LIMIT, ocr: R10_READ_OCR })
    const r10Side = r10.stats?.sidecar || readSidecarPath(r10Pdf)
    const r10Text = existsSync(r10Side) ? readFileSync(r10Side, 'utf8') : ''
    const r10Flat = r10Text.replace(/\s+/g, '')
    const r10TailSentence = `Thisisfixturepage${R10_PDF_PAGES}ofthetextlayergate.`
    const r10Pages = (r10Flat.match(/Fixturepage\d+/g) || []).length
    ok('R10-1: PDF 降级 + 内联截断时 sidecar 拿到整篇（sidecarChars 与 sidecar 文件都超过截断面）',
      r10.truncated === true && r10.stats?.fallback === 'sidecar'
      && Number(r10.stats?.sidecarChars) > R10_READ_LIMIT && r10Text.length > r10.content.length,
      `truncated=${r10.truncated} fallback=${r10.stats?.fallback} sidecarChars=${r10.stats?.sidecarChars} limit=${R10_READ_LIMIT} content=${r10.content.length} sidecar文件=${r10Text.length}`)
    ok('R10-2: sidecar 含被护栏截掉的末页独有标记（不重不漏），且返回值不泄漏 __fullBody',
      r10Flat.includes(R10_TAIL_MARK) && r10Flat.includes(r10TailSentence) && r10Pages >= R10_PDF_PAGES
      && !('__fullBody' in r10) && !JSON.stringify(r10).includes('__fullBody'),
      `尾标记=${r10Flat.includes(R10_TAIL_MARK)} 末页句=${r10Flat.includes(r10TailSentence)} 页数=${r10Pages}/${R10_PDF_PAGES} content=${r10.content.length} sidecar文件=${r10Text.length}`)
    rmSync(r10Side, { force: true })
    rmSync(r10Pdf, { force: true })
    rmSync(ocrCachePath(r10Pdf), { force: true })
  }

  // ---- sidecar合并：manifest 的 src / retry 必须**跨批累积**，不能被后加的段吃掉 ----
  {
    const mPdf = join(OUT, 'hint-probe.pdf')
    const mSide = ocrCachePath(mPdf)
    if (!eng) {
      ok('sidecar合并：本地引擎缺失 —— 未执行', true, '跳过')
    } else {
      writeOcrSidecar(mPdf, `covered: 1 | total: 3 | parser: ${PDF_PARSER_VERSION} | src: vision=1 | retry: 1=2`,
        '# OCR\n\n## 第 1 页（OCR）\n历史识别文本\n')
      await call('office_read', { path: mPdf, pages: '2', ocr: 'always', ocrEngine: 'local', limit: 200 })
      const head = readFileSync(mSide, 'utf8').split('\n').slice(0, 2).join('\n')
      const srcSeg = (/src: ([^|\n]*)/.exec(head) || [, ''])[1]
      const retrySeg = (/retry: ([^|\n]*)/.exec(head) || [, ''])[1]
      ok('sidecar合并：src 段不再被同行后面的 retry 段吃到（旧正则 [^\\n]+ 的坑）',
        srcSeg.includes('vision=1') && srcSeg.includes('rapidocr=2'), `src:"${srcSeg.trim()}"`)
      ok('sidecar合并：历史 retry 不被下一批写抹掉', retrySeg.includes('1=2'), `retry:"${retrySeg.trim()}"`)
      ok('sidecar合并：covered 合并推进（不回退）', /covered: 1-2/.test(head), (/covered: [^|\n]*/.exec(head) || ['<无>'])[0])
      // 历史 retry 记录只进 manifest，绝不能冒充成"本轮救回的页"（否则 stats 会谎报做了重试）
      const histHit = await call('office_read', { path: mPdf, pages: '1', ocr: 'always', ocrEngine: 'local', limit: 200 })
      ok('sidecar合并：历史 retry 只进 manifest，不冒充本轮 stats.ocrRetried',
        histHit.stats?.ocrRetried === undefined && histHit.stats?.ocrRetryScale === undefined
        && /OCR 缓存/.test(String(histHit.content)), `ocrRetried=${JSON.stringify(histHit.stats?.ocrRetried)}`)
      rmSync(mSide, { force: true })
    }
  }

  // ===========================================================================
  // 第四轮补丁（需求 1-5）：HTML 读取 / 批量 stats / 页级续读与批间衔接 /
  // 字符级质量启发 / CSV 编码收敛
  // ===========================================================================

  // ---------------- 需求 1：HTML 读取与转换 ----------------
  {
    const htmlFixture = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>需求1夹具 · HTML读取</title>
<style>.hidden{display:none} body{color:#333}</style>
<script>var evil = "<p>script 里的标签不是正文</p>";</script>
</head>
<body>
<h1>测试主标题</h1>
<p>第一段：A&amp;B 合并&#8212;破折号&#x201C;引号&#150;短横线&nbsp;空格 &lt;转义&gt;。</p>
<div class="card">
  <h2>第二章 嵌套列表</h2>
  <ul>
    <li>外层一
      <ol>
        <li>内层甲</li>
        <li>内层乙</li>
      </ol>
    </li>
    <li>外层二</li>
  </ul>
</div>
<table class="tb">
  <tr><th>列一</th><th>列二</th></tr>
  <tr><td>甲,含逗号</td><td>乙"含引号"</td></tr>
</table>
<blockquote>引用块内容：要保留语义顺序。</blockquote>
<pre>pre 里的
  原样空白</pre>
<p>图片：<img src="data:image/png;base64,AAAA" alt="示意"> 与 <a href="https://example.com/x?a=1&amp;b=2">外链</a>。</p>
<p style="display:none">这一段是隐藏的，不应出现</p>
<div>最后一段正文。</div>
</body>
</html>`
    const htmlFile = join(OUT, 'r1.html')
    writeFileSync(htmlFile, htmlFixture, 'utf8')

    const hm = await call('office_read', { path: htmlFile, as: 'meta' })
    ok('R1 meta：format=html 且带 document 统计', hm.format === 'html' && hm.stats.blocks > 0 && hm.stats.characters > 0, JSON.stringify(hm.stats))
    ok('R1 meta：textLayerUsable 正常给出', hm.stats.textLayerUsable === true, JSON.stringify(hm.stats.textLayerUsable))

    const hmd = await call('office_read', { path: htmlFile, as: 'markdown' })
    const hcontent = String(hmd.content)
    ok('R1 markdown：标题语义（h1/h2→#）', hcontent.includes('# 测试主标题') && hcontent.includes('## 第二章 嵌套列表'), hcontent.slice(0, 120))
    ok('R1 markdown：实体解码（named+数字+CP1252 别名）', hcontent.includes('A&B') && hcontent.includes('—') && hcontent.includes('“') && hcontent.includes('–'),
      hcontent.slice(hcontent.indexOf('第一段'), hcontent.indexOf('第一段') + 60))
    ok('R1 markdown：style/script 剥离', !hcontent.includes('display:none') && !hcontent.includes('var evil') && !hcontent.includes('script 里的标签'), '')
    ok('R1 markdown：嵌套列表保留层级（li→-，缩进表层级）', /\n {2}- 内层甲/.test(hcontent) && /\n- 外层二/.test(hcontent),
      hcontent.slice(hcontent.indexOf('外层一'), hcontent.indexOf('外层一') + 60))
    ok('R1 markdown：表格转管道表', hcontent.includes('| 列一 | 列二 |') && hcontent.includes('甲,含逗号') && hcontent.includes('乙"含引号"'), '')
    ok('R1 markdown：quote/pre 结构与语义顺序', hcontent.includes('> 引用块内容：要保留语义顺序。') && hcontent.includes('```\npre 里的\n  原样空白\n```'), '')
    ok('R1 markdown：外链保留、data: URI 丢弃、图前图后正文都保留', hcontent.includes('[外链](https://example.com/x?a=1&b=2)') && hcontent.includes('图片：') && hcontent.includes('与') && hcontent.includes('。'),
      hcontent.slice(hcontent.indexOf('图片'), (hcontent.indexOf('图片') || 0) + 90))
    ok('R1 markdown：隐藏子树不出现', !hcontent.includes('隐藏的'), '')

    const htxt = await call('office_read', { path: htmlFile, as: 'text' })
    const tcontent = String(htxt.content)
    ok('R1 text：无 markdown 残留标记（纯文本+制表符表格）', !tcontent.includes('## ') && !tcontent.includes('| 列一') && tcontent.includes('列一\t列二'), tcontent.slice(0, 120))
    ok('R1 text：正文与段落边界保留', tcontent.includes('测试主标题') && tcontent.includes('第一段') && tcontent.includes('最后一段正文'), '')

    const c1 = await call('office_convert', { source: htmlFile, target: join(OUT, 'r1.md') })
    ok('R1 convert html→md', existsSync(c1.target) && (await readFile(join(OUT, 'r1.md'), 'utf8')).includes('测试主标题'))
    const c2 = await call('office_convert', { source: htmlFile, target: join(OUT, 'r1.docx') })
    ok('R1 convert html→docx 可回读', existsSync(c2.target) && (await textOf(join(OUT, 'r1.docx'))).includes('第二章 嵌套列表'))
    const c3 = await call('office_convert', { source: htmlFile, target: join(OUT, 'r1.txt') })
    ok('R1 convert html→txt', existsSync(c3.target) && (await readFile(join(OUT, 'r1.txt'), 'utf8')).includes('A&B'))
    const c4 = await call('office_create', { path: join(OUT, 'r1-from-html.md'), from: htmlFile })
    ok('R1 create from=html（转换式创建）', existsSync(c4.path))

    // 真实样本（存在才跑）：上一轮 OCR 重建的两份学习笔记 HTML —— 含大量表格/嵌套列表
    for (const real of ['<work>/deepseek9/sample-A-样本.html',
      '<work>/deepseek9/sample-topic-2-样本-汇总.html']) {
      if (!existsSync(real)) continue
      const rmeta = await call('office_read', { path: real, as: 'meta' })
      ok(`R1 真实样本 meta：${basename(real)}`, rmeta.format === 'html' && rmeta.stats.blocks > 50 && rmeta.stats.textLayerUsable === true, JSON.stringify(rmeta.stats))
      const rmd = await call('office_read', { path: real })
      ok(`R1 真实样本正文：${basename(real)}`, String(rmd.content).includes('背诵') && String(rmd.content).length > 5000, `${String(rmd.content).length} 字符`)
    }
  }

  // ---------------- 需求 2：批量文件 stats 探测 ----------------
  {
    const scanDir = join(OUT, 'scan-dir')
    mkdirSync(scanDir, { recursive: true })
    const cleanPdf = join(scanDir, 'scan-clean.pdf')
    {
      const blocks = []
      for (let p = 1; p <= 3; p++) {
        if (p > 1) blocks.push({ type: 'pagebreak' })
        blocks.push({ type: 'paragraph', runs: [{ text: `这是第 ${p} 页的正文内容，用于批量扫描统计测试，语义完整。`.repeat(3) }] })
      }
      writeFileSync(cleanPdf, writePdf({ kind: 'document', meta: {}, blocks }, {}))
    }
    // 乱码 PDF：25 页 PUA —— 文字层不可用 → garbledPages>0、OCR 建议 2 批（纯 CPU，无 OCR 消耗）
    writeFileSync(join(scanDir, 'scan-garbled.pdf'), buildGarbledFixture({ pages: 25 }))
    writeFileSync(join(scanDir, 'scan-table.csv'), '城市,人口\n唐山,770\n', 'utf8')
    writeFileSync(join(scanDir, 'scan-note.md'), '# 笔记\n\n- 一\n- 二\n', 'utf8')
    writeFileSync(join(scanDir, 'scan-page.html'), '<html><head><title>t</title></head><body><h1>标题</h1><p>正文段落，批量扫描用。</p></body></html>', 'utf8')
    writeFileSync(join(scanDir, 'ignore-me.xyz'), '不是办公文件', 'utf8')

    const scan = await call('office_read', { paths: [scanDir] })
    ok('R2 目录扫描：5 个受支持文件（.xyz 被忽略、不计 failed）', scan.format === 'scan' && scan.total === 5 && scan.ok === 5 && scan.failed === 0, `total=${scan.total} ok=${scan.ok} failed=${scan.failed}`)
    const cleanRow = scan.files.find(f => f.name === 'scan-clean.pdf')
    ok('R2 pdf 行字段与 meta.stats 同名', cleanRow && ['format', 'pages', 'characters', 'textLayerUsable', 'scannedPages', 'garbledPages', 'suggestedBatches'].every(k => k in cleanRow), JSON.stringify(cleanRow))
    ok('R2 非 pdf 行也有 characters/textLayerUsable/建议批次', scan.files.filter(f => f.format !== 'pdf').every(f => 'characters' in f && 'textLayerUsable' in f && 'suggestedBatches' in f), '')
    ok('R2 正文绝不返回（行里没有 content 字段）', scan.files.every(f => !('content' in f)), '')
    ok('R2 干净 PDF 行：textLayerUsable=true、garbledPages=0', cleanRow.textLayerUsable === true && cleanRow.garbledPages === 0 && cleanRow.pages === 3, '')
    const gRow = scan.files.find(f => f.name === 'scan-garbled.pdf')
    ok('R2 乱码 PDF 行：不可用 + garbledPages>0 + 建议 2 批 OCR', gRow && gRow.textLayerUsable === false && gRow.garbledPages > 0 && gRow.suggestedBatches === 2, JSON.stringify({ usable: gRow?.textLayerUsable, garbled: gRow?.garbledPages, batches: gRow?.suggestedBatches }))
    ok('R2 content 是清单表格', String(scan.content).includes('| 文件 |') && String(scan.content).includes('scan-clean.pdf'), String(scan.content).split('\n')[0])

    const mix = await call('office_read', { paths: [cleanPdf, join(scanDir, '不存在的.pdf'), join(scanDir, 'scan-note.md')] })
    ok('R2 混合清单：坏路径进 skipped，不连累整批', mix.total === 2 && mix.ok === 2 && Array.isArray(mix.skipped) && mix.skipped.length === 1, JSON.stringify(mix.skipped))

    let scanThrew = ''
    try { await call('office_read', { paths: [] }) } catch (e) { scanThrew = e.message }
    ok('R2 paths 为空报错（带下一步）', /paths 为空/.test(scanThrew), scanThrew.slice(0, 80))
    scanThrew = ''
    try { await call('office_read', { paths: [join(OUT, 'no-such-dir')] }) } catch (e) { scanThrew = e.message }
    ok('R2 全部无效时报错并列出原因', /没有可扫描的文件/.test(scanThrew) && /路径不存在/.test(scanThrew), scanThrew.slice(0, 120))
    scanThrew = ''
    try { await call('office_read', { paths: [cleanPdf], as: 'markdown' }) } catch (e) { scanThrew = e.message }
    ok('R2 批量形态不支持 as=markdown（显式报错不静默）', /不支持 as/.test(scanThrew), scanThrew.slice(0, 80))
    scanThrew = ''
    try { await call('office_read', {}) } catch (e) { scanThrew = e.message }
    ok('R2 path/paths 都缺时报错', /path/.test(scanThrew) && /paths/.test(scanThrew), scanThrew.slice(0, 80))
    scanThrew = ''
    try { await call('office_read', { path: cleanPdf, paths: [cleanPdf] }) } catch (e) { scanThrew = e.message }
    ok('R2 path 与 paths 互斥报错', /二选一/.test(scanThrew), scanThrew.slice(0, 80))
  }

  // ---------------- 需求 3：页级续读 + 批间衔接 ----------------
  {
    const book = join(OUT, 'r3-book.pdf')
    {
      const lines = []
      for (let p = 1; p <= 5; p++) {
        if (p > 1) lines.push('---page---')
        lines.push(`## 第 ${p} 章标题`, `这是第 ${p} 页的正文段落。本页语义必须跨批完整：${'连续内容'.repeat(6)}页尾句子在第 ${p} 页收束。`)
      }
      await call('office_create', { path: book, markdown: lines.join('\n\n') })
    }
    const stripMarkers = s => String(s).replace(/<!--[^>]*-->\s*/g, '')
    const whole = await call('office_read', { path: book, as: 'markdown' })
    ok('R3 整本读取不带批间字段（旧调用零影响）', whole.stats.pageFrom === undefined && whole.stats.prevTail === undefined && whole.stats.nextPage === undefined, '')

    const b1 = await call('office_read', { path: book, as: 'markdown', pages: '1-2' })
    const b2 = await call('office_read', { path: book, as: 'markdown', pages: '3-5' })
    ok('R3 分批 stats：pageFrom/pageTo/nextPage', b1.stats.pageFrom === 1 && b1.stats.pageTo === 2 && b1.stats.nextPage === 3
      && b2.stats.pageFrom === 3 && b2.stats.pageTo === 5 && b2.stats.nextPage === undefined,
      `b1:${b1.stats.pageFrom}-${b1.stats.pageTo}→${b1.stats.nextPage} b2:${b2.stats.pageFrom}-${b2.stats.pageTo}→${b2.stats.nextPage}`)
    ok('R3 上一批末行回看（prevTail 指向第 2 页末尾）', b2.stats.prevTail?.page === 2 && String(b2.stats.prevTail.text).length > 0, JSON.stringify(b2.stats.prevTail))
    ok('R3 下一批首行预览（nextHead 指向第 3 页开头）', b1.stats.nextHead?.page === 3 && String(b1.stats.nextHead.text).length > 0, JSON.stringify(b1.stats.nextHead))
    ok('R3 回看/预览在 100 字符量级', String(b1.stats.nextHead.text).length <= 102 && String(b2.stats.prevTail.text).length <= 102, `${String(b1.stats.nextHead.text).length}/${String(b2.stats.prevTail.text).length}`)
    ok('R3 notice 带批间衔接摘要（boundary 缺省时）', /批间衔接/.test(String(b1.notice)) && /批间衔接/.test(String(b2.notice)), String(b1.notice || '').slice(0, 120))

    const joined = `${stripMarkers(b1.content)}\n\n${stripMarkers(b2.content)}`
    ok('R3 分批拼接 ≡ 整本读取（段落完整性 diff）', joined === stripMarkers(whole.content), `joined=${joined.length} whole=${stripMarkers(whole.content).length}`)

    const p23 = await call('office_read', { path: book, as: 'markdown', pageFrom: 2, pageTo: 3 })
    const pages23 = await call('office_read', { path: book, as: 'markdown', pages: '2-3' })
    ok('R3 pageFrom/pageTo ≡ pages="N-M"（正文逐字一致）', p23.content === pages23.content, '')
    ok('R3 pageFrom/pageTo 的 stats 回显', p23.stats.pageFrom === 2 && p23.stats.pageTo === 3, '')
    const p4 = await call('office_read', { path: book, as: 'markdown', pageFrom: 4 })
    ok('R3 pageFrom 单给 = 从第 4 页读到最后一页', p4.stats.pageFrom === 4 && p4.stats.pageTo === 5, '')

    let r3threw = ''
    try { await call('office_read', { path: book, pages: '1-2', pageFrom: 1 }) } catch (e) { r3threw = e.message }
    ok('R3 pages 与 pageFrom 互斥报错', /不能同时给/.test(r3threw), r3threw.slice(0, 80))
    r3threw = ''
    try { await call('office_read', { path: book, pageFrom: 3, pageTo: 2 }) } catch (e) { r3threw = e.message }
    ok('R3 pageFrom>pageTo 报错（页码范围无效）', /页码范围无效/.test(r3threw), r3threw.slice(0, 80))

    const bb = await call('office_read', { path: book, as: 'markdown', pages: '2-3', boundary: true })
    const b3 = await call('office_read', { path: book, as: 'markdown', pages: '2-3' })
    ok('R3 boundary=true：正文首尾内联回看/预览', String(bb.content).startsWith('> 〔dsh-office 批间回看') && String(bb.content).includes('〔dsh-office 批间预览'), String(bb.content).slice(0, 90))
    const stripBoundary = s => String(s).split('\n').filter(l => !l.includes('〔dsh-office 批间')).join('\n').replace(/^\n+|\s+$/g, '')
    ok('R3 剔除内联衔接行后与默认形态逐字一致（拼接可剔除）', stripBoundary(stripMarkers(bb.content)) === stripMarkers(b3.content).trim(), '')
    ok('R3 boundary=true 时 notice 不重复批间摘要', !/批间衔接/.test(String(bb.notice || '')), String(bb.notice || '<无 notice>').slice(0, 80))
  }

  // ---------------- 需求 4：字符级质量启发（structural 只在页面级启用） ----------------
  {
    const hi = 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞß'
    const qHi = textQuality(hi.repeat(3))
    ok('R4 Latin-1 高带占优（mojibake/CID 错映射）被检出', qHi.garbled && qHi.reasons.some(r => /Latin-1 扩展/.test(r)), qHi.reasons.join('；'))
    const qNv = textQuality(('bcdfghjklmnpqrstv ').repeat(12))
    ok('R4 连续非词典字符（无元音字母串）被检出', qNv.garbled && qNv.reasons.some(r => /非词典/.test(r)), qNv.reasons.join('；'))
    const qDup = textQuality(Array.from({ length: 12 }, () => '第一行重复出现的解析碎片').join('\n'), null, { structural: true })
    ok('R4 重复行率被检出（structural）', qDup.garbled && qDup.reasons.some(r => /重复行/.test(r)), qDup.reasons.join('；'))
    const qShort = textQuality('第\n一\n行\n都\n是\n单\n个\n汉\n字\n的\n页\n面', null, { structural: true })
    ok('R4 单字/双字行率被检出（逐字断行）', qShort.garbled && qShort.reasons.some(r => /单字\/双字行/.test(r)), qShort.reasons.join('；'))

    const cleanCjk = '这是一段完全正常的中文正文，字符语义完整、行结构正常，质量门不应该误报这段文字。'.repeat(6)
    ok('R4 正常中文页不误报', !textQuality(cleanCjk, null, { structural: true }).garbled, '')
    ok('R4 正常英文页不误报', !textQuality('The quick brown fox jumps over the lazy dog near the river bank every morning. '.repeat(6), null, { structural: true }).garbled, '')
    ok('R4 俄文页不误报（希腊/西里尔只参与相对判据）', !textQuality('Съешь же ещё этих мягких французских булок, да выпей чаю. '.repeat(6), null, { structural: true }).garbled, '')
    const codeLines = Array.from({ length: 10 }, (_, i) => `const getUserInfo${i} = document.querySelector("#app-${i}"); return getUserInfo${i};`)
    ok('R4 代码页不误报（标识符不构成无元音连排、行不重复）', !textQuality(codeLines.join('\n'), null, { structural: true }).garbled, '')
    const toc = Array.from({ length: 14 }, (_, i) => `第${i + 1}章 标题内容..................${i + 3}`).join('\n')
    ok('R4 目录点前导页不误报（第三轮判据保持）', !textQuality(toc, null, { structural: true }).garbled, '')
    ok('R4 相对判据：干净书中混入西里尔页被检出', textQuality('Съешь же ещё этих мягких французских булок, да выпей чаю. '.repeat(8), { cjkRatio: 0.7, oddCharRatio: 0 }, { structural: true }).garbled, '')
    // 整本口径（structural=false，finishRead/convert 探针）：版面启发不启用 → 不新增 sidecar 误报
    const qDoc = textQuality(Array.from({ length: 12 }, () => '第一行重复出现的解析碎片').join('\n'))
    ok('R4 整本口径不启用版面启发（误报不高于现水平）', !qDoc.garbled, qDoc.reasons.join('；'))
    // 旧判据回归：私用区/替换字符/控制符三条照旧独立命中
    ok('R4 旧判据回归：PUA 页仍被检出', textQuality('\uE0A1\uE0A2\uE0A3'.repeat(30)).garbled, '')
  }

  // ---------------- 需求 5：CSV/文本写出的编码与转义收敛 ----------------
  {
    const csvPath = join(OUT, 'r5.csv')
    await call('office_create', {
      path: csvPath,
      workbook: { sheets: [{ name: 'S1', rows: [['城市', '备注'], ['唐山', '含,逗号'], ['保定', '含"引号"'], ['北京', '含\n换行']] }] },
    })
    const raw = readFileSync(csvPath)
    ok('R5 csv 默认 UTF-8 with BOM（Excel 双击不乱码）', raw[0] === 0xef && raw[1] === 0xbb && raw[2] === 0xbf, raw.subarray(0, 4).toString('hex'))
    const csvText = raw.toString('utf8')
    ok('R5 逗号字段转义（RFC 4180）', csvText.includes('"含,逗号"'), csvText.split('\r\n')[1])
    ok('R5 引号字段转义（内部引号成对双写）', csvText.includes('"含""引号"""'), csvText.split('\r\n')[2])
    ok('R5 换行字段整体加引号（内嵌真实换行）', csvText.includes('"含\n换行"'), '')
    const reread = await textOf(csvPath)
    ok('R5 回读无损（唐山/含,逗号/含"引号" 都在）', reread.includes('唐山') && reread.includes('含,逗号') && reread.includes('含"引号"'), reread.slice(0, 160))

    const csvNoBom = join(OUT, 'r5-nobom.csv')
    await call('office_create', { path: csvNoBom, table: { rows: [['a', 'b']] }, encoding: 'utf-8' })
    const raw2 = readFileSync(csvNoBom)
    ok('R5 encoding=utf-8 显式无 BOM', !(raw2[0] === 0xef && raw2[1] === 0xbb && raw2[2] === 0xbf), raw2.subarray(0, 4).toString('hex'))
    const mdBom = join(OUT, 'r5-bom.md')
    await call('office_create', { path: mdBom, markdown: '# 有BOM\n', encoding: 'utf-8-sig' })
    ok('R5 md 目标也可指定 utf-8-sig', readFileSync(mdBom)[0] === 0xef, '')

    await call('office_edit', { path: csvPath, operations: [{ op: 'append_rows', rows: [['广州', '100']] }] })
    ok('R5 office_edit 后 BOM 保留（不再被静默剥掉）', readFileSync(csvPath)[0] === 0xef, '')

    await call('office_convert', { source: join(OUT, 'data.csv'), target: join(OUT, 'r5-conv.csv') })
    ok('R5 convert → csv 默认带 BOM', readFileSync(join(OUT, 'r5-conv.csv'))[0] === 0xef, '')
    let encThrew = ''
    try { await call('office_create', { path: join(OUT, 'r5.xlsx'), workbook: { sheets: [{ name: 'S', rows: [['a']] }] }, encoding: 'utf-8' }) } catch (e) { encThrew = e.message }
    ok('R5 encoding 只对文本类目标生效（xlsx 显式报错）', /encoding/.test(encThrew) && /csv\/tsv\/md\/txt/.test(encThrew), encThrew.slice(0, 90))
  }

  // ===========================================================================
  // 阶段一（第五轮）：中文 PDF 内嵌字体子集  /  阶段二：PDF 产出质量门
  // 前缀：嵌字体： / 质量门：
  // ===========================================================================
  {
    // ---- 阶段一 a/c：字节级断言 + office_read 提取不丢字符 ----
    const emb = join(OUT, 'embed-cn.pdf')
    const embRet = await call('office_create', {
      path: emb,
      markdown: '# 嵌入字体验收\n\n中文正文：需求是指消费者愿意购买的数量。符号 ☆ ★ ① ② ✅ ❌ 地球 🌍 混排 The quick brown fox。\n\n| 区域 | 营收 |\n| --- | --- |\n| 华北 | 520 |\n',
    })
    const embRaw = readFileSync(emb, 'latin1')
    ok('嵌字体：FontFile2 > 0', (embRaw.match(/\/FontFile2/g) || []).length > 0, `${statSync(emb).size} 字节`)
    ok('嵌字体：BaseFont 带 XXXXXX+ 子集前缀', /\/BaseFont\/[A-Z]{6}\+/.test(embRaw))
    ok('嵌字体：ToUnicode > 0', (embRaw.match(/\/ToUnicode/g) || []).length > 0)
    ok('嵌字体：Identity-H + CIDFontType2', /Encoding\/Identity-H/.test(embRaw) && /CIDFontType2/.test(embRaw))
    const embRead = await call('office_read', { path: emb })
    ok('嵌字体：office_read 提取 ☆ ① 不再是 ?', embRead.content.includes('☆') && embRead.content.includes('①') && !/[☆★①②]\?/.test(embRead.content),
      embRead.content.slice(0, 70).replace(/\n/g, ' '))
    ok('嵌字体：emoji 提取保真（🌍✅）', embRead.content.includes('🌍') && embRead.content.includes('✅'))
    ok('嵌字体：体积护栏（样本 <2MB；9000 字实测 0.39MB）', statSync(emb).size < 2 * 1048576, `${statSync(emb).size} 字节`)

    // ---- 阶段二 a：好 PDF → stats.pdfQuality（只增不改、返回值无损）----
    ok('质量门：stats.pdfQuality = embedded=true / renderCheck=pass / firstPageBytes>0',
      embRet.stats?.pdfQuality?.embedded === true && embRet.stats.pdfQuality.renderCheck === 'pass'
      && embRet.stats.pdfQuality.firstPageBytes > 0, JSON.stringify(embRet.stats?.pdfQuality))
    ok('质量门：create 返回值仍无损', losslessJsonProblem(embRet) === null)

    // ---- 阶段二 b：坏 PDF（未嵌字体的中文）被拦，四要素 + 逃生通道，不落盘 ----
    process.env.DSH_OFFICE_PDF_EMBED_CJK = '0'
    const badGate = join(OUT, 'gate-should-not-exist.pdf')
    rmSync(badGate, { force: true })
    let gateErr = ''
    try { await call('office_create', { path: badGate, markdown: '# 拦我\n\n这段中文必须触发字体未嵌入拦截。\n' }) } catch (e) { gateErr = String(e && e.message) }
    delete process.env.DSH_OFFICE_PDF_EMBED_CJK
    ok('质量门：未嵌字体的中文 PDF 被拦（四要素 + 逃生通道可复制）',
      ['页码=', '格式=', '根因=', '下一步='].every(k => gateErr.includes(k)) && /逃生通道/.test(gateErr) && /ExportAsFixedFormat/.test(gateErr),
      gateErr.slice(0, 130))
    ok('质量门：拒绝时绝不产出目标文件', !existsSync(badGate))

    // ---- 阶段二 b2：渲染为空白 → fail（字节级诱过 FontFile 串、WinRT 实际渲染空白）----
    process.env.DSH_OFFICE_PDF_EMBED_CJK = '0'
    const blankModel = { kind: 'document', meta: {}, blocks: [{ type: 'paragraph', runs: [{ text: '中文渲染空白探针字符' }] }] }
    const stBuf = writePdf(blankModel, {})
    delete process.env.DSH_OFFICE_PDF_EMBED_CJK
    const trick = Buffer.concat([stBuf, Buffer.from('\n% /FontFile2 probe\n', 'latin1')])
    const blankProbe = join(OUT, 'gate-blank-probe.pdf')
    rmSync(blankProbe, { force: true })
    let blankErr = ''
    try { await pdfOutputGate(blankProbe, trick, blankModel, { embedded: false }) } catch (e) { blankErr = String(e && e.message) }
    ok('质量门：渲染为空白 → fail（四要素）且不落盘',
      /渲染为空白/.test(blankErr) && ['页码=', '格式=', '根因=', '下一步='].every(k => blankErr.includes(k)) && !existsSync(blankProbe),
      blankErr.slice(0, 150))

    // ---- 阶段二 d：DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1 → skipped ----
    process.env.DSH_OFFICE_PDF_SKIP_RENDER_CHECK = '1'
    const skipRet = await call('office_create', { path: join(OUT, 'gate-skip.pdf'), markdown: '# 跳过渲染\n\n中文正文一行。\n' })
    delete process.env.DSH_OFFICE_PDF_SKIP_RENDER_CHECK
    ok('质量门：DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1 → renderCheck=skipped',
      skipRet.stats?.pdfQuality?.renderCheck === 'skipped' && skipRet.stats.pdfQuality.embedded === true,
      JSON.stringify(skipRet.stats?.pdfQuality))

    // ---- 阶段二：无可见文本（hr-only）不被"空白"误杀（既有 fixture 的命门）----
    const hrRet = await call('office_create', { path: join(OUT, 'gate-hr.pdf'), document: { blocks: [{ type: 'hr' }] } })
    ok('质量门：无可见文本文档不误杀（renderCheck=pass）',
      hrRet.stats?.pdfQuality?.renderCheck === 'pass', JSON.stringify(hrRet.stats?.pdfQuality))

    // ---- 阶段二：convert → pdf 走同一条闸 ----
    const convGate = await call('office_convert', { source: join(OUT, 'note.md'), target: join(OUT, 'gate-conv.pdf') })
    ok('质量门：convert → pdf 带 stats.pdfQuality（embedded=true / pass）',
      convGate.stats?.pdfQuality?.embedded === true && convGate.stats.pdfQuality.renderCheck === 'pass',
      JSON.stringify(convGate.stats?.pdfQuality))

    // ---- R18 任务 A：渲染校验的暂存件必须落 %TEMP%（本机 WinRT 只读得到 %TEMP% 内的 PDF）----
    {
      ok('R18-质量门：默认渲染暂存目录在 tmpdir() 下、**不在目标目录旁边**',
        !process.env.DSH_OFFICE_PDF_GATE_DIR
        && resolve(pdfGateStagingRoot()).startsWith(resolve(tmpdir())),
        pdfGateStagingRoot())
      const nt = nonTempWritableDir()
      if (!nt) {
        ok('R18-质量门：非 TEMP 目标目录 create/convert PDF → 跳过（找不到可写的非 TEMP、非插件目录）',
          true, '沙箱 workspace-write 下只能写 %TEMP%；用 danger-full-access 重跑即可覆盖')
      } else {
        const p1 = join(nt, 'r18-create.pdf')
        rmSync(p1, { force: true })
        const r1 = await call('office_create', { path: p1, markdown: '# 中文标题\n\n这是一段中文正文，用来验证质量门。\n' })
        ok('R18-质量门：**非 TEMP** 目标目录 create PDF 成功（旧版 100% 报"第 1 页渲染失败"）',
          existsSync(p1) && r1.stats?.pdfQuality?.renderCheck === 'pass',
          `nt=${nt} ${JSON.stringify(r1.stats?.pdfQuality)}`)
        const mdNt = join(nt, 'r18-note.md')
        writeFileSync(mdNt, '# 转换标题\n\n转换后的中文正文。\n', 'utf8')
        const p2 = join(nt, 'r18-conv.pdf')
        rmSync(p2, { force: true })
        const r2 = await call('office_convert', { source: mdNt, target: p2 })
        ok('R18-质量门：**非 TEMP** 目标目录 convert → pdf 成功',
          existsSync(p2) && r2.stats?.pdfQuality?.renderCheck === 'pass', JSON.stringify(r2.stats?.pdfQuality))
        ok('R18-质量门：非 TEMP 目标目录里不留临时件 / 渲染副本（之前这里会落下 gate 半成品）',
          readdirSync(nt).every(n => !n.startsWith('.dsh-tmp-') && !n.endsWith('.part')),
          readdirSync(nt).join(',').slice(0, 90))
        process.env.DSH_OFFICE_PDF_EMBED_CJK = '0'
        const bm = { kind: 'document', meta: {}, blocks: [{ type: 'paragraph', runs: [{ text: '中文渲染空白探针字符' }] }] }
        const sb = writePdf(bm, {})
        delete process.env.DSH_OFFICE_PDF_EMBED_CJK
        const trick2 = Buffer.concat([sb, Buffer.from('\n% /FontFile2 probe\n', 'latin1')])
        const p3 = join(nt, 'r18-blank.pdf')
        rmSync(p3, { force: true })
        let e3 = ''
        try { await pdfOutputGate(p3, trick2, bm, { embedded: false }) } catch (e) { e3 = String(e.message) }
        ok('R18-质量门：非 TEMP 目录里"渲染为空白"仍被拒、目标不存在、不留临时件',
          /渲染为空白/.test(e3) && !existsSync(p3)
          && readdirSync(nt).every(n => !n.startsWith('.dsh-tmp-') && !n.endsWith('.part')),
          e3.slice(0, 100))
      }
      // DSH_OFFICE_PDF_GATE_DIR 覆盖生效（指向 %TEMP% 内的子目录 → 仍能成功）
      const override = join(tmpdir(), 'r18-gate-override')
      process.env.DSH_OFFICE_PDF_GATE_DIR = override
      try {
        const p4 = join(OUT, 'r18-gate-override.pdf')
        rmSync(p4, { force: true })
        const r4 = await call('office_create', { path: p4, markdown: '# 覆盖暂存目录\n\n中文正文。\n' })
        ok('R18-质量门：DSH_OFFICE_PDF_GATE_DIR 覆盖生效（指向 %TEMP% 内子目录仍产出成功）',
          resolve(pdfGateStagingRoot()) === resolve(override) && existsSync(p4)
          && r4.stats?.pdfQuality?.renderCheck === 'pass', pdfGateStagingRoot())
      } finally { delete process.env.DSH_OFFICE_PDF_GATE_DIR }
      ok('R18-质量门：覆盖的暂存目录在渲染后被清空（暂存件不残留）',
        !existsSync(override) || readdirSync(override).length === 0,
        existsSync(override) ? `残余: ${readdirSync(override).join(',')}` : '(目录不存在)')
      if (nt) { try { rmSync(nt, { recursive: true, force: true }) } catch { /* ignore */ } }
    }
  }

  // ===========================================================================
  // 阶段三/四：原生 Office 开箱冒烟（Word COM）+ .html 目标
  // 前缀：开箱： / html：
  // ===========================================================================
  {
    // 修复 0x800A1401 的一线防线：写出的 docx 必须能被真实 Word 打开并导出 PDF。
    // 路径只走环境变量（避免中文/引号在命令行上被切碎）；脚本 ASCII-only。
    const NATIVE_PS = `
$ErrorActionPreference = 'Continue'
$src = $env:DSH_OFFICE_NATIVE_SRC
$out = $env:DSH_OFFICE_NATIVE_PDF
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300
Remove-Item (Join-Path $env:APPDATA 'Microsoft\\Word\\*.asd') -Force -ErrorAction SilentlyContinue
try { $w = New-Object -ComObject Word.Application } catch { 'NATIVE_NO_WORD ' + $_.Exception.Message; exit 0 }
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($src, $false, $true)
  if ($out) { $d.ExportAsFixedFormat($out, 17) }
  $n = $d.Paragraphs.Count
  $d.Close(0)
  "NATIVE_OPEN_OK paragraphs=$n"
} catch {
  'NATIVE_OPEN_FAIL ' + $_.Exception.Message
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
`
    const wordProbe = (src, pdfOut) => {
      const env = { ...process.env, DSH_OFFICE_NATIVE_SRC: src }
      if (pdfOut) env.DSH_OFFICE_NATIVE_PDF = pdfOut
      else delete env.DSH_OFFICE_NATIVE_PDF
      return spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', NATIVE_PS],
        { encoding: 'utf8', windowsHide: true, timeout: 240000, env })
    }
    const probeText = r => `${(r && r.stdout) || ''}${(r && r.stderr) || ''}`.trim()
    const noWord = t => !t || /NATIVE_NO_WORD|Retrieving the COM class factory|80040154|not registered|is not recognized/i.test(t)
    const brief = t => t.replace(/\s+/g, ' ').slice(0, 150)

    if (process.platform !== 'win32') {
      ok('开箱：Word COM 打开产出的 .docx（非 Windows → 跳过）', true, `platform=${process.platform}`)
      ok('开箱：Word 打开产出的 .html（非 Windows → 跳过）', true, `platform=${process.platform}`)
    } else {
      const nativeDocx = join(OUT, 'native-open.docx')
      await call('office_create', { path: nativeDocx, markdown: '# 原生开箱冒烟\n\n中文正文一行，用于 Word COM 打开验证。\n\n- 列表项\n' })
      const nativePdf = join(OUT, 'native-open.pdf')
      rmSync(nativePdf, { force: true })
      const dTxt = probeText(wordProbe(nativeDocx, nativePdf))
      if (noWord(dTxt)) {
        ok('开箱：Word COM 打开产出的 .docx（Word 不可用 → 跳过）', true, brief(dTxt))
      } else {
        ok('开箱：Word 16.0 能打开产出的 .docx（不再 0x800A1401）', /NATIVE_OPEN_OK/.test(dTxt), brief(dTxt))
        ok('开箱：Word 能把产出的 .docx 导出 PDF（ExportAsFixedFormat）',
          existsSync(nativePdf) && statSync(nativePdf).size > 2000, `${existsSync(nativePdf) ? statSync(nativePdf).size : 0} bytes`)
      }

      const nativeHtml = join(OUT, 'native-open.html')
      await call('office_create', { path: nativeHtml, markdown: '# Word 开箱冒烟\n\n中文正文一行。\n\n| 列 | 值 |\n| --- | --- |\n| 甲 | 1 |\n' })
      const hTxt = probeText(wordProbe(nativeHtml, ''))
      if (noWord(hTxt)) ok('开箱：Word 打开产出的 .html（Word 不可用 → 跳过）', true, brief(hTxt))
      else ok('开箱：Word 能打开产出的 .html', /NATIVE_OPEN_OK/.test(hTxt), brief(hTxt))
    }

    // ---- 阶段四：.html 目标是语义化 HTML5，且 md→html→md 往返不丢内容 ----
    const HTML_MD = `# 季度经营报告

本期营收 **1280** 万元，同比增长 *12%*，见 [官网](https://example.com)，字段名 \`textLayerUsable\`。

## 要点

- 零第三方依赖
- 全格式互转
  - 嵌套一层
- 行列齐全

1. 第一条
2. 第二条

> 引用：数据以审计后为准

| 区域 | 营收 |
| --- | --- |
| 华北 | 1280 |
| 华东 | 960 |

\`\`\`
code line <tag> & "quoted"
\`\`\`

---
`
    const htmlSrc = join(OUT, 'rt-src.md')
    const htmlMid = join(OUT, 'rt-mid.html')
    const htmlBack = join(OUT, 'rt-back.md')
    const mdDirect = join(OUT, 'rt-direct.md')
    writeFileSync(htmlSrc, HTML_MD, 'utf8')
    const created = await call('office_create', { path: htmlMid, markdown: HTML_MD })
    const htmlText = readFileSync(htmlMid, 'utf8')
    ok('html：create 产出语义化 HTML5（doctype/charset/h1/table/strong/blockquote）',
      /^<!DOCTYPE html>/.test(htmlText) && /<meta charset="utf-8">/.test(htmlText)
      && /<h1>/.test(htmlText) && /<table>/.test(htmlText) && /<strong>/.test(htmlText) && /<blockquote>/.test(htmlText),
      `${created.bytes} bytes`)
    ok('html：UTF-8 且中文不转实体',
      htmlText.includes('季度经营报告') && !/&#\d+;/.test(htmlText) && !/&lt;h1/.test(htmlText))
    const htmlRead = await call('office_read', { path: htmlMid })
    ok('html：office_read 回读保留结构（标题/加粗/表格）',
      htmlRead.content.includes('# 季度经营报告') && htmlRead.content.includes('**1280**') && htmlRead.content.includes('| 华北 | 1280 |'),
      htmlRead.content.split('\n')[0])
    await call('office_convert', { source: htmlSrc, target: mdDirect })
    await call('office_convert', { source: htmlMid, target: htmlBack })
    ok('html：md→html→md 往返逐字不丢内容',
      readFileSync(htmlBack, 'utf8') === readFileSync(mdDirect, 'utf8'),
      `${readFileSync(htmlBack, 'utf8').length} vs ${readFileSync(mdDirect, 'utf8').length} 字符`)
    const convHtml = join(OUT, 'rt-conv.html')
    const convRet = await call('office_convert', { source: htmlSrc, target: convHtml })
    ok('html：convert → html 走同一条写出端', convRet.targetFormat === 'html' && /<h1>/.test(readFileSync(convHtml, 'utf8')))
  }

  // ===========================================================================
  // 第十二轮（自 WorkBuddy 侧移植 + DSH 侧补兜底）：插图链路 / AES-256 / 公式重算 / 稿纸网格
  // ===========================================================================

  /**
   * OOXML 包结构体检（纯 JS）：①每个部件都必须被 Default 或 Override 覆盖；
   * ②每个 Override 的 PartName 必须真实存在；③`extended-properties` 不许挂在
   * wordprocessingml/presentationml/spreadsheetml 命名空间下（历史病灶）。
   */
  function checkOoxmlPackage(buf, label) {
    const z = openZip(buf)
    const ct = z.getText('[Content_Types].xml') || ''
    // 属性顺序不固定（真实样本里 ContentType 可能写在 Extension 之前），逐标签抓取。
    const defaults = new Map()
    for (const m of ct.matchAll(/<Default\b[^>]*?\/?>/gi)) {
      const e = /Extension="([^"]+)"/i.exec(m[0])?.[1]
      const t = /ContentType="([^"]+)"/i.exec(m[0])?.[1]
      if (e && t) defaults.set(e.toLowerCase(), t)
    }
    const overrides = new Map()
    for (const m of ct.matchAll(/<Override\b[^>]*?\/?>/gi)) {
      const p = /PartName="([^"]+)"/i.exec(m[0])?.[1]
      const t = /ContentType="([^"]+)"/i.exec(m[0])?.[1]
      if (p && t) overrides.set(p, t)
    }
    const problems = []
    const names = z.names.filter(n => !n.endsWith('/'))
    for (const n of names) {
      const ext = (n.split('.').pop() || '').toLowerCase()
      if (!overrides.has('/' + n) && !defaults.has(ext)) problems.push(`部件无内容类型: ${n}`)
    }
    for (const p of overrides.keys()) if (!z.has(p.replace(/^\//, ''))) problems.push(`Override 指向不存在的部件: ${p}`)
    for (const [part, type] of overrides) {
      if (/openxmlformats-officedocument\.(wordprocessingml|presentationml|spreadsheetml)\.extended-properties\+xml$/.test(type)) {
        problems.push(`extended-properties 内容类型非法（历史病灶）: ${part} → ${type}`)
      }
    }
    return { label, ok: problems.length === 0, problems, parts: names.length }
  }

  // ---------------- R12-0：OOXML 包结构（Word 拒开防回归；插图后同样要过） ----------------
  {
    const pkgDocx = join(OUT, 'r12-pkg.docx')
    await call('office_create', { path: pkgDocx, markdown: '# 标题\n\n正文\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n' })
    const r0 = checkOoxmlPackage(readFileSync(pkgDocx), 'r12-docx')
    ok('R12-0 包结构：docx 内容类型完备且无非法 extended-properties',
      r0.ok, r0.ok ? `${r0.parts} 部件` : r0.problems.join('；'))
    {
      const z = openZip(readFileSync(pkgDocx))
      const ctBad = (z.getText('[Content_Types].xml') || '').replace(
        'application/vnd.openxmlformats-officedocument.extended-properties+xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.extended-properties+xml')
      const entries = z.names.map(n => ({ name: n, data: n === '[Content_Types].xml' ? ctBad : z.get(n) }))
      const r = checkOoxmlPackage(Buffer.from(makeZip(entries)), 'negative-control')
      ok('R12-0 包结构：负向控制 —— 塞回病灶内容类型必须被判不合规',
        !r.ok && r.problems.some(p => /extended-properties 内容类型非法/.test(p)), r.problems.join('；'))
    }
    {
      const z = openZip(readFileSync(pkgDocx))
      const ctSrc = z.getText('[Content_Types].xml') || ''
      const ctSwapped = ctSrc.replace(/<Default\b[^>]*?\/?>/gi, tag => {
        const e = /Extension="([^"]+)"/i.exec(tag)?.[1]
        const t = /ContentType="([^"]+)"/i.exec(tag)?.[1]
        return (e && t) ? `<Default ContentType="${t}" Extension="${e}"/>` : tag
      })
      const entries = z.names.map(n => ({ name: n, data: n === '[Content_Types].xml' ? ctSwapped : z.get(n) }))
      const r = checkOoxmlPackage(Buffer.from(makeZip(entries)), 'ctype-attr-order')
      ok('R12-0 包结构：属性顺序颠倒（ContentType 在前）不得误报', r.ok, r.problems.join('；'))
    }
  }

  // ---------------- R12-1：插图链路（markdown → 模型 / docx 插图 / PDF 图片 / 降级不静默） ----------------
  {
    const imgDoc = markdownToDocument('# 图\n\n![红色方块](./a.png)\n\n正文\n')
    ok('R12-1 插图：markdown 独立成行的图片解析成 image 块（旧版变成 "!"+链接）',
      imgDoc.blocks.map(b => b.type).join(',') === 'heading,image,paragraph'
      && imgDoc.blocks[1].alt === '红色方块' && imgDoc.blocks[1].name === './a.png',
      JSON.stringify(imgDoc.blocks.map(b => b.type)))

    const pngRgb = join(OUT, 'r12-img-rgb.png')
    const pngAlpha = join(OUT, 'r12-img-alpha.png')
    writeFileSync(pngRgb, makeTestPng(120, 60, 2, (r, x, y, ch) => {
      const o = x * ch
      const on = x > 10 && x < 110 && y > 10 && y < 50
      r[o] = on ? 200 : 255; r[o + 1] = on ? 40 : 255; r[o + 2] = on ? 40 : 255
    }))
    writeFileSync(pngAlpha, makeTestPng(80, 80, 6, (r, x, y, ch) => {
      const o = x * ch
      r[o] = 20; r[o + 1] = 90; r[o + 2] = 200; r[o + 3] = (x - 40) ** 2 + (y - 40) ** 2 < 900 ? 255 : 0
    }))

    // —— 输入端 → HTML 写出端（image 块往返）——
    const mdImg = join(OUT, 'r12-img.md')
    const htmlImg = join(OUT, 'r12-img.html')
    writeFileSync(mdImg, `# 图册\n\n![红色方块](${pngRgb})\n\n图后正文\n`, 'utf8')
    await call('office_convert', { source: mdImg, target: htmlImg })
    const htmlImgText = readFileSync(htmlImg, 'utf8')
    ok('R12-1 插图：md → html 写出真 <figure><img src>（旧版只是占位文本）',
      /<figure><img src="[^"]*r12-img-rgb\.png" alt="红色方块"><\/figure>/.test(htmlImgText),
      (htmlImgText.match(/<figure[^>]*>.*<\/figure>/) || ['<无>'])[0].slice(0, 120))
    const mdBack = join(OUT, 'r12-img-back.md')
    await call('office_convert', { source: htmlImg, target: mdBack })
    ok('R12-1 插图：html → md 图片块往返（alt 与路径都保留）',
      /!\[红色方块\]\([^)]*r12-img-rgb\.png\)/.test(readFileSync(mdBack, 'utf8')),
      (readFileSync(mdBack, 'utf8').match(/!\[[^\]]*\]\([^)]*\)/) || ['<无>'])[0])

    // —— docx 插图（office_edit）——
    const docxPath = join(OUT, 'r12-img-edit.docx')
    await call('office_create', { path: docxPath, markdown: '# 插图\n\n第一段正文。\n\n第二段正文。\n' })
    const beforeRids = (() => {
      const z0 = openZip(readFileSync(docxPath))
      return (z0.getText('word/_rels/document.xml.rels').match(/Id="([^"]+)"/g) || []).length
    })()
    const edRes = await call('office_edit', {
      path: docxPath,
      operations: [
        { op: 'append_image', path: pngRgb, alt: '红色方块', width: 200 },
        { op: 'insert_image', path: pngAlpha, alt: '蓝色圆', width: 150, after: '第一段正文' },
      ],
    })
    const zImg = openZip(readFileSync(docxPath))
    const relsImg = zImg.getText('word/_rels/document.xml.rels')
    const rids = [...relsImg.matchAll(/Id="([^"]+)"/g)].map(m => m[1])
    const media = zImg.names.filter(n => n.startsWith('word/media/'))
    ok('R12-1 插图：docx zip 里真的多了 media 部件', media.length === 2, media.join(', '))
    ok('R12-1 插图：image 关系 id 唯一且递增（撞号会让 Word 判包损坏）',
      new Set(rids).size === rids.length && rids.length === beforeRids + 2, rids.join(','))
    ok('R12-1 插图：`w:drawing` 写进 document.xml 且引用正确的 rId',
      /<w:drawing>/.test(zImg.getText('word/document.xml')) && /r:embed="rId\d+"/.test(zImg.getText('word/document.xml')))
    ok('R12-1 插图：[Content_Types].xml 补了 png 的 Default',
      /<Default\s+Extension="png"\s+ContentType="image\/png"/.test(zImg.getText('[Content_Types].xml')))
    ok('R12-1 插图：插图后的包结构仍然合规', checkOoxmlPackage(readFileSync(docxPath), 'docx-with-images').ok)
    ok('R12-1 插图：插到锚点段落之后（不是无脑文末）',
      /第一段正文。<\/w:t><\/w:r><\/w:p><w:p><w:r><w:drawing>/.test(zImg.getText('word/document.xml')))
    ok('R12-1 插图：编辑返回里报了操作摘要',
      edRes.summary.some(s => /append_image/.test(s)) && edRes.summary.some(s => /insert_image/.test(s)))
    const docxBack = await textOf(docxPath)
    ok('R12-1 插图：读回时看得见图片位置（`[图片]`，旧版整段丢弃）',
      (docxBack.match(/\[图片\]/g) || []).length === 2, JSON.stringify(docxBack.replace(/\s+/g, ' ').slice(0, 90)))
    let imgErr = ''
    try { await call('office_edit', { path: docxPath, operations: [{ op: 'append_image', path: join(OUT, 'r12-img.md') }] }) } catch (e) { imgErr = String(e.message) }
    ok('R12-1 插图：给个非图片文件 → 明确报错（不静默跳过）', /【插图】/.test(imgErr), imgErr.slice(0, 90))
    let noSrcErr = ''
    try { await call('office_edit', { path: docxPath, operations: [{ op: 'append_image' }] }) } catch (e) { noSrcErr = String(e.message) }
    // 任务四起这一步在**参数层**就被拦下（带 operations[0] 下标定位），比执行期更早、更清楚
    ok('R12-1 插图：缺 path/base64 → 参数层明确报错（带 operations[N] 定位）',
      /需要 "path" 或 "base64" 之一/.test(noSrcErr) && /operations\[0\]/.test(noSrcErr), noSrcErr)
    let anchorMiss = null
    try {
      anchorMiss = await call('office_edit', {
        path: docxPath,
        operations: [{ op: 'insert_image', path: pngRgb, alt: '锚点未命中', after: '这段文字不存在' }],
      })
    } catch (e) { anchorMiss = { error: String(e.message) } }
    ok('R12-1 插图：after 锚点未命中不静默（notice 明说退化为文末）',
      /锚点/.test(String(anchorMiss?.notice || '')), String(anchorMiss?.notice || anchorMiss?.error || '<无>').slice(0, 100))

    // —— PDF 图片写出（直连 writePdf，拿 info 报告）——
    const imgPdf = join(OUT, 'r12-img.pdf')
    const pdfInfo = {}
    const pdfBytes = writePdf(markdownToDocument(
      `# 图片测试\n\n![方](${pngRgb})\n\n图中之后的正文。\n\n![圆](${pngAlpha})\n\n![缺](C:/nope/missing.png)\n`), { info: pdfInfo })
    writeFileSync(imgPdf, pdfBytes)
    const ipRaw = pdfBytes.toString('latin1')
    ok('R12-1 插图：PDF 写出含 /XObject 图片与 Do 绘制算子',
      /\/XObject/.test(ipRaw) && (ipRaw.match(/ Do /g) || []).length === 2)
    ok('R12-1 插图：PNG 走 FlateDecode、带 alpha 的挂 /SMask（不猜、不丢透明）',
      /\/FlateDecode/.test(ipRaw) && /\/SMask/.test(ipRaw) && /\/Subtype\/Image/.test(ipRaw))
    ok('R12-1 插图：读不到的图片不静默（imagesSkipped 有来源+原因）',
      pdfInfo.imagesSkipped?.length === 1 && /missing\.png/.test(pdfInfo.imagesSkipped[0].name),
      JSON.stringify(pdfInfo.imagesSkipped?.[0]))
    const imgRead = await textOf(imgPdf)
    ok('R12-1 插图：含图 PDF 的正文照常可读（图片不影响文字层）', imgRead.includes('图中之后的正文'))
    const convPdf = join(OUT, 'r12-img-conv.pdf')
    const conv = await call('office_convert', { source: mdImg, target: convPdf })
    ok('R12-1 插图：convert → pdf 的 stats.pdfQuality 记账图片数',
      conv.stats?.pdfQuality?.images === 1, JSON.stringify(conv.stats?.pdfQuality))

    // —— 需求 1（第二轮收口）：docx 写出端**内嵌**图片 —— 第六轮的降级用例在此翻转 ——
    // 第六轮因范围隔离只能降级为字面文本；第二轮 `writeDocx` 补上 image 分支后，
    // 原始验收项（"含图 HTML → docx 的 word/media 真有图片"）**已达成**。
    const mdToDocx = join(OUT, 'r12-img-fallback.docx')
    const fb = await call('office_create', { path: mdToDocx, markdown: `# 图\n\n![红色方块](${pngRgb})\n\n正文\n` })
    const zEmb = openZip(readFileSync(mdToDocx))
    ok('R12-1 降级→内嵌：md → docx 真的内嵌图片（验收项已达成：word/media 非空 + stats.imageMedia + notice 尺寸）',
      zEmb.names.filter(n => n.startsWith('word/media/')).length === 1
      && (fb.stats?.imageMedia || []).length === 1
      && /原始 120×60 → /.test(String(fb.notice || '')),
      JSON.stringify({ media: zEmb.names.filter(n => n.startsWith('word/media/')), notice: String(fb.notice || '').slice(0, 80) }))
    ok('R12-1 降级→内嵌：图片的 alt 写进 docx 的 w:drawing 替代文字（信息不丢）',
      /descr="红色方块"/.test(zEmb.getText('word/document.xml')))
    const htmlToDocx = join(OUT, 'r12-img-fromhtml.docx')
    const fb2 = await call('office_convert', { source: htmlImg, target: htmlToDocx })
    const zFb = openZip(readFileSync(htmlToDocx))
    const mediaFb = zFb.names.filter(n => n.startsWith('word/media/'))
    ok('R12-1 降级→内嵌：含图 HTML → docx 的 word/media 真有图片且 rId 唯一递增',
      mediaFb.length === 1
      && (() => { const rids = [...(zFb.getText('word/_rels/document.xml.rels').matchAll(/Id="([^"]+)"/g) || [])].map(m => m[1]); return new Set(rids).size === rids.length && rids.length === 3 })()
      && checkOoxmlPackage(readFileSync(htmlToDocx), 'docx-embedded').ok,
      mediaFb.join(','))
    ok('R12-1 降级→内嵌：该情形同样有 notice/stats 说明（内嵌成功也记账，不静默）',
      (fb2.stats?.imageMedia || []).length === 1 && /内嵌 1 张图片/.test(String(fb2.notice || '')),
      String(fb2.notice || '').slice(0, 90))
    // —— 兜底仍在：**仍未内嵌图片**的写出端（.odt）必须显式降级并双账 ——
    const mdToOdt = join(OUT, 'r12-img-fallback.odt')
    const fbOdt = await call('office_create', { path: mdToOdt, markdown: `# 图\n\n![红色方块](${pngRgb})\n\n正文\n` })
    ok('R12-1 降级：仍未内嵌图片的 .odt 不静默丢图（stats.imageFallback + notice 双账）',
      fbOdt.stats?.imageFallback?.length === 1 && /不内嵌/.test(String(fbOdt.notice || '')),
      JSON.stringify({ fb: fbOdt.stats?.imageFallback?.length, notice: String(fbOdt.notice || '').slice(0, 60) }))
    ok('R12-1 降级：.odt 里图片的 alt 与路径以字面 markdown 保留（不丢信息）',
      (await textOf(mdToOdt)).includes(`![红色方块](${pngRgb})`))

    // —— 需求 4b：`.html` 的 office_edit（核实后本来就通，本轮补回归用例与文档口径）——
    const editHtml = join(OUT, 'r12-edit.html')
    writeFileSync(editHtml, '<html><body><p>旧文案</p></body></html>', 'utf8')
    const edHtml = await call('office_edit', { path: editHtml, operations: [{ op: 'replace_text', find: '旧文案', replace: '新文案' }] })
    const editedHtmlText = readFileSync(editHtml, 'utf8')
    ok('R12-4b html：office_edit 对 .html 做文件级文本替换（不改结构、不重排版）',
      editedHtmlText.includes('新文案') && editedHtmlText.includes('<p>') && !editedHtmlText.includes('旧文案'),
      `format=${edHtml.format}`)
  }

  // ===========================================================================
  // R12-2：AES-256（R5/R6）加密 PDF 读取
  // ===========================================================================
  //
  // 夹具走**自造**路线（零第三方依赖）：`node:crypto` 有全部所需原语（SHA-256/384/512、
  // AES-128-CBC、AES-256-CBC/ECB），按 ISO 32000-2 §7.6.4.3.3 + Algorithm 2.B 自己组装加密
  // PDF 完全可行 —— 不依赖任何外部样本。⚠ 夹具与解析器同源（同一份 `hash2B`），
  // 所以它证明的是"读取链路端到端可用 + 失败路径不猜"，而不是"与第三方实现逐字节一致"。

  /**
   * 造一个 AES-256 加密的单页 PDF。
   * @param {{R?:number, password?:string, corruptPerms?:boolean, content?:string}} o
   */
  function buildAes256Fixture({ R = 6, password = '', corruptPerms = false, content = '' } = {}) {
    const pw = Buffer.from(password, 'utf8')
    const id0 = randomBytes(16)
    const fileKey = randomBytes(32)
    const ZERO = Buffer.alloc(16)
    const sha256 = b => createHash('sha256').update(b).digest()
    // 口径：用户验证哈希带 id0（R6 规范），所有者带 U；R5 只换哈希函数
    const h = (salt, udata) => (R >= 6 ? hash2B(password, salt, udata) : sha256(Buffer.concat([pw, salt, udata])))
    const cbc = (plain, key, iv, pad) => {
      const c = createCipheriv('aes-256-cbc', key, iv)
      if (!pad) c.setAutoPadding(false)
      return Buffer.concat([c.update(plain), c.final()])
    }
    const uVS = randomBytes(8); const uKS = randomBytes(8)
    const U = Buffer.concat([h(uVS, id0), uVS, uKS])
    const UE = cbc(fileKey, h(uKS, id0), ZERO, false)
    const oVS = randomBytes(8); const oKS = randomBytes(8)
    const O = Buffer.concat([h(oVS, U), oVS, oKS])
    const OE = cbc(fileKey, h(oKS, U), ZERO, false)
    const permsPlain = Buffer.alloc(16)
    permsPlain.writeInt32LE(-1, 0)
    permsPlain.writeUInt32LE(0xffffffff, 4)
    permsPlain[8] = corruptPerms ? 0x00 : 0x54          // 'T'
    permsPlain[9] = 0x61; permsPlain[10] = 0x64; permsPlain[11] = 0x62
    const ec = createCipheriv('aes-256-ecb', fileKey, null)
    ec.setAutoPadding(false)
    const Perms = Buffer.concat([ec.update(permsPlain), ec.final()])
    const text = content || `AES-256 R${R} decryption works`
    const plain = Buffer.from(`BT /F1 14 Tf 1 0 0 1 40 760 Tm (${text}) Tj ET`, 'latin1')
    const iv = randomBytes(16)
    const stream = Buffer.concat([iv, cbc(plain, fileKey, iv, true)])   // PDF 的 AES 流用 PKCS#7 填充

    const hex = b => Buffer.from(b).toString('hex')
    const objs = {}
    objs[1] = '<</Type/Catalog/Pages 2 0 R>>'
    objs[2] = '<</Type/Pages/Kids[3 0 R]/Count 1>>'
    objs[3] = '<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Resources<</Font<</F1 4 0 R>>>>/Contents 6 0 R>>'
    objs[4] = '<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>'
    objs[5] = `<</Filter/Standard/V 5/R ${R}/Length 256/P -1`
      + `/O <${hex(O)}>/U <${hex(U)}>/OE <${hex(OE)}>/UE <${hex(UE)}>/Perms <${hex(Perms)}>`
      + `/CF<</StdCF<</CFM/AESV3/Length 32/AuthEvent/DocOpen>>>>/StmF/StdCF/StrF/StdCF>>`
    objs[6] = `<</Length ${stream.length}>>\nstream\n${stream.toString('latin1')}\nendstream`
    let out = '%PDF-1.7\n%\xE9\xEA\xF0\xF1\n'
    const offs = {}
    for (const id of [1, 2, 3, 4, 5, 6]) { offs[id] = Buffer.byteLength(out, 'latin1'); out += `${id} 0 obj\n${objs[id]}\nendobj\n` }
    const xref = Buffer.byteLength(out, 'latin1')
    out += 'xref\n0 7\n0000000000 65535 f \n'
    for (const id of [1, 2, 3, 4, 5, 6]) out += `${String(offs[id]).padStart(10, '0')} 00000 n \n`
    out += `trailer\n<< /Size 7 /Root 1 0 R /Encrypt 5 0 R /ID[<${hex(id0)}> <${hex(id0)}>] >>\nstartxref\n${xref}\n%%EOF\n`
    return Buffer.from(out, 'latin1')
  }

  {
    for (const R of [5, 6]) {
      const p = join(OUT, `r12-aes256-r${R}.pdf`)
      writeFileSync(p, buildAes256Fixture({ R }))
      const before = statSync(p)
      const r = await call('office_read', { path: p, as: 'markdown' })
      const after = statSync(p)
      ok(`R12-2 AES-256：R${R} 空口令透明解密并读出正文`,
        String(r.content).includes(`AES-256 R${R} decryption works`),
        `${String(r.content).replace(/\s+/g, ' ').slice(0, 64)}`)
      ok(`R12-2 AES-256：R${R} 不改写原文件（大小/mtime 不变）`,
        after.size === before.size && after.mtimeMs === before.mtimeMs, `${after.size} 字节`)
    }
    // 真·打开密码：绝不猜测，必须明确报错（且给下一步）
    const locked = join(OUT, 'r12-aes256-locked.pdf')
    writeFileSync(locked, buildAes256Fixture({ R: 6, password: 'secret' }))
    let lockErr = ''
    try { await call('office_read', { path: locked, as: 'markdown' }) } catch (e) { lockErr = String(e.message) }
    ok('R12-2 AES-256：真打开密码 → 明确报错且不猜口令',
      /打开密码|AES-256/.test(lockErr) && !/decryption works/.test(lockErr),
      lockErr.replace(/\s+/g, ' ').slice(0, 140))
    ok('R12-2 AES-256：报错里带可复制的下一步', /另存为未加密|阅读器/.test(lockErr))
    // /Perms 校验：把已知明文破坏掉 → 必须判"密钥不对"，而不是默默返回垃圾
    const badPerms = join(OUT, 'r12-aes256-badperms.pdf')
    writeFileSync(badPerms, buildAes256Fixture({ R: 6, corruptPerms: true }))
    let permsErr = ''
    try { await call('office_read', { path: badPerms, as: 'markdown' }) } catch (e) { permsErr = String(e.message) }
    ok('R12-2 AES-256：/Perms 校验不过 → 报需要密码（绝不接受错误密钥）',
      /打开密码|Perms/.test(permsErr), permsErr.replace(/\s+/g, ' ').slice(0, 120))
    // 截断文件：不许抛裸异常，也不许静默返回垃圾
    const whole = buildAes256Fixture({ R: 6 })
    const cut = join(OUT, 'r12-aes256-truncated.pdf')
    writeFileSync(cut, whole.subarray(0, Math.floor(whole.length * 0.6)))
    let cutErr = ''
    let cutRes = null
    try { cutRes = await call('office_read', { path: cut, as: 'meta' }) } catch (e) { cutErr = String(e.message) }
    const cutOk = cutErr
      ? /页码=|格式=|根因=|下一步=/.test(cutErr)
      : (cutRes?.stats?.textLayerUsable === false || Number(cutRes?.stats?.characters) === 0)
    ok('R12-2 AES-256：截断文件 → 不抛裸异常、也不静默返回垃圾',
      cutOk,
      (cutErr || `pages=${cutRes?.stats?.pages} chars=${cutRes?.stats?.characters} usable=${cutRes?.stats?.textLayerUsable}`).slice(0, 120))
  }

  // ---------------- R12-3：xlsx 公式重算（recalc=true，显式 opt-in） ----------------
  {
    const xlPath = join(OUT, 'r12-recalc.xlsx')
    await call('office_create', {
      path: xlPath,
      workbook: {
        sheets: [{
          name: 'Sheet1',
          rows: [
            ['名称', '分数', '系数'],
            ['甲', 90, 1.1],
            ['乙', 80, 1.2],
            ['丙', 70, 0.9],
            ['合计', { v: '', f: 'SUM(B2:B4)' }, { v: '', f: 'ROUND(AVERAGE(C2:C4),2)' }],
            ['个数', { v: '', f: 'COUNT(B2:B4)' }, { v: '', f: 'COUNTIF(B2:B4,">75")' }],
            ['判定', { v: '', f: 'IF(B5>200,"达标","未达标")' }, { v: '', f: 'MAX(B2:B4)' }],
            ['查表', { v: '', f: 'VLOOKUP("乙",A2:C4,2)' }, { v: '', f: 'VLOOKUP("丁",A2:C4,2)' }],
            ['跨表', { v: '', f: 'SUM(Sheet1!B2:B4)' }, { v: '', f: 'B5&"元"' }],
            ['不支持', { v: '', f: 'LET(1,2)' }, { v: '', f: '1/0' }],
          ],
        }],
      },
    })
    const rec = await call('office_read', { path: xlPath, as: 'meta', recalc: true })
    const rp = rec.stats.recalc
    ok('R12-3 公式重算：只对显式 recalc=true 生效，且统计齐备',
      Boolean(rp) && rp.formulaCells === 12 && rp.evaluated === 9 && rp.unsupported === 1 && rp.errors === 2,
      JSON.stringify({ f: rp?.formulaCells, ok: rp?.evaluated, un: rp?.unsupported, err: rp?.errors }))
    const byCell = Object.fromEntries((rp?.details || []).map(d => [d.cell, d]))
    ok('R12-3 公式重算：未支持的函数显式标 unsupported 且保留原值（绝不猜）',
      /unsupported: LET/.test(byCell.B10?.error || '') && String(byCell.B10?.value).includes('LET'),
      JSON.stringify(byCell.B10))
    ok('R12-3 公式重算：错误值带单元格地址返回（#DIV/0! / #N/A）',
      byCell.C10?.error === '#DIV/0!' && byCell.C8?.error === '#N/A',
      JSON.stringify({ C10: byCell.C10?.error, C8: byCell.C8?.error }))
    const mdCalc = await call('office_read', { path: xlPath, as: 'markdown', recalc: true })
    const rowsCalc = mdCalc.content.split('\n').filter(l => l.startsWith('|')).map(l => l.split('|').map(s => s.trim()))
    const cellAt = (label, col) => (rowsCalc.find(r => r[1] === label) || [])[col]
    ok('R12-3 公式重算：SUM / AVERAGE+ROUND / COUNT / COUNTIF 结果正确',
      cellAt('合计', 2) === '240' && cellAt('合计', 3) === '1.07' && cellAt('个数', 2) === '3' && cellAt('个数', 3) === '2',
      `${cellAt('合计', 2)}/${cellAt('合计', 3)}/${cellAt('个数', 2)}/${cellAt('个数', 3)}`)
    ok('R12-3 公式重算：IF / VLOOKUP / 跨表引用 / 字符串拼接正确',
      cellAt('判定', 2) === '达标' && cellAt('查表', 2) === '80' && cellAt('跨表', 2) === '240' && cellAt('跨表', 3) === '240元',
      `${cellAt('判定', 2)}/${cellAt('查表', 2)}/${cellAt('跨表', 2)}/${cellAt('跨表', 3)}`)
    ok('R12-3 公式重算：错误值在表格里照 Excel 显示（#N/A / #DIV/0!）',
      cellAt('查表', 3) === '#N/A' && cellAt('不支持', 3) === '#DIV/0!',
      `${cellAt('查表', 3)}/${cellAt('不支持', 3)}`)
    const noRecalc = await call('office_read', { path: xlPath, as: 'markdown' })
    ok('R12-3 公式重算：不设 recalc 时返回值与旧版一致（公式原样保留、stats 无 recalc）',
      !noRecalc.stats.recalc && noRecalc.content.includes('=SUM(B2:B4)') && !noRecalc.content.includes('| 240'),
      (noRecalc.content.match(/\| 合计 \|[^\n]*/) || [''])[0])
    const nonSheet = await call('office_read', { path: join(OUT, 'r12-img-back.md'), as: 'meta', recalc: true })
    ok('R12-3 公式重算：非表格文件上 recalc 只记 skipped（不报错也不假装算过）',
      nonSheet.stats.recalc?.skipped !== undefined, JSON.stringify(nonSheet.stats.recalc))
  }

  // ---------------- R12-4：申论稿纸网格（docx w:docGrid） ----------------
  {
    const gridPath = join(OUT, 'r12-grid.docx')
    await call('office_create', { path: gridPath, markdown: '# 申论作答\n\n第一段。\n', grid: '20x25' })
    const zg = openZip(readFileSync(gridPath))
    const gridDoc = zg.getText('word/document.xml')
    ok('R12-4 稿纸网格：writeDocx 写出 `<w:docGrid type="linesAndChars">`',
      /<w:docGrid w:type="linesAndChars" w:linePitch="\d+" w:charSpace="\d+"\/>/.test(gridDoc),
      (gridDoc.match(/<w:docGrid[^>]*\/>/) || ['<无>'])[0])
    ok('R12-4 稿纸网格：20 字×25 行的换算符合 A4 可用区',
      /w:linePitch="558"/.test(gridDoc) && /w:charSpace="211"/.test(gridDoc))
    let gridErr = ''
    try { await call('office_create', { path: join(OUT, 'r12-grid.odt'), markdown: '# T', grid: '20x25' }) } catch (e) { gridErr = String(e.message) }
    ok('R12-4 稿纸网格：非 docx 目标显式报错（不静默忽略）', /【稿纸网格】/.test(gridErr))
    let gridBad = ''
    try { await call('office_create', { path: join(OUT, 'r12-grid2.docx'), markdown: '# T', grid: 'abc' }) } catch (e) { gridBad = String(e.message) }
    ok('R12-4 稿纸网格：非法取值显式报错并给可照抄的写法', /无法识别/.test(gridBad) && /20x25/.test(gridBad))
    ok('R12-4 稿纸网格：包结构仍合规', checkOoxmlPackage(readFileSync(gridPath), 'grid').ok)
    ok('R12-4 稿纸网格：读回内容不受影响', (await textOf(gridPath)).includes('第一段'))
  }

  // ===========================================================================
  // 第二轮收口批次（R13-）：docx 写出端插图 / 图片解码面 / 公式扩展 / AES 交叉验证
  // ===========================================================================

  // ---------------- R13-1：docx 写出端 image 分支（需求 1 / 4c / 4e） ----------------
  {
    const r13Png = join(OUT, 'r13-img.png')
    writeFileSync(r13Png, makeTestPng(120, 60, 2, (r, x, y, ch) => {
      const o = x * ch
      const on = x > 10 && x < 110 && y > 10 && y < 50
      r[o] = on ? 200 : 255; r[o + 1] = on ? 40 : 255; r[o + 2] = on ? 40 : 255
    }))

    // 无图文档必须"不多写任何东西"（第六轮产物逐字节口径）
    const plainDocx = join(OUT, 'r13-plain.docx')
    await call('office_create', { path: plainDocx, markdown: '# 标题\n\n正文\n' })
    const zp = openZip(readFileSync(plainDocx))
    ok('R13-1 docx：无图文档不多写 media / 图片关系 / 图片内容类型 / w:drawing（逐字节口径）',
      zp.names.filter(n => n.startsWith('word/media/')).length === 0
      && !/<Relationship[^>]*Type="[^"]*\/image"/.test(zp.getText('word/_rels/document.xml.rels'))
      && !/<Default[^>]*Extension="(png|jpeg|jpg|gif|bmp)"/.test(zp.getText('[Content_Types].xml'))
      && !/<w:drawing>/.test(zp.getText('word/document.xml')),
      zp.names.join(','))

    // 尺寸策略（需求 4e）：显式规则 + 换算过程可回放
    const sizeAuto = imageSizeFor(1200, 800)
    const sizeExplicit = imageSizeFor(120, 60, 200)
    ok('R13-1 尺寸策略：省略 width → 原图像素 × 72/96；超 A4 可用宽等比缩到 451.3pt',
      Math.abs(sizeAuto.naturalWidthPt - 900) < 0.01 && sizeAuto.capped === true
      && Math.abs(sizeAuto.widthPt - DOCX_USABLE_WIDTH_PT) < 0.01
      && Math.abs(sizeAuto.heightPt - 300.87) < 0.05,
      JSON.stringify(sizeAuto))
    ok('R13-1 尺寸策略：显式 width 优先、高度按纵横比、rule 可回放',
      sizeExplicit.widthPt === 200 && sizeExplicit.heightPt === 100 && sizeExplicit.rule === 'width 参数',
      JSON.stringify(sizeExplicit))

    // 媒体去重（需求 4c）：同一张图两次 → 只 1 个部件 / 1 条关系 / imageReused 记账
    const dupDocx = join(OUT, 'r13-dup.docx')
    const dupRes = await call('office_create', { path: dupDocx, markdown: `# 去重\n\n![甲](${r13Png})\n\n![乙](${r13Png})\n` })
    const zd = openZip(readFileSync(dupDocx))
    const dupRids = [...(zd.getText('word/_rels/document.xml.rels').matchAll(/Id="([^"]+)"/g) || [])].map(m => m[1])
    ok('R13-1 去重：同一张图出现两次只写 1 个媒体部件 + 1 条图片关系（内容 SHA-256 去重）',
      zd.names.filter(n => n.startsWith('word/media/')).length === 1
      && dupRids.length === 3 && new Set(dupRids).size === 3
      && (dupRes.stats?.imageReused || []).length === 1 && /复用/.test(String(dupRes.notice || '')),
      JSON.stringify({ rids: dupRids, reused: dupRes.stats?.imageReused?.length }))
    ok('R13-1 docx：含图产物包结构体检通过，且两个 drawing 都在（复用关系但仍画两次）',
      checkOoxmlPackage(readFileSync(dupDocx), 'dup').ok
      && (zd.getText('word/document.xml').match(/<w:drawing>/g) || []).length === 2)

    // 失败路径：读不到 / 格式不认识 → 记账 + 字面文本，绝不静默丢图
    const missDocx = join(OUT, 'r13-miss.docx')
    const missRes = await call('office_create', { path: missDocx, markdown: '# 缺图\n\n![没有](C:/nope/r13-missing.png)\n\n正文\n' })
    ok('R13-1 失败路径：读不到的图片不静默丢图（imagesSkipped + notice + 字面文本保留）',
      (missRes.stats?.imagesSkipped || []).length === 1 && /未能内嵌/.test(String(missRes.notice || ''))
      && (await textOf(missDocx)).includes('![没有](C:/nope/r13-missing.png)'),
      JSON.stringify({ skipped: missRes.stats?.imagesSkipped?.[0]?.reason?.slice(0, 40), notice: String(missRes.notice || '').slice(0, 60) }))
    const badImgPath = join(OUT, 'r13-bad.png')
    writeFileSync(badImgPath, Buffer.from('not a png at all'))
    const badRes = await call('office_create', { path: join(OUT, 'r13-bad.docx'), markdown: `# 坏图\n\n![坏](${badImgPath})\n` })
    ok('R13-1 失败路径：不认识的图片格式同样记账 + 保留字面文本（不抛裸异常）',
      (badRes.stats?.imagesSkipped || []).length === 1 && /不认识的图片格式/.test(badRes.stats.imagesSkipped[0].reason),
      JSON.stringify(badRes.stats?.imagesSkipped?.[0]))

    // office_edit：append_markdown 里的图片真的落进 media；再插同图复用既有部件
    const editDocxPath = join(OUT, 'r13-edit.docx')
    await call('office_create', { path: editDocxPath, markdown: '# 编辑\n\n正文\n' })
    const edRes2 = await call('office_edit', { path: editDocxPath, operations: [{ op: 'append_markdown', markdown: `![追加](${r13Png})\n` }] })
    const ze = openZip(readFileSync(editDocxPath))
    ok('R13-1 office_edit：append_markdown 里的图片真的写进 word/media + 关系 + w:drawing',
      ze.names.filter(n => n.startsWith('word/media/')).length === 1
      && /<w:drawing>/.test(ze.getText('word/document.xml'))
      && (edRes2.stats?.imageMedia || []).length === 1
      && checkOoxmlPackage(readFileSync(editDocxPath), 'edit-image').ok,
      JSON.stringify({ media: ze.names.filter(n => n.startsWith('word/media/')), notice: String(edRes2.notice || '').slice(0, 60) }))
    const edRes3 = await call('office_edit', { path: editDocxPath, operations: [{ op: 'append_image', path: r13Png, alt: '再来一张' }] })
    const ze3 = openZip(readFileSync(editDocxPath))
    ok('R13-1 office_edit：再插同一张图复用既有媒体部件（不重写部件、不新增关系 id）',
      ze3.names.filter(n => n.startsWith('word/media/')).length === 1
      && (edRes3.stats?.imageReused || []).length === 1
      && [...(ze3.getText('word/_rels/document.xml.rels').matchAll(/Id="([^"]+)"/g) || [])].length === 3,
      JSON.stringify(edRes3.stats?.imageReused))

    // 原生开箱：Word 16.0 能打开含图 docx，且能看到 2 个内联图形
    const R13_WORD_PS = `
$ErrorActionPreference = 'Continue'
$src = $env:DSH_OFFICE_NATIVE_SRC
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300
try { $w = New-Object -ComObject Word.Application } catch { 'NATIVE_NO_WORD ' + $_.Exception.Message; exit 0 }
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($src, $false, $true)
  $n = $d.Paragraphs.Count
  $shapes = $d.InlineShapes.Count
  $d.Close(0)
  "NATIVE_OPEN_OK paragraphs=$n inlineShapes=$shapes"
} catch {
  'NATIVE_OPEN_FAIL ' + $_.Exception.Message
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
`
    if (process.platform !== 'win32') {
      ok('R13-1 开箱：Word 打开含图 docx（非 Windows → 跳过）', true, `platform=${process.platform}`)
    } else {
      const env = { ...process.env, DSH_OFFICE_NATIVE_SRC: editDocxPath }
      delete env.DSH_OFFICE_NATIVE_PDF
      const r = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', R13_WORD_PS],
        { encoding: 'utf8', windowsHide: true, timeout: 240000, env })
      const t = `${(r && r.stdout) || ''}${(r && r.stderr) || ''}`.trim()
      if (!t || /NATIVE_NO_WORD|Retrieving the COM class factory|80040154|not registered/i.test(t)) {
        ok('R13-1 开箱：Word 打开含图 docx（Word 不可用 → 跳过）', true, t.replace(/\s+/g, ' ').slice(0, 120))
      } else {
        ok('R13-1 开箱：Word 16.0 能打开含图的 .docx 并识别出 2 个内联图形',
          /NATIVE_OPEN_OK/.test(t) && /inlineShapes=2/.test(t), t.replace(/\s+/g, ' ').slice(0, 120))
      }
    }
  }

  // ---------------- R13-4：图片解码面（需求 4a/4b/4c/4f） ----------------
  {
    const idxPng = makeIndexedPng(4, 2, [[255, 0, 0], [0, 255, 0], [0, 0, 255], [255, 255, 255]], [0, 1, 2, 3, 3, 2, 1, 0], [255, 128, 255, 255])
    const rawIdx = pngToRaw(idxPng)
    ok('R13-4 解码面：调色板 PNG（colorType 3）用 PLTE + tRNS 展开成 RGB + alpha',
      Boolean(rawIdx) && rawIdx.width === 4 && rawIdx.height === 2 && Boolean(rawIdx.alpha)
      && rawIdx.data[0] === 255 && rawIdx.data[1] === 0 && rawIdx.data[2] === 0
      && rawIdx.alpha[1] === 128 && rawIdx.alpha[2] === 255,
      JSON.stringify({ w: rawIdx?.width, h: rawIdx?.height, a: rawIdx?.alpha ? [...rawIdx.alpha] : null }))
    const idxPngPath = join(OUT, 'r13-idx.png')
    writeFileSync(idxPngPath, makeIndexedPng(8, 8, [[200, 30, 30], [255, 255, 255]], Array.from({ length: 64 }, (_, i) => (i % 8 < 4 ? 0 : 1))))
    const infoIdx = {}
    const pdfIdx = writePdf(markdownToDocument(`# 索引图\n\n![索引](${idxPngPath})\n\n后面正文\n`), { info: infoIdx })
    const idxPdfPath = join(OUT, 'r13-idx.pdf')
    writeFileSync(idxPdfPath, pdfIdx)
    ok('R13-4 PDF：调色板 PNG 走 FlateDecode 内嵌（不再整张 skipped），正文仍可读',
      /\/XObject/.test(pdfIdx.toString('latin1')) && (infoIdx.imagesSkipped || []).length === 0
      && infoIdx.images?.length === 1 && (await textOf(idxPdfPath)).includes('后面正文'),
      JSON.stringify({ skipped: infoIdx.imagesSkipped, images: infoIdx.images }))

    const bmp = makeTestBmp24(6, 4, (x) => (x < 3 ? [255, 0, 0] : [0, 0, 255]))
    const rawBmp = bmpToRaw(bmp)
    ok('R13-4 解码面：24-bit BMP → RGB 原始采样（自下而上行序已还原）',
      Boolean(rawBmp) && rawBmp.width === 6 && rawBmp.height === 4
      && rawBmp.data[0] === 255 && rawBmp.data[1] === 0 && rawBmp.data[2] === 0
      && rawBmp.data[5 * 3] === 0 && rawBmp.data[5 * 3 + 2] === 255,
      JSON.stringify({ w: rawBmp?.width, h: rawBmp?.height, first: rawBmp ? [...rawBmp.data.subarray(0, 6)] : null }))
    const bmpPath = join(OUT, 'r13-img.bmp')
    writeFileSync(bmpPath, bmp)
    const infoBmp = {}
    writePdf(markdownToDocument(`# BMP\n\n![位图](${bmpPath})\n`), { info: infoBmp })
    ok('R13-4 PDF：BMP 走 FlateDecode 内嵌且不记账跳过',
      (infoBmp.imagesSkipped || []).length === 0 && infoBmp.images?.length === 1,
      JSON.stringify(infoBmp.imagesSkipped))

    const gif = makeTestGif(4, 2, [[255, 0, 0], [0, 255, 0], [0, 0, 255]], [0, 1, 2, 0], 1)
    const rawGif = gifToRaw(gif)
    ok('R13-4 解码面：GIF 首帧 LZW 解码 + 透明索引（索引 1 处 alpha=0）',
      Boolean(rawGif) && rawGif.width === 4 && rawGif.height === 2 && Boolean(rawGif.alpha)
      && rawGif.alpha[0] === 255 && rawGif.alpha[1] === 0
      && rawGif.data[0] === 255 && rawGif.data[1] === 0,
      JSON.stringify({ a: rawGif?.alpha ? [...rawGif.alpha] : null }))
    const gifPath = join(OUT, 'r13-img.gif')
    writeFileSync(gifPath, gif)
    const infoGif = {}
    writePdf(markdownToDocument(`# GIF\n\n![动图](${gifPath})\n`), { info: infoGif })
    ok('R13-4 PDF：GIF 首帧内嵌（透明索引挂 /SMask）', (infoGif.imagesSkipped || []).length === 0 && infoGif.images?.length === 1)

    const badBmp = Buffer.concat([Buffer.from('BM'), randomBytes(60)])
    const badGif = Buffer.concat([Buffer.from('GIF89a'), randomBytes(40)])
    ok('R13-4 坏字节：坏 BMP / 坏 GIF 解不出采样时返回 null（调用方记账，绝不抛裸异常）',
      bmpToRaw(badBmp) === null && bmpToRaw(Buffer.from('nope')) === null
      && gifToRaw(badGif) === null && imageToRaw(Buffer.from('xx')) === null
      && bmpToRaw(Buffer.alloc(4)) === null)

    const infoDup = {}
    const pdfDup = writePdf(markdownToDocument(`# 去重\n\n![甲](${idxPngPath})\n\n![乙](${idxPngPath})\n`), { info: infoDup })
    ok('R13-4 PDF 去重：同一张图出现两次只写 1 个 /XObject（复用资源名），Do 仍画两次',
      infoDup.images?.length === 1 && (infoDup.imageReused || []).length === 1
      && (pdfDup.toString('latin1').match(/ Do /g) || []).length === 2,
      JSON.stringify({ images: infoDup.images, reused: infoDup.imageReused?.length }))

    const infoLeft = {}
    const pdfLeft = writePdf({ kind: 'document', meta: {}, blocks: [{ type: 'image', alt: '甲', name: idxPngPath }] }, { info: infoLeft })
    const infoCenter = {}
    const pdfCenter = writePdf({ kind: 'document', meta: {}, blocks: [{ type: 'image', alt: '甲', name: idxPngPath, align: 'center' }] }, { info: infoCenter })
    const cmOf = buf => (buf.toString('latin1').match(/q [\d.]+ 0 0 [\d.]+ -?[\d.]+ -?[\d.]+ cm \/Im1 Do Q/) || [''])[0]
    const leftCm = cmOf(pdfLeft)
    const centerCm = cmOf(pdfCenter)
    const cmX = s => Number(s.split(' ')[5])
    ok('R13-4 PDF 对齐：align=center 把图片画到可用区中间；缺省仍是 marginX（旧字节不变）',
      Boolean(leftCm) && Boolean(centerCm) && leftCm !== centerCm && Math.abs(cmX(centerCm) - cmX(leftCm)) > 50,
      `${leftCm} | ${centerCm}`)
  }

  // ---------------- R13-3：公式子集扩展（需求 3b/3c/3d） ----------------
  {
    const fx = join(OUT, 'r13-formula.xlsx')
    await call('office_create', {
      path: fx,
      workbook: {
        sheets: [{
          name: 'Sheet1',
          rows: [
            ['名称', '值', '类'],
            ['甲', 90, 'A'],
            ['乙', 80, 'B'],
            ['丙', 70, 'A'],
            ['丁', 60, 'A'],
            ['戊', 50, 'B'],
            ['SUMIF_A', { v: '', f: 'SUMIF(C2:C6,"A",B2:B6)' }, ''],
            ['AVERAGEIF_A', { v: '', f: 'AVERAGEIF(C2:C6,"A",B2:B6)' }, ''],
            ['COUNTIFS_AB', { v: '', f: 'COUNTIFS(C2:C6,"A",B2:B6,">75")' }, ''],
            ['SUM区域', { v: '', f: 'SUM(B2:B6)' }, ''],
            ['SUM直接', { v: '', f: 'SUM("3",TRUE)' }, ''],
            ['MAX区域', { v: '', f: 'MAX(B2:B6)' }, ''],
            ['LEFT', { v: '', f: 'LEFT("abcdef",3)' }, ''],
            ['RIGHT', { v: '', f: 'RIGHT("abcdef",2)' }, ''],
            ['MID', { v: '', f: 'MID("abcdef",2,3)' }, ''],
            ['LEN', { v: '', f: 'LEN("中文ab")' }, ''],
            ['TRIM', { v: '', f: 'TRIM("  a   b  ")' }, ''],
            ['UPPER', { v: '', f: 'UPPER("ab")' }, ''],
            ['LOWER', { v: '', f: 'LOWER("AB")' }, ''],
            ['CONCAT', { v: '', f: 'CONCAT("a",1,TRUE)' }, ''],
            ['CONCATENATE', { v: '', f: 'CONCATENATE("a","b")' }, ''],
            ['TEXT', { v: '', f: 'TEXT(1234.5,"#,##0.00")' }, ''],
            ['TEXT日期', { v: '', f: 'TEXT(45000,"yyyy-mm-dd")' }, ''],
            ['VALUE', { v: '', f: 'VALUE("1,234.5")' }, ''],
            ['VALUE坏', { v: '', f: 'VALUE("abc")' }, ''],
            ['整列', { v: '', f: 'SUM(A:A)' }, ''],
            ['外部', { v: '', f: '[Book1]Sheet1!A1' }, ''],
            ['xlfn', { v: '', f: '_xlfn.CONCAT("a","b")' }, ''],
            ['跨表未加载', { v: '', f: 'NoSheet!A1' }, ''],
          ],
        }],
      },
    })
    const fr = await call('office_read', { path: fx, as: 'meta', recalc: true })
    const fp = fr.stats.recalc
    const fcell = Object.fromEntries((fp?.details || []).map(d => [d.cell, d]))
    const mdFx = await call('office_read', { path: fx, as: 'markdown', recalc: true })
    const rowsFx = mdFx.content.split('\n').filter(l => l.startsWith('|')).map(l => l.split('|').map(s => s.trim()))
    const cellFx = label => (rowsFx.find(r => r[1] === label) || [])[2]
    ok('R13-3 公式：SUMIF / AVERAGEIF / COUNTIFS 正确（判据与区域分开，不摊平）',
      cellFx('SUMIF_A') === '220' && Math.abs(Number(cellFx('AVERAGEIF_A')) - 220 / 3) < 1e-9 && cellFx('COUNTIFS_AB') === '1',
      `${cellFx('SUMIF_A')}/${cellFx('AVERAGEIF_A')}/${cellFx('COUNTIFS_AB')}`)
    ok('R13-3 公式：SUM 对区域求和、直接参数换算（区域文本/布尔口径见下面直连用例）',
      cellFx('SUM区域') === '350' && cellFx('SUM直接') === '4' && cellFx('MAX区域') === '90',
      `${cellFx('SUM区域')}/${cellFx('SUM直接')}/${cellFx('MAX区域')}`)
    // —— 直连求值器：xlsx writer 会把纯数字串写成数字单元格，所以"区域里的文本/布尔"这条
    //    必须用内存 model 精确控制值类型（真 Excel 口径：区域忽略文本与布尔、直接参数参与换算）——
    const { recalcWorkbook } = await import('./formula.js')
    const rr = recalcWorkbook({
      kind: 'workbook',
      sheets: [{
        name: 'Sheet1',
        rows: [
          ['名称', '值', '类'],
          ['甲', 90, 'A'],
          ['乙', 80, 'B'],
          ['丙', 70, 'A'],
          ['丁', '90', 'A'],
          ['戊', true, 'B'],
          ['SUMIF_A', { v: '', f: 'SUMIF(C2:C6,"A",B2:B6)' }],
          ['AVERAGEIF_A', { v: '', f: 'AVERAGEIF(C2:C6,"A",B2:B6)' }],
          ['COUNTIFS_AB', { v: '', f: 'COUNTIFS(C2:C6,"A",B2:B6,">75")' }],
          ['SUM区域', { v: '', f: 'SUM(B2:B6)' }],
          ['SUM直接', { v: '', f: 'SUM("3",TRUE)' }],
          ['COUNT区域', { v: '', f: 'COUNT(B2:B6)' }],
          ['MAX区域', { v: '', f: 'MAX(B2:B6)' }],
          ['COUNTIF通配', { v: '', f: 'COUNTIF(C2:C6,"?")' }],
        ],
      }],
    })
    const dv = (r, c) => rr.model.sheets[0].rows[r][c].v
    ok('R13-3 直连（真 Excel 口径）：区域里的文本/布尔被忽略 —— SUM=240、COUNT=3、MAX=90',
      dv(9, 1) === 240 && dv(11, 1) === 3 && dv(12, 1) === 90,
      `${dv(9, 1)}/${dv(11, 1)}/${dv(12, 1)}`)
    ok('R13-3 直连：SUMIF / AVERAGEIF 的求和区同样忽略文本/布尔（160 / 80）',
      dv(6, 1) === 160 && dv(7, 1) === 80, `${dv(6, 1)}/${dv(7, 1)}`)
    ok('R13-3 直连：数值条件不命中文本单元格（COUNTIFS ">75" 只数数字 → 1）',
      dv(8, 1) === 1, String(dv(8, 1)))
    ok('R13-3 直连：直接参数里的文本/布尔参与换算（SUM("3",TRUE)=4）', dv(10, 1) === 4, String(dv(10, 1)))
    ok('R13-3 直连：COUNTIF 的 ? 通配符命中全部单字符文本（5）', dv(13, 1) === 5, String(dv(13, 1)))
    ok('R13-3 公式：文本函数族 LEFT/RIGHT/MID/LEN/TRIM/UPPER/LOWER/CONCAT/CONCATENATE 正确',
      cellFx('LEFT') === 'abc' && cellFx('RIGHT') === 'ef' && cellFx('MID') === 'bcd'
      && cellFx('LEN') === '4' && cellFx('TRIM') === 'a b'
      && cellFx('UPPER') === 'AB' && cellFx('LOWER') === 'ab'
      && cellFx('CONCAT') === 'a1TRUE' && cellFx('CONCATENATE') === 'ab',
      [cellFx('LEFT'), cellFx('RIGHT'), cellFx('MID'), cellFx('LEN'), cellFx('TRIM'), cellFx('CONCAT')].join('/'))
    ok('R13-3 公式：TEXT 数字格式子集生效、日期格式显式 unsupported（绝不猜）',
      cellFx('TEXT') === '1,234.50' && /unsupported/.test(fcell.B23?.error || ''),
      `${cellFx('TEXT')} / ${fcell.B23?.error}`)
    ok('R13-3 公式：VALUE 解析千分位、坏输入报 #VALUE!（带单元格地址）',
      cellFx('VALUE') === '1234.5' && fcell.B25?.error === '#VALUE!',
      `${cellFx('VALUE')} / ${fcell.B25?.error}`)
    ok('R13-3 公式：整列引用 / 外部工作簿 / _xlfn 前缀 / 未加载工作表全部显式记账（带地址）',
      /整列引用/.test(fcell.B26?.error || '') && /不支持的引用/.test(fcell.B27?.error || '')
      && /unsupported: _XLFN.CONCAT/.test(fcell.B28?.error || '') && fcell.B29?.error === '#REF!',
      [fcell.B26?.error, fcell.B27?.error, fcell.B28?.error, fcell.B29?.error].join(' | ').slice(0, 160))
    ok('R13-3 公式：unsupportedFunctions 汇总与 supportedFunctions 清单同步',
      (fp?.supportedFunctions || []).includes('SUMIF') && (fp?.supportedFunctions || []).includes('CONCATENATE')
      && (fp?.unsupportedTokens || []).some(t => /整列引用|不支持的引用|_XLFN/.test(t)),
      JSON.stringify(fp?.unsupportedTokens))

    // 循环引用：A→B→A 必须显式 #CIRC! 且不挂死
    const circPath = join(OUT, 'r13-circ.xlsx')
    await call('office_create', {
      path: circPath,
      workbook: { sheets: [{ name: 'Sheet1', rows: [['a', { v: '', f: 'B2' }], ['b', { v: '', f: 'B1' }]] }] },
    })
    const circRes = await call('office_read', { path: circPath, as: 'meta', recalc: true })
    const circCells = Object.fromEntries((circRes.stats.recalc?.details || []).map(d => [d.cell, d.error]))
    ok('R13-3 失败口径：循环引用显式 #CIRC!（带地址、不挂死）',
      circCells.B1 === '#CIRC!' && circCells.B2 === '#CIRC!', JSON.stringify(circCells))

    // 大表性能 + 护栏（需求 3d）
    const bigRows = [['序号', '值']]
    for (let i = 0; i < 5000; i++) bigRows.push([i + 1, i % 2 ? { v: '', f: `A${i + 2}*2` } : i])
    const bigPath = join(OUT, 'r13-big.xlsx')
    await call('office_create', { path: bigPath, workbook: { sheets: [{ name: 'Sheet1', rows: bigRows }] } })
    const t0 = Date.now()
    const bigRes = await call('office_read', { path: bigPath, as: 'meta', recalc: true })
    const bigMs = Date.now() - t0
    ok('R13-3 性能：1 万单元格量级（5000 行 / 2500 个公式）重算在实测量级内跑完',
      bigRes.stats.recalc?.evaluated === 2500 && bigMs < 120000,
      `${bigMs}ms / evaluated=${bigRes.stats.recalc?.evaluated}`)
    process.env.DSH_OFFICE_RECALC_MAX_CELLS = '100'
    const bigRes2 = await call('office_read', { path: bigPath, as: 'meta', recalc: true })
    delete process.env.DSH_OFFICE_RECALC_MAX_CELLS
    ok('R13-3 护栏：DSH_OFFICE_RECALC_MAX_CELLS 超限 → 显式 skipped + 原值保留（不静默截断）',
      /跳过重算/.test(bigRes2.stats.recalc?.skipped || '') && bigRes2.stats.recalc?.evaluated === 0,
      String(bigRes2.stats.recalc?.skipped || '').slice(0, 80))
    const bigRes3 = await call('office_read', { path: bigPath, as: 'meta', recalc: true })
    ok('R13-3 护栏：删掉环境变量后行为回到无上限（opt-in，不设时与旧版一致）',
      bigRes3.stats.recalc?.evaluated === 2500)
  }

  // ---------------- R13-2：AES-128 对照夹具 + 交叉验证结论（需求 2） ----------------
  {
    const AES_PAD = Buffer.from([
      0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
      0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
    ])
    const md5 = b => createHash('md5').update(b).digest()
    const padPw = s => {
      const p = Buffer.from(String(s ?? ''), 'latin1')
      return p.length >= 32 ? p.subarray(0, 32) : Buffer.concat([p, AES_PAD.subarray(0, 32 - p.length)])
    }
    /**
     * 造一个 **R4/V4 + AESV2** 加密单页 PDF —— 与 R12-2 的 AES-256 夹具**同一骨架**
     * （同 writePdf 风格 xref/trailer/`/Encrypt` 位置），只是换成 Algorithm 3/5 + AES-128-CBC。
     * 用途：让支持 AES-128 的第三方实现（WinRT）来判定"骨架本身对不对"。
     */
    function buildAes128Fixture({ password = '', content = '' } = {}) {
      const R = 4
      const P = -1
      const id0 = randomBytes(16)
      const up = padPw(password)
      const op = padPw(password)
      let oKey = md5(op)
      for (let i = 0; i < 50; i++) oKey = md5(oKey.subarray(0, 16))
      oKey = oKey.subarray(0, 16)
      let O = rc4(oKey, up)
      for (let i = 1; i <= 19; i++) O = rc4(Buffer.from(oKey.map(b => b ^ i)), O)
      const pbuf = Buffer.alloc(4)
      pbuf.writeInt32LE(P)
      let fk = md5(Buffer.concat([up, O, pbuf, id0]))
      for (let i = 0; i < 50; i++) fk = md5(fk.subarray(0, 16))
      const fileKey = fk.subarray(0, 16)
      let x = md5(Buffer.concat([AES_PAD, id0]))
      x = rc4(fileKey, x)
      for (let i = 1; i <= 19; i++) x = rc4(Buffer.from(fileKey.map(b => b ^ i)), x)
      const U = Buffer.concat([x.subarray(0, 16), randomBytes(16)])
      const objKey = (num, gen) => md5(Buffer.concat([
        fileKey,
        Buffer.from([num & 0xff, (num >> 8) & 0xff, (num >> 16) & 0xff, gen & 0xff, (gen >> 8) & 0xff]),
        Buffer.from('sAlT', 'latin1'),
      ])).subarray(0, 16)
      const encStream = (num, plain) => {
        const iv = randomBytes(16)
        const c = createCipheriv('aes-128-cbc', objKey(num, 0), iv)
        return Buffer.concat([iv, c.update(plain), c.final()])
      }
      const text = content || 'AES-128 R4 decryption works'
      const plain = Buffer.from(`BT /F1 14 Tf 1 0 0 1 40 760 Tm (${text}) Tj ET`, 'latin1')
      const stream = encStream(6, plain)
      const objs = {}
      objs[1] = '<</Type/Catalog/Pages 2 0 R>>'
      objs[2] = '<</Type/Pages/Kids[3 0 R]/Count 1>>'
      objs[3] = '<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Resources<</Font<</F1 4 0 R>>>>/Contents 6 0 R>>'
      objs[4] = '<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>'
      objs[5] = `<</Filter/Standard/V 4/R 4/Length 128/P ${P}`
        + `/O <${O.toString('hex')}>/U <${U.toString('hex')}>`
        + '/CF<</StdCF<</CFM/AESV2/Length 16>>>>/StmF/StdCF/StrF/StdCF>>'
      objs[6] = `<</Length ${stream.length}>>\nstream\n${stream.toString('latin1')}\nendstream`
      let out = '%PDF-1.6\n%\xE9\xEA\xF0\xF1\n'
      const offs = {}
      for (const id of [1, 2, 3, 4, 5, 6]) { offs[id] = Buffer.byteLength(out, 'latin1'); out += `${id} 0 obj\n${objs[id]}\nendobj\n` }
      const xref = Buffer.byteLength(out, 'latin1')
      out += 'xref\n0 7\n0000000000 65535 f \n'
      for (const id of [1, 2, 3, 4, 5, 6]) out += `${String(offs[id]).padStart(10, '0')} 00000 n \n`
      out += `trailer\n<< /Size 7 /Root 1 0 R /Encrypt 5 0 R /ID[<${id0.toString('hex')}> <${id0.toString('hex')}>] >>\nstartxref\n${xref}\n%%EOF\n`
      return Buffer.from(out, 'latin1')
    }

    const aes128 = join(OUT, 'r13-aes128-r4.pdf')
    writeFileSync(aes128, buildAes128Fixture({}))
    const before128 = statSync(aes128)
    const r128 = await call('office_read', { path: aes128, as: 'markdown' })
    const after128 = statSync(aes128)
    ok('R13-2 AES-128 对照：自造 R4/V4+AESV2 夹具能被读出正文（加密方向与解析器一致）',
      String(r128.content).includes('AES-128 R4 decryption works'),
      String(r128.content).replace(/\s+/g, ' ').slice(0, 64))
    ok('R13-2 AES-128 对照：透明解密不改写原文件（大小/mtime 不变）',
      after128.size === before128.size && after128.mtimeMs === before128.mtimeMs, `${after128.size} 字节`)
    const locked128 = join(OUT, 'r13-aes128-locked.pdf')
    writeFileSync(locked128, buildAes128Fixture({ password: 'secret' }))
    let lockErr128 = ''
    try { await call('office_read', { path: locked128, as: 'markdown' }) } catch (e) { lockErr128 = String(e.message) }
    ok('R13-2 AES-128 对照：真口令 → 明确报需要密码且不猜口令（口径与 R2/R4 一致）',
      /打开密码/.test(lockErr128) && !/decryption works/.test(lockErr128), lockErr128.replace(/\s+/g, ' ').slice(0, 120))

    // —— 判别实验：同一骨架的 AES-128，交给微软实现（WinRT）打开 ——
    if (process.platform !== 'win32') {
      ok('R13-2 判别实验：WinRT 打开 AES-128 夹具（非 Windows → 跳过）', true, `platform=${process.platform}`)
    } else {
      // R18 任务 A（连带修复）：这两条用例过去必 FAIL，根因**不是**加密骨架，而是调用方式：
      //   ① 探针 PDF 写在工作区（本机 WinRT 的 StorageFile 只读得到 %TEMP% 内的文件）；
      //   ② 测试自己 spawnSync + 管道 stdio，受限沙箱下直接 EPERM，stdout/stderr 全空。
      // 现在与实现**同源**：探针 PDF 落到 pdfGateStagingRoot()（渲染暂存目录），
      // 渲染统一走 runRenderScript()（它自带"管道被拒 → stdio:'ignore' 重来"的兜底）。
      // 断言一个字没放宽：探针本身有效 = 明文 PDF 必须真的渲染出 PNG。
      const probeStage = join(pdfGateStagingRoot(), `r13-probe-${process.pid}-${Date.now()}`)
      mkdirSync(probeStage, { recursive: true })
      const aesStage = join(probeStage, 'aes128.pdf')
      const plainStage = join(probeStage, 'plain.pdf')
      writeFileSync(aesStage, buildAes128Fixture({}))
      writeFileSync(plainStage, writePdf(markdownToDocument('# 明文对照\n'), {}))
      const renderDir128 = join(OUT, 'r13-winrt-aes128')
      const renderDirPlain = join(OUT, 'r13-winrt-plain')
      rmSync(renderDir128, { recursive: true, force: true })
      rmSync(renderDirPlain, { recursive: true, force: true })
      const w1 = runRenderScript(aesStage, renderDir128, ['1'])
      const w1txt = `${(w1 && w1.stdout) || ''}${(w1 && w1.stderr) || ''}${(w1 && w1.error) || ''}`.replace(/\s+/g, ' ').trim()
      const pngs128 = existsSync(renderDir128) ? readdirSync(renderDir128).filter(f => f.endsWith('.png')) : []
      // 明文对照（探针本身有效）
      const w2 = runRenderScript(plainStage, renderDirPlain, ['1'])
      const pngsPlain = existsSync(renderDirPlain) ? readdirSync(renderDirPlain).filter(f => f.endsWith('.png')) : []
      rmSync(probeStage, { recursive: true, force: true })
      const winrtAes128 = pngs128.length >= 1
      ok('R13-2 判别实验（结论已实测记录）：WinRT 打开同骨架 AES-128 夹具 → '
        + (winrtAes128 ? '能开 ⇒ 骨架正确；AES-256 打不开指向 WinRT 不支持 R5/R6 或 V5 字典细节' : '打不开 ⇒ 骨架本身有问题，需逐条复核加密字典'),
        pngsPlain.length >= 1, // 判据：探针本身有效（明文对照能渲染）；AES-128 的结果作为**结论**记录
        `plainPng=${pngsPlain.length} aes128Png=${pngs128.length} ${w1txt.slice(0, 90)}`)
      ok('R13-2 判别实验：探针本身有效（同一脚本打开明文 PDF 能渲染出 PNG）',
        pngsPlain.length >= 1, `plainPng=${pngsPlain.length}${w2 && w2.error ? ` ${w2.error.message}` : ''}`)
    }

    // —— 真实样本槽位（需求 2d）：主人手上有真实 R5/R6 样本时一条命令复验 ——
    const realSample = process.env.DSH_OFFICE_TEST_AES256_PDF
    if (!realSample) {
      ok('R13-2 真实样本槽位：DSH_OFFICE_TEST_AES256_PDF 未设 → 跳过（设上即可复验，自证风险的门在这里）', true,
        '未设 DSH_OFFICE_TEST_AES256_PDF；设置后本用例会用 office_read 打开该文件并断言能读出正文')
    } else {
      const rs = await call('office_read', { path: realSample, as: 'markdown' })
      ok('R13-2 真实样本：第三方生成的 AES-256 PDF 能被本插件读出正文（自证风险就此关闭）',
        String(rs.content || '').trim().length > 0, `${String(rs.content || '').length} 字符`)
    }
  }

  // ---------------- R13-7：pptx 被 PowerPoint 16.0 拒开（需求 7） ----------------
  {
    const deck = join(OUT, 'r13-deck.pptx')
    await call('office_create', {
      path: deck,
      slides: {
        slides: [
          { layout: 'title', title: '演示标题', subtitle: '副标题' },
          { title: '要点页', bullets: [{ text: '一级要点', level: 0 }, { text: '二级要点', level: 1 }, { text: '三级要点', level: 2 }], notes: '这里是备注。' },
          { title: '第三页', bullets: [{ text: '收尾', level: 0 }] },
        ],
      },
    })
    const deckText = await textOf(deck)
    const zd = openZip(readFileSync(deck))
    const notesRels = zd.getText('ppt/notesMasters/_rels/notesMaster1.xml.rels')
    ok('R13-7 pptx：备注母版有**独立**主题部件（notesMaster→theme2，slideMaster→theme1）',
      /Target="\.\.\/theme\/theme2\.xml"/.test(notesRels)
      && zd.names.includes('ppt/theme/theme2.xml')
      && /PartName="\/ppt\/theme\/theme2\.xml"/.test(zd.getText('[Content_Types].xml'))
      && /Target="\.\.\/theme\/theme1\.xml"/.test(zd.getText('ppt/slideMasters/_rels/slideMaster1.xml.rels')),
      notesRels.replace(/\s+/g, ' ').slice(0, 120))
    // 负向控制：把 notesMaster 的主题改回与 slideMaster 共用 theme1 → PowerPoint 必须仍拒开
    const badRels = notesRels.replace('theme2.xml', 'theme1.xml')
    const badEntries = zd.names.filter(n => !n.endsWith('/')).map(n => ({
      name: n,
      data: n === 'ppt/notesMasters/_rels/notesMaster1.xml.rels' ? badRels : zd.get(n),
    }))
    const badDeck = join(OUT, 'r13-deck-badrel.pptx')
    writeFileSync(badDeck, Buffer.from(makeZip(badEntries)))
    const badDeckText = await textOf(badDeck)
    ok('R13-7 pptx：修复不动读回内容（theme2 版与 theme1 版读回逐字一致）',
      badDeckText === deckText, `${deckText.length} vs ${badDeckText.length} 字符`)

    const PPT_PS = `
$ErrorActionPreference = 'Continue'
$src = $env:DSH_OFFICE_PPTX
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300
try { $p = New-Object -ComObject PowerPoint.Application } catch { 'PPT_NO_POWERPOINT ' + $_.Exception.Message; exit 0 }
try {
  $d = $p.Presentations.Open($src, -1, 0, 0)
  $n = $d.Slides.Count
  $d.Close()
  "PPT_OPEN_OK slides=$n"
} catch {
  'PPT_OPEN_FAIL ' + $_.Exception.Message
} finally {
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
`
    const pptProbe = src => spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PPT_PS],
      { encoding: 'utf8', windowsHide: true, timeout: 240000, env: { ...process.env, DSH_OFFICE_PPTX: src } })
    const pptText = r => `${(r && r.stdout) || ''}${(r && r.stderr) || ''}`.trim()
    if (process.platform !== 'win32') {
      ok('R13-7 pptx：PowerPoint 打开产出的 pptx（非 Windows → 跳过）', true, `platform=${process.platform}`)
    } else {
      const gt = pptText(pptProbe(deck))
      const bt = pptText(pptProbe(badDeck))
      if (!gt || /PPT_NO_POWERPOINT|80040154|Retrieving the COM class factory|not registered/i.test(gt)) {
        ok('R13-7 pptx：PowerPoint 打开产出的 pptx（PowerPoint 不可用 → 跳过）', true, gt.replace(/\s+/g, ' ').slice(0, 120))
      } else {
        ok('R13-7 pptx：PowerPoint 16.0 能打开产出的 .pptx（不再 0x80070570）',
          /PPT_OPEN_OK/.test(gt) && /slides=3/.test(gt), gt.replace(/\s+/g, ' ').slice(0, 140))
        ok('R13-7 pptx：负向控制 —— notesMaster 改回共用 theme1 必须仍被 PowerPoint 拒开',
          /PPT_OPEN_FAIL/.test(bt), bt.replace(/\s+/g, ' ').slice(0, 140))
      }
    }
  }

  // ---------------- error handling ----------------
  let threw = false
  try { await call('office_read', { path: join(OUT, 'not-exist.docx') }) } catch { threw = true }
  ok('缺失文件报错', threw)

  threw = false
  try { await call('office_create', { path: join(OUT, 'x.unknownext') }) } catch { threw = true }
  ok('不支持格式报错', threw)

  threw = false
  try { await call('office_read', { path: docx, bogus: 1 }) } catch { threw = true }
  ok('未知参数报错', threw)

  threw = false
  try { await call('office_edit', { path: docx, operations: [{ op: 'nope' }] }) } catch { threw = true }
  ok('非法操作报错', threw)

  // ===========================================================================
  // R18：大 stored 图片条目端到端 / 大输入质量门基准 / zip maxOutputLength 早停 RSS
  // 前缀：R18-大图： / R18-基准： / R18-zip：
  // ===========================================================================
  {
    // ---- E2 真实大 stored 图片条目：2000×2000 PNG（噪声内容 → PNG 本身不可再压）----
    {
      // 噪声源必须**不可压缩**：LCG 的低位有强周期性，deflate 能把 12 MB 压到 0.1 MB，
      // 那样就不是"大 stored 条目"了。改成从一个 64 KiB 随机池循环取样 —— 池距 65536 大于
      // deflate 的 32 KiB 窗口，LZ77 找不到可复用匹配，才能真的产出 ~12 MB 的 PNG。
      const noisePool = randomBytes(65536)
      let ni = 0
      const rand = () => noisePool[ni++ & 0xffff]
      const bigPngPath = join(OUT, 'r18-big-2000x2000.png')
      writeFileSync(bigPngPath, makeTestPng(2000, 2000, 2, (r, x) => {
        r[x * 3] = rand(); r[x * 3 + 1] = rand(); r[x * 3 + 2] = rand()
      }))
      const pngBytes = statSync(bigPngPath).size
      const docxBig = join(OUT, 'r18-big-image.docx')
      await call('office_create', { path: docxBig, markdown: '# 大图宿主\n\n锚点段落：把图插在这行之后。\n' })
      const t0 = Date.now()
      const ins = await call('office_edit', { path: docxBig, operations: [{ op: 'insert_image', path: bigPngPath, after: '锚点段落' }] })
      const insMs = Date.now() - t0
      const buf = readFileSync(docxBig)
      const zb = openZip(buf)
      const mediaName = zb.names.find(n => n.startsWith('word/media/'))
      const mediaSize = mediaName ? Buffer.from(zb.get(mediaName)).length : 0
      ok('R18-大图：2000×2000 噪声 PNG 作为**大媒体条目**写进 docx（内容逐字节一致）',
        /image1\.png$/.test(String(mediaName)) && mediaSize === pngBytes && mediaSize > 10 * 1024 * 1024,
        `${mediaName} ${(mediaSize / 1048576).toFixed(2)} MiB = 源 PNG ${(pngBytes / 1048576).toFixed(2)} MiB（${insMs}ms）`)
      // 该条目必须是 **store（method 0）**：PNG 已是压缩数据，二次压缩只会更大 → 这正是"合法大 stored 条目"场景
      const eocdAt = buf.length - 22
      let q = buf.readUInt32LE(eocdAt + 16)
      let method = -1
      while (q + 4 <= buf.length && buf.readUInt32LE(q) === 0x02014b50) {
        const nlen = buf.readUInt16LE(q + 28)
        const nm = buf.subarray(q + 46, q + 46 + nlen).toString('utf8')
        if (nm === mediaName) { method = buf.readUInt16LE(buf.readUInt32LE(q + 42) + 8); break }
        q += 46 + nlen + buf.readUInt16LE(q + 30) + buf.readUInt16LE(q + 32)
      }
      ok('R18-大图：该条目以 store（method 0）落地 —— 合法大 stored 条目不被限额误拒', method === 0, `method=${method}`)
      const bigRead = await call('office_read', { path: docxBig, as: 'markdown' })
      ok('R18-大图：含大媒体条目的 docx 仍能被 office_read 完整读回（守门断言）',
        String(bigRead.content).includes('锚点段落'), `${String(bigRead.content).length} 字符`)
      ok('R18-大图：大条目远低于单条目上限（阈值有依据，不是"刚好擦边"）',
        mediaSize < maxEntryBytes() / 10,
        `media=${(mediaSize / 1048576).toFixed(2)}MiB，单条目上限=${(maxEntryBytes() / 1048576).toFixed(0)}MiB`)
      ok('R18-大图：insert_image 的字节账包含媒体 + XML（不是只写了 XML）',
        typeof ins.bytes === 'number' && ins.bytes > mediaSize, `${ins.bytes} 字节`)
    }

    // ---- E3 大输入基准：textQuality 在 8 MiB 级输入上不栈溢出、耗时在实测量级内 ----
    // 详细数字（耗时 / RSS）落在 DEVELOPMENT.md 的 R18 交付报告里；这里只做守门。
    const benchCases = [
      ['8 MiB 连续字母', 'a'.repeat(8 * 1024 * 1024)],
      ['8 MiB 连续句点（点前导）', '.'.repeat(8 * 1024 * 1024)],
      ['8 MiB 正常中文', '中文测试内容。这是一段正常的中文正文。'.repeat(140000)],
      ['8 MiB 混合', 'The quick brown fox 中文混排段落。'.repeat(200000)],
    ]
    for (const [label, s] of benchCases) {
      const rss0 = process.memoryUsage().rss
      const t0 = Date.now()
      let q = null
      let err = ''
      try { q = textQuality(s) } catch (e) { err = String(e && e.message) }
      const ms = Date.now() - t0
      ok(`R18-基准：textQuality ${label} → 不栈溢出、耗时在实测量级内`, !err && ms < 8000 && !!q,
        err || `${ms}ms chars=${q.chars} rss+${((process.memoryUsage().rss - rss0) / 1048576).toFixed(1)}MiB`)
    }
    // E3' 口径等价守门：R18 把 textQuality 里两条**整串正则**改成手写扫描（见 index.js 的注释），
    // 判据本身一个数都不能变 —— 用小样本与旧"正则 + split"参考实现逐样本对照。
    {
      const refNoVowel = str => {
        let runs = 0, noVowel = 0
        for (const m of str.matchAll(/[A-Za-z]{6,}/g)) {
          runs += m[0].length
          for (const seg of m[0].split(/[aeiouAEIOU]+/)) if (seg.length >= 8) noVowel += seg.length
        }
        return { runs, noVowel }
      }
      const samples = [
        'abcdefgh', 'bcdfghjklmnp', 'a'.repeat(40), 'bcd'.repeat(20), 'Hello world xyz',
        '中文abcDEFghijklmnop', 'AEIOUaeiou', 'qwrtypsdfghjklzxcvbnm', 'abc def ghijklmnop qrstuvwx',
        '......', 'a......b', '第1章 标题..................3'.repeat(3),
      ]
      const bad = samples.filter(str => {
        const r = refNoVowel(str)
        return textQuality(str).noVowelRatio !== (r.runs ? r.noVowel / r.runs : 0)
      })
      ok('R18-基准：无元音连排口径与旧正则+split 参考实现逐样本一致（手写扫描不改判据）',
        bad.length === 0, bad.length ? bad.join(' | ') : `${samples.length} 个样本一致`)
      // 点前导口径等价：连续 ≥3 个句点整段不计入可见字符；<3 个照常计入
      const vis = str => textQuality(str).visible
      ok('R18-基准：点前导口径与旧版一致（"..."→0 可见字符、"."→1、2 个点→2）',
        vis('...') === 0 && vis('.') === 1 && vis('..') === 2 && vis('a...b') === 2 && vis('....') === 0,
        `${vis('...')}/${vis('.')}/${vis('..')}/${vis('a...b')}/${vis('....')}`)
    }

    // ---- E4 zip maxOutputLength 早停：RSS 增量必须远小于"真解压完"（阈值实测+自校准）----
    {
      const bombSize = 64 * 1024 * 1024
      const zipBuf = Buffer.from(makeZip([{ name: 'bomb.bin', data: Buffer.alloc(bombSize) }]))
      const compSize = zipBuf.readUInt32LE(18)
      const dataStart = 30 + zipBuf.readUInt16LE(26) + zipBuf.readUInt16LE(28)
      const rawDeflate = Buffer.from(zipBuf.subarray(dataStart, dataStart + compSize))
      // 把中央目录声明的 uncSize 改小（1 MiB）→ 读侧 cap 变小 → zlib 必须在流没解完时早停
      zipBuf.writeUInt32LE(1024 * 1024, zipBuf.readUInt32LE(zipBuf.length - 22 + 16) + 24)
      const rss0 = process.memoryUsage().rss
      let bombMsg = ''
      try { openZip(zipBuf).get('bomb.bin') } catch (e) { bombMsg = String(e.message) }
      const cappedDelta = process.memoryUsage().rss - rss0
      // 干净基线：把同一段 deflate 流交给**子进程**真解压完 64 MiB，量它自己的 RSS 增量
      const refLog = join(OUT, 'r18-rss-ref.txt')
      const rawFile = join(OUT, 'r18-raw-deflate.bin')
      rmSync(refLog, { force: true })
      writeFileSync(rawFile, rawDeflate)
      const childCode = "import { inflateRawSync } from 'node:zlib'\n"
        + "import { readFileSync, writeFileSync } from 'node:fs'\n"
        + `const z = readFileSync(${JSON.stringify(rawFile)})\n`
        + 'const r0 = process.memoryUsage().rss\n'
        + 'const out = inflateRawSync(z)\n'
        + `writeFileSync(${JSON.stringify(refLog)}, (process.memoryUsage().rss - r0) + ":" + out.length)\n`
      const child = runWithTimeout(180, childCode)
      const refTxt = existsSync(refLog) ? readFileSync(refLog, 'utf8').trim() : ''
      const refDelta = Number(refTxt.split(':')[0])
      const refLen = Number(refTxt.split(':')[1])
      ok('R18-zip：说谎的声明尺寸 → maxOutputLength 早停，RSS 增量远小于"真解压完"（自校准阈值）',
        /zip 炸弹防护/.test(bombMsg) && cappedDelta < 8 * 1024 * 1024
        && Number.isFinite(refDelta) && refDelta > cappedDelta * 4 && refLen === bombSize,
        `早停 Δ=${(cappedDelta / 1048576).toFixed(1)}MiB；子进程真解压 Δ=${Number.isFinite(refDelta) ? (refDelta / 1048576).toFixed(1) : refTxt}MiB`
        + `（${refLen} 字节，status=${child.status}）compress ${compSize} 字节`)
      rmSync(rawFile, { force: true })
      rmSync(refLog, { force: true })
    }

    // ---- R19-D3：8 MiB 单行输入端到端守门 ----
    // 旧版 `textQuality` 的整串正则在"单个巨大匹配"上抛 `RangeError`（宿主级异常、整条 office_read 作废）。
    // 这里从**工具边界**守门：md 与"超大 XML 部件"两条路都必须正常返回。
    {
      const bigMd = join(OUT, 'r19-8mib-single-line.md')
      writeFileSync(bigMd, 'a'.repeat(8 * 1024 * 1024), 'utf8')
      const t0 = Date.now()
      let rMd = null, eMd = ''
      try { rMd = await call('office_read', { path: bigMd, as: 'text' }) } catch (e) { eMd = String(e.message) }
      const mdMs = Date.now() - t0
      ok('R19-大输入：`office_read` 读 8 MiB 单行 .md → 不抛 RangeError（旧版整串正则会崩在宿主级）',
        !eMd && typeof rMd?.content === 'string' && !/Maximum call stack/.test(eMd) && mdMs < 30000,
        eMd ? eMd.slice(0, 140) : `${mdMs}ms，content ${String(rMd.content).length} 字符`)
      ok('R19-大输入：超长正文要么完整返回、要么按内联护栏截断并带说明（不静默丢）',
        typeof rMd?.content === 'string' && rMd.content.length > 0
        && (rMd.content.length <= 200000
          || /上限|截断|truncate/i.test(String(rMd.stats?.truncateNote || '') + String(rMd.content).slice(-400))),
        `content=${String(rMd?.content || '').length} 字符，truncateNote=${String(rMd?.stats?.truncateNote || '(none)').slice(0, 80)}`)

      // 超大 XML 部件：造一个 document.xml 里含 8 MiB 单行文本的真 .docx（走 XML 解析 + 质量画像）
      const docxBig = join(OUT, 'r19-8mib-single-line.docx')
      await call('office_create', { path: docxBig, markdown: '# 大输入\n\n正文。\n' })
      const z0 = openZip(readFileSync(docxBig))
      const bigXml = z0.getText('word/document.xml').replace('</w:body>',
        `<w:p><w:r><w:t>${'b'.repeat(8 * 1024 * 1024)}</w:t></w:r></w:p></w:body>`)
      const rebuilt = makeZip(z0.names.map(n => ({
        name: n, data: n === 'word/document.xml' ? Buffer.from(bigXml, 'utf8') : Buffer.from(z0.get(n)),
      })))
      writeFileSync(docxBig, rebuilt)
      const t1 = Date.now()
      let rDx = null, eDx = ''
      try { rDx = await call('office_read', { path: docxBig, as: 'text' }) } catch (e) { eDx = String(e.message) }
      const dxMs = Date.now() - t1
      ok('R19-大输入：`office_read` 读"8 MiB 单行 XML 部件"的 .docx → 不抛 RangeError、不超时',
        !eDx && typeof rDx?.content === 'string' && !/Maximum call stack/.test(eDx) && dxMs < 60000,
        eDx ? eDx.slice(0, 140) : `${dxMs}ms，content ${String(rDx.content).length} 字符`)

      // 解析层点位（R19 任务 D2 审计发现、本轮修掉的三处"复合体量词"RangeError）
      let mgErr = '', mgBlocks = 0
      try { mgBlocks = markdownToDocument('-'.repeat(2 * 1024 * 1024)).blocks.length } catch (e) { mgErr = String(e.message) }
      ok('R19-大输入：`markdownToDocument` 单行 2 MiB "-" → 不 RangeError，且真 hr 仍被识别',
        !mgErr && mgBlocks === 1 && markdownToDocument('---\n\n正文\n').blocks[0].type === 'hr',
        mgErr || `${mgBlocks} 块；"---"→${markdownToDocument('---\n\n正文\n').blocks[0].type}`)
      let xErr = ''
      let xAttrs = ''
      try {
        xAttrs = JSON.stringify(parseXML('<w:p a="1" b=\'2\' c=3/>').children[0].attrs)
        if (parseXML(`<x${' a=a'.repeat(2000000)}>`).children.length !== 1) xErr = '属性序列解析结果异常'
      } catch (e) { xErr = String(e.message) }
      ok('R19-大输入：`parseXML` 单标签内 2e6 个属性 → 不 RangeError（OOXML / ODF 部件入口）',
        !xErr, xErr.slice(0, 130) || 'ok')
      ok('R19-大输入：`parseXML` 正常属性仍解析正确（手写扫描与旧正则口径一致）',
        xAttrs === '{"a":"1","b":"2","c":"3"}', xAttrs)
      const odtBad = makeZip([
        { name: 'mimetype', data: 'application/vnd.oasis.opendocument.text', store: true },
        { name: 'content.xml', data: '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>'
          + `&#${'9'.repeat(400)};` + '</text:p></office:text></office:body></office:document-content>' },
      ])
      let odtBigErr = ''
      try { readOdt(odtBad) } catch (e) { odtBigErr = String(e.message) }
      ok('R19-大输入：含 400 位数字实体的 .odt → 不再 RangeError（`odf.js::decodeNumeric` 过去缺守卫，可达）',
        !odtBigErr, odtBigErr.slice(0, 130) || 'ok')
      process.env.DSH_OFFICE_PDF_SKIP_RENDER_CHECK = '1'
      let pgErr = ''
      try { await pdfOutputGate(join(OUT, 'r19-gate-hr.pdf'), Buffer.from('%PDF-1.4\n'), '-'.repeat(8 * 1024 * 1024), {}) } catch (e) { pgErr = String(e.message) }
      delete process.env.DSH_OFFICE_PDF_SKIP_RENDER_CHECK
      ok('R19-大输入：`pdfOutputGate` 整篇 8 MiB 单行 "-" → 不 RangeError（hr-only 仍算无可见文本）',
        !pgErr, pgErr.slice(0, 130) || 'ok')
    }
  }

  console.log(results.join('\n'))
  console.log(`\n${failures === 0 ? '✅ ALL PASS' : `❌ ${failures} FAILED`}  (${results.length} checks)  产物：${OUT}`)
  process.exitCode = failures === 0 ? 0 : 1
}

main().catch(e => {
  console.log(results.join('\n'))
  console.error('\nHARNESS ERROR:', e)
  process.exitCode = 2
})
