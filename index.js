// dsh-office — system-level Office all-in-one plugin for the DeepSeek Harness.
//
// Registers four host tools (office_read / office_create / office_edit /
// office_convert) into the shared tool registry, so every session of every
// profile can use them. Zero third-party dependencies: Node built-ins only,
// with the OOXML / ODF / PDF / OLE2 implementations shipped in this directory.
import { readFile, writeFile, mkdir, stat, rename, open, rm } from 'node:fs/promises'
import { existsSync, mkdirSync, readFileSync, writeFileSync, statSync, readdirSync, rmSync, openSync, readSync, closeSync, writeSync, renameSync, fsyncSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { tmpdir } from 'node:os'
import { resolve as resolvePath, dirname, extname, join, basename, relative as relativePath } from 'node:path'
import { openZip, makeZip } from './zip.js'
import {
  sniff, resolveKind, decodeTextBytes, parseDelimited, renderDelimited, tableMd,
  markdownToDocument, markdownToSlides, documentToMarkdown,
  normalizeDocument, normalizeWorkbook, normalizeSlides,
} from './model.js'
import { parseXML, serializeXML, findAll, encodeEntities } from './xml.js'
import { readDocx, writeDocx, replaceTextInDocument, appendBlocksToDocument, imageParagraphXml, insertImageParagraph, ensureContentTypeDefault, normalizeGrid, imageSizeFor } from './docx.js'
import { readXlsx, writeXlsx, colName } from './xlsx.js'
import { readPptx, writePptx, replaceTextInPptxPart } from './pptx.js'
import { serializeSlideXmlForEdit, slideRelsForEdit, notesXmlForEdit, notesRelsForEdit } from './pptx-edit.js'
import { readOdt, readOds, readOdp, writeOdt, writeOds, writeOdp } from './odf.js'
import { readPdfText, writePdf } from './pdf.js'
import { parseOle2, readLegacyDoc, readLegacyXls, readLegacyPpt, textToBlocks } from './legacy.js'
import { cropPngBand, pngInkCoverage } from './png.js'
// 第十二轮需求 1b/1c：零依赖图片嗅探/解码（docx 插图的尺寸、PDF `/XObject` 的采样数据）。
import { readImageBytes, sniffImage, imageExt, imageMime } from './image.js'
// 第十二轮需求 3：xlsx 公式重算（纯 TS 子集求值器；只有显式 `recalc=true` 才启用）。
import { recalcWorkbook } from './formula.js'
import { findEngine, ocrImages, gateResult, engineLabel, LOCAL_ENGINE, LOCAL_MAX_IMAGES } from './rapidocr.js'
// 需求 1：HTML 读取（零依赖 DOM 解析器，见 html.js）。带 ?v= 缓存爆破参数：
// 与 cordis.patch.yml 的热换机制同款——以后改 html.js，把这里的 v 加一即可随 index.js 一起热生效。
import { htmlToDocument, plainTextOf, documentToHtml } from './html.js?v=1'

const name = 'dsh-wen'
const inject = ['tools']

/** This plugin's own directory (holds pdf-render.ps1). */
const PLUGIN_DIR = dirname(fileURLToPath(import.meta.url))
const OCR_DEFAULT_PAGES = 3
const OCR_MAX_PAGES = 20

const MAX_FILE_BYTES = 200 * 1024 * 1024
const READ_CAP = 200000

// ---------------------------------------------------------------------------
// tool definition helper — produces the registry shape that
// @deepseek-ai/dsh-tools' defineTool() produces, without importing it.
// ---------------------------------------------------------------------------

const TYPE_NAMES = { string: '字符串', integer: '整数', number: '数字', boolean: '布尔值', array: '数组', object: '对象' }
const typeName = t => TYPE_NAMES[t] || String(t)

/** 单类型判定（validator 与 schema 共用同一套口径）。 */
function typeOk(t, v) {
  switch (t) {
    case 'string': return typeof v === 'string'
    case 'integer': return Number.isInteger(v)
    case 'number': return typeof v === 'number'
    case 'boolean': return typeof v === 'boolean'
    case 'array': return Array.isArray(v)
    case 'object': return typeof v === 'object' && v !== null && !Array.isArray(v)
    default: return true
  }
}

function jsonSchemaOf(spec) {
  // 任务四：`spec.types` 声明"同一参数接受多种类型"（grid 的数字与字符串）→ `oneOf`。
  // 依据 host 的 JSON Schema 子集：`oneOf` 是一等关键字（≥2 支、不与 type 同级），
  // 而 `type: [...]` 数组是**被明确拒绝**的写法（"type arrays are not supported"）。
  const types = Array.isArray(spec.types) && spec.types.length ? spec.types : null
  const base = types && types.length > 1 ? { oneOf: types.map(t => ({ type: t })) } : { type: spec.type }
  // enum 不能与 oneOf 同级（host 子集硬规则），多类型参数靠 validator 的 enum 段兜底
  if (spec.enum && !base.oneOf) base.enum = spec.enum
  if (spec.description) base.description = spec.description
  if (spec.type === 'array') base.items = spec.items ? jsonSchemaOf(spec.items) : { type: 'string' }
  if (spec.type === 'object') {
    base.additionalProperties = spec.additionalProperties !== false
    if (spec.properties) {
      base.properties = Object.fromEntries(Object.entries(spec.properties).map(([k, v]) => [k, jsonSchemaOf(v)]))
      const req = Object.entries(spec.properties).filter(([, v]) => v.required).map(([k]) => k)
      if (req.length) base.required = req   // required 必须都在 properties 里（host 子集硬规则）
    }
  }
  return base
}

/**
 * 元素级校验（任务四）：`operations[i]` 过去只声明 `{type:'object'}`，元素内部零约束，
 * 模型给出的 `{}`、`[{op:'nope'}]`、缺 `newName` 的 `rename_sheet` 全都被静默放到执行期，
 * 甚至落到实现里变成 `Sheet1 → undefined`。这里给**元素**一份带下标的校验：
 *   ① 必须是对象；② `op` 必填且取自实现支持的 16 个取值；
 *   ③ 容器无关的必需字段（同一 op 跨容器都必需的那些）缺失即报错，错误里带 `operations[N]`。
 * 容器相关必需性仍由执行期精确报错（各格式的兜底文案本来就有"可用: …"清单）。
 */
function itemViolations(path, el, rules) {
  const errs = []
  const ops = rules?.ops || []
  // R18 任务 D2：把三种"非对象元素"分开口径，文案里直接说清**实际收到了什么**。
  // 旧文案对 `null` / 数字 / 数组一视同仁（"应为对象"），排查时看不出是哪种。
  if (el === null || el === undefined) {
    return [`${path} 不能为空（${el === null ? 'null' : 'undefined'}）；每一项都应是对象，形如 {op:"replace_text", find:"旧", replace:"新"}`]
  }
  if (typeof el !== 'object' || Array.isArray(el)) {
    const got = Array.isArray(el) ? '数组' : (TYPE_NAMES[typeof el] || typeof el)
    return [`${path} 应为对象（形如 {op:"replace_text", find:"旧", replace:"新"}），实际收到${got}`]
  }
  const op = el.op
  if (op === undefined || op === null || op === '') {
    return [`${path} 缺少必填字段 "op"（可用: ${ops.join(' / ')}）`]
  }
  if (typeof op !== 'string' || !ops.includes(op)) {
    return [`${path}.op 未识别：${JSON.stringify(op)}（可用: ${ops.join(' / ')}）`]
  }
  for (const need of rules.required?.[op] || []) {
    const hit = need.any
      ? need.any.some(f => el[f] !== undefined && el[f] !== '')
      : el[need.field] !== undefined && el[need.field] !== ''
    if (hit) continue
    errs.push(need.any
      ? `${path} 的 "${op}" 需要 ${need.any.map(f => `"${f}"`).join(' 或 ')} 之一`
      : `${path} 的 "${op}" 缺少必填字段 "${need.field}"`)
  }
  for (const [field, t] of Object.entries(rules.fields?.[op] || {})) {
    if (el[field] !== undefined && el[field] !== null && !typeOk(t, el[field])) {
      errs.push(`${path}.${field} 应为${typeName(t)}`)
    }
  }
  return errs
}

function defineToolLite(options) {
  const parameters = {
    type: 'object',
    properties: Object.fromEntries(Object.entries(options.parameters).map(([k, v]) => [k, jsonSchemaOf(v)])),
    required: Object.entries(options.parameters).filter(([, v]) => v.required).map(([k]) => k),
  }
  const violations = args => {
    const errs = []
    const input = args && typeof args === 'object' ? args : {}
    for (const [k, spec] of Object.entries(options.parameters)) {
      const val = input[k]
      if (val === undefined || val === null) {
        if (spec.required) errs.push(`缺少必填参数 "${k}"`)
        continue
      }
      const types = Array.isArray(spec.types) && spec.types.length ? spec.types : spec.type ? [spec.type] : null
      if (types && !types.some(t => typeOk(t, val))) {
        // 单类型文案逐字不变（既有调用方/测试依赖它）；多类型才用"…或…"
        errs.push(types.length === 1
          ? `参数 "${k}" 应为${typeName(types[0])}`
          : `参数 "${k}" 应为${types.map(typeName).join(' 或 ')}`)
      } else if (Array.isArray(val) && spec.items?.itemRules) {
        // R18 任务 D2：用**下标循环**而不是 forEach —— 稀疏数组（`new Array(2)`）里的空洞
        // forEach 会直接跳过，于是 `operations: [ <hole> ]` 一路滑到执行期变成裸 TypeError。
        for (let i = 0; i < val.length; i++) errs.push(...itemViolations(`${k}[${i}]`, val[i], spec.items.itemRules))
      }
      if (spec.enum && !spec.enum.includes(val)) errs.push(`参数 "${k}" 只能是 ${spec.enum.map(e => JSON.stringify(e)).join(' / ')}`)
    }
    for (const k of Object.keys(input)) {
      if (!(k in options.parameters)) errs.push(`未知参数 "${k}"（可用: ${Object.keys(options.parameters).join(', ')}）`)
    }
    return errs
  }
  return {
    name: options.name,
    description: options.description,
    parameters,
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        // R18 任务 D1：只声明"**若**返回里有 content，它必须是字符串"，**不加 required**。
        // 目的：将来某条返回把 content 写成对象/数组时，host 侧立刻以 ToolOutputError 拦下，
        // 而不是等到调用方 `JSON.parse` 才炸（as="json" 的契约就是"content 是 JSON 文本"）。
        // 动手前已核实：本文件所有 `content` 赋值都是字符串（capText/capWithOffset 的纯前缀、
        // PDF 的 `JSON.stringify(bare)`、office_read 正文、批量盘点管道表、convert 报告）。
        properties: { content: { type: 'string' } },
      },
      // render 投影与 execute 走同一条出站边界：host 侧同样会校验 render 的无损性
      render: options.render && ((a, v) => finalizeToolValue(options.render(a, v), a || {})),
    },
    async execute(args, exec) {
      const errs = violations(args)
      if (errs.length) throw new Error(errs.join('；'))
      try {
        // 唯一出站终点：消毒 + 压平 + 记账。返回值绝不会再被 host 判 "not lossless JSON"。
        return finalizeToolValue(await options.execute(args, exec), args || {})
      } catch (e) {
        throw sanitizeThrown(e)
      }
    },
  }
}

// ---------------------------------------------------------------------------
// path / io helpers
// ---------------------------------------------------------------------------

function sessionCwd(exec) {
  try {
    return exec?.agent?.session?.header?.cwd || process.cwd()
  } catch {
    return process.cwd()
  }
}

function hostPath(input, exec) {
  const raw = String(input ?? '').trim().replace(/^["']|["']$/g, '')
  if (!raw) throw new Error('path 不能为空')
  let p = raw
  if (p === '~' || p.startsWith('~/') || p.startsWith('~\\')) {
    p = resolvePath(process.env.USERPROFILE || process.env.HOME || '.', p.slice(1).replace(/^[/\\]/, ''))
  }
  const isAbsolute = /^[a-zA-Z]:[\\/]/.test(p) || p.startsWith('\\\\') || p.startsWith('/')
  return resolvePath(isAbsolute ? p : resolvePath(sessionCwd(exec), p))
}

function extOf(file) {
  return extname(file).toLowerCase().replace(/^\./, '')
}

function throwIfAborted(exec) {
  if (exec?.signal?.aborted) throw new Error('操作已取消')
}

async function readBuffer(file) {
  const info = await stat(file).catch(() => undefined)
  if (!info) throw new Error(`文件不存在: ${file}`)
  if (!info.isFile()) throw new Error(`不是普通文件: ${file}`)
  if (info.size > MAX_FILE_BYTES) throw new Error(`文件过大 (${Math.round(info.size / 1048576)}MB)，上限 ${MAX_FILE_BYTES / 1048576}MB`)
  return readFile(file)
}

async function saveBuffer(file, bytes) {
  // 任务二：一切落盘都走"同目录唯一临时件 → 完整写入 → 成功后替换目标"，
  // 不再直接 writeFile(file)（open 即截断，中断就留半截文件；编辑场景还会毁原件）。
  return writeFileAtomic(file, bytes)
}

// ---------------------------------------------------------------------------
// 原子落盘（任务二）：目标同目录唯一临时文件 → 完整写入 → 成功后替换目标
// ---------------------------------------------------------------------------
//
// 病灶：`saveBuffer` 旧实现是 `writeFile(file, bytes)`（flag 'w'，**open 那刻就截断**）——
// 写盘中断/失败会留下半截文件；`office_edit` 的落盘点 100% 覆盖原件，等于"改坏原件"。
// 而 `pdfOutputGate` 早就用"临时件 + 校验 + 改名"做对了，本段把它抽成公共两段式，
// 让 create / edit / convert / sidecar 全部复用同一套语义（**不叠加两层**）。
//
// 三条 Windows 实测约束（见 r16 审计）：
//  ① `rename` 覆盖**已存在**的目标是成功的 —— 所以正常路径直接 rename，无需先删目标；
//  ② 目标上有任何句柄（编辑器/杀毒/同步盘）时 rename 会 EPERM/EBUSY —— 有限退避重试，
//     仍失败就**保留原目标**并抛错；**绝不 unlink 目标**（那会绕过只读保护、并制造
//     "目标短暂消失"的窗口，而且真正被锁住时 unlink 同样 EBUSY，救不了它想救的场景）；
//  ③ 跨卷 rename 会 EXDEV —— 临时件必须放在**目标同目录**，绝不能用 tmpdir()。
//
// 临时名带 `O_EXCL`（`open(..., 'wx')`）而不只靠随机后缀：并发下唯一性由内核保证。
// 残留的临时件以 `.part` 结尾：`extOf()` 取到的是 "part"，不在 `SCAN_EXTS` 里，
// 因此不会被 `office_read paths=[目录]` 当成用户文档盘点（旧 `.dsh-pdf-gate-*.pdf` 会）。
const ATOMIC_TMP_PREFIX = '.dsh-tmp-'
const ATOMIC_ERROR_TAIL = '；下一步=确认目标目录可写、且文件未被其他程序占用（编辑器/杀毒/同步盘）后重试'
let atomicSeq = 0

/** 测试用故障注入点（只在本模块的 helper 内判断，不污染其它探针）。 */
function atomicFault(stage) {
  return process.env.DSH_OFFICE_ATOMIC_FAULT === stage
}

function atomicFaultError(stage) {
  return new Error(`注入故障（DSH_OFFICE_ATOMIC_FAULT=${stage}）`)
}

function atomicFsyncEnabled() {
  return process.env.DSH_OFFICE_ATOMIC_FSYNC !== '0'
}

/** 目标同目录唯一临时路径：保留原扩展名（`pdf-render.ps1` 按内容加载，零成本保险）。 */
function tempPathBeside(file) {
  const ext = extname(file)
  const stem = basename(file, ext).replace(/[^\w.-]/g, '_').slice(0, 24) || 'out'
  const salt = randomBytes(8).toString('hex')
  return join(dirname(file), `${ATOMIC_TMP_PREFIX}${stem}-${process.pid.toString(36)}-${(atomicSeq++).toString(36)}-${salt}${ext}.part`)
}

/** 写盘失败的四要素错误：说清目标、根因、原文件状态、下一步。 */
function atomicWriteError(file, e, tmp) {
  const root = String(e?.message || e)
  const left = tmp && existsSync(tmp) ? `；临时文件未能清理：${tmp}` : ''
  return new Error(`【写盘失败｜四要素】目标=${file}；根因=${root}；`
    + '已保留原文件（未被截断、未被删除），临时文件已清理' + left + ATOMIC_ERROR_TAIL)
}

function sleepSync(ms) {
  try { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms) } catch { /* 退化：直接重试 */ }
}

const PUBLISH_RETRY_DELAYS = [60, 120, 180]

/** 只发布：把临时件改名到目标（Windows 上 rename 能覆盖已存在目标）。 */
async function publishTemp(tmp, file) {
  if (atomicFault('publish')) throw atomicFaultError('publish')
  for (let i = 0; ; i++) {
    try {
      await rename(tmp, file)
      return file
    } catch (e) {
      const transient = /^(EPERM|EBUSY|EACCES)$/.test(String(e?.code || ''))
      if (!transient || i >= PUBLISH_RETRY_DELAYS.length) throw e
      await new Promise(r => setTimeout(r, PUBLISH_RETRY_DELAYS[i]))
    }
  }
}

function publishTempSync(tmp, file) {
  if (atomicFault('publish')) throw atomicFaultError('publish')
  for (let i = 0; ; i++) {
    try {
      renameSync(tmp, file)
      return file
    } catch (e) {
      const transient = /^(EPERM|EBUSY|EACCES)$/.test(String(e?.code || ''))
      if (!transient || i >= PUBLISH_RETRY_DELAYS.length) throw e
      sleepSync(PUBLISH_RETRY_DELAYS[i])
    }
  }
}

/** 只写临时件（`O_EXCL` 保证并发唯一），返回临时路径；失败自行清理后抛出。 */
async function writeTempBeside(file, bytes) {
  await mkdir(dirname(file), { recursive: true })
  const tmp = tempPathBeside(file)
  let fh
  try {
    if (atomicFault('temp-open')) throw atomicFaultError('temp-open')
    fh = await open(tmp, 'wx')
    if (atomicFault('temp-write')) throw atomicFaultError('temp-write')
    await fh.writeFile(bytes)
    // fsync 是 best-effort：吃掉"断电后 0 字节/半截"这一整类用户可见风险（实测 8MiB 仅 +4.3ms）。
    // 目录 fsync 在 Windows 上不可用（EPERM），故不承诺"改名本身跨掉电持久"。
    if (atomicFsyncEnabled()) { try { await fh.sync() } catch { /* 网络盘/OneDrive 上失败无所谓 */ } }
    await fh.close()
    fh = undefined
    return tmp
  } catch (e) {
    if (fh) { try { await fh.close() } catch { /* ignore */ } }
    try { await rm(tmp, { force: true }) } catch { /* ignore */ }
    throw e
  }
}

function writeTempBesideSync(file, data) {
  mkdirSync(dirname(file), { recursive: true })
  const tmp = tempPathBeside(file)
  let fd
  try {
    if (atomicFault('temp-open')) throw atomicFaultError('temp-open')
    fd = openSync(tmp, 'wx')
    if (atomicFault('temp-write')) throw atomicFaultError('temp-write')
    writeSync(fd, data)
    if (atomicFsyncEnabled()) { try { fsyncSync(fd) } catch { /* ignore */ } }
    closeSync(fd)
    fd = undefined
    return tmp
  } catch (e) {
    if (fd !== undefined) { try { closeSync(fd) } catch { /* ignore */ } }
    try { rmSync(tmp, { force: true }) } catch { /* ignore */ }
    throw e
  }
}

/**
 * 原子写文件：完整写入临时件 → 成功后替换目标。失败时清理临时件并**保留旧目标**。
 * @returns {Promise<string>} 目标路径
 */
export async function writeFileAtomic(file, bytes) {
  const tmp = await writeTempBeside(file, bytes).catch(e => {
    throw atomicWriteError(file, e, undefined)
  })
  try {
    await publishTemp(tmp, file)
    return file
  } catch (e) {
    try { await rm(tmp, { force: true }) } catch { /* ignore */ }
    throw atomicWriteError(file, e, tmp)
  }
}

/** 同步版：sidecar 写入要保持"同步返回、失败 undefined"的既有契约。 */
export function writeFileAtomicSync(file, data) {
  const tmp = writeTempBesideSync(file, data)
  try {
    publishTempSync(tmp, file)
    return file
  } catch (e) {
    try { rmSync(tmp, { force: true }) } catch { /* ignore */ }
    throw atomicWriteError(file, e, tmp)
  }
}

// ---------------------------------------------------------------------------
// PDF 渲染校验的**暂存目录**（R18 任务 A）
// ---------------------------------------------------------------------------
//
// 病灶（R18 实测，探针 `work/r18-probe/probe-winrt.mjs`）：WinRT 的
// `StorageFile.GetFileFromPathAsync` 在本机**只允许 `%TEMP%` 下的文件** ——
// 工作区 / 用户目录下的 PDF 一律「拒绝访问，该项目没有位于应用程序可以访问的位置」。
// 旧实现把渲染校验用的副本写在**目标同目录**（`writeTempBeside(file, …)`），于是
// **在任何非 TEMP 目录 create / convert 一个 PDF 都会 100% 失败**（报"第 1 页渲染失败…逃生通道"）；
// R16 只是把测试产物目录挪到 TEMP 绕开了它，代码没修。
//
// 现在：渲染副本一律落 `tmpdir()`（可用 `DSH_OFFICE_PDF_GATE_DIR` 覆盖），渲染完立即删除；
// 校验**通过后**才走公共两段式（目标同目录唯一临时件 + rename）发布 —— 「未经校验不发布」的语义
// 一字不变，也不再为了渲染往目标目录写任何东西。
export function pdfGateStagingRoot() {
  const raw = process.env.DSH_OFFICE_PDF_GATE_DIR
  const set = raw && String(raw).trim()
  return resolvePath(set ? String(raw).trim() : join(tmpdir(), 'dsh-office-pdfgate'))
}

// ---------------------------------------------------------------------------
// 阶段二：PDF 产出质量门（G2：宁可不产出，也不写"渲染为空白"的垃圾）
// ---------------------------------------------------------------------------

/** 逃生通道（实测可用）：md → odt → Word COM → ExportAsFixedFormat(17)。 */
const PDF_ESCAPE = '逃生通道（已实测可用）：md → office_create(.odt) → PowerShell Word COM（New-Object -ComObject Word.Application; $word.Documents.Open($odt,$false,$true); $doc.ExportAsFixedFormat($out,17)）导出 PDF（系统会对字体子集自动嵌入）'

/**
 * `office_create` / `office_convert` 对 **PDF 目标**的出站质量门（阶段二）。
 *
 * 两级判据（OCR 不参与，纯字节/像素启发式，不消耗视觉额度）：
 *  1. 字节级：正文含 CJK 却没有任何 `/FontFile*` → 未内嵌字体
 *     （阶段一验收后收紧为 hard-fail；`DSH_OFFICE_PDF_EMBED_CJK=0` 的显式回退同样拦）。
 *  2. 渲染级（主判据）：bytes 先写 **`tmpdir()` 下的暂存件**（R18 任务 A：可被
 *     `DSH_OFFICE_PDF_GATE_DIR` 覆盖）→ `pdf-render.ps1` 渲染第 1 页 →
 *     "有可见文本却整页黑像素 < 40 个（≈ 完全没渲染）"，或"文本 ≥ 24 字却 PNG < 30KB
 *     且墨迹 < 0.1%"判渲染为空白。hr-only / 仅分页标记的文档不算"有可见文本"，不误杀。
 *     暂存件**必须在 `%TEMP%` 下**：WinRT 读不到工作区 / 用户目录的 PDF（见 pdfGateStagingRoot）。
 * 通过 → 复用公共两段式（目标同目录唯一临时件 + rename）发布到目标；
 * 失败 → 删除暂存件与临时件并抛四要素错误（绝不落盘、绝不在校验前碰目标）。
 * `DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` 或非 Windows → `renderCheck="skipped"`（仍写字节级结论）。
 *
 * @param {string} file 目标路径
 * @param {Buffer} bytes 待落盘字节
 * @param {object|string} model 内容模型（判"有无可见文本/含不含 CJK"）
 * @param {{notes?:string[]}} info writePdf 回填的嵌入状态
 * @returns {Promise<{embedded:boolean, renderCheck:'pass'|'skipped', firstPageBytes:number}>}
 */
/**
 * R19 任务 D：把"去掉 markdown 分隔线行"从**复合体量词**正则
 * `/^[ \t]*([-*_])(?:[ \t]*\1){2,}[ \t]*$/gm` 改成逐行手写判据 ——
 * 旧正则在"整篇 8 MiB 单行 `-`"（合法 markdown）上抛
 * `RangeError: Maximum call stack size exceeded`（探针 `work/r19-probe/probe-biginput-sites.mjs`）。
 * 语义与旧版一致：命中行**只删内容、保留换行**（等价于把该行替换成空串）。
 */
function stripMarkdownHrLines(s) {
  const out = []
  let i = 0
  while (i <= s.length) {
    const nl = s.indexOf('\n', i)
    const line = nl < 0 ? s.slice(i) : s.slice(i, nl)
    out.push(isHrLineTab(line) ? '' : line)
    if (nl < 0) break
    i = nl + 1
  }
  return out.join('\n')
}

/** 该行是否只由空格 / `\t` 与同一个 `-` / `*` / `_` 组成、且该字符 ≥3 个。 */
function isHrLineTab(line) {
  let c = ''
  let count = 0
  for (let k = 0; k < line.length; k++) {
    const ch = line[k]
    if (ch === ' ' || ch === '\t') continue
    if (c === '') {
      if (ch !== '-' && ch !== '*' && ch !== '_') return false
      c = ch
    } else if (ch !== c) return false
    count++
  }
  return count >= 3
}

export async function pdfOutputGate(file, bytes, model, info = {}) {
  // "可见文本"：Markdown 分隔线（模型里 hr 渲染为 `---`）与分页注释不算文本，
  // 否则 hr-only / pagebreak-only 的文档会走进"有文本"分支被墨迹判据误杀。
  const text = stripMarkdownHrLines(String(typeof model === 'string' ? model : modelToText(model)))
    .replace(/<!--\s*pagebreak\s*-->/gi, '')
    .replace(/\s/g, '')
  const hasText = !!text
  const hasCjk = /[\u3400-\u9fff\u3000-\u303f\uff00-\uffef\u2e80-\u2eff\uf900-\ufaff]/.test(text)
  // 页数：辅助"空白"判据必须与"只渲染了首页"**同源**——整篇文本量对多页文档是错口径。
  // 第八轮实测（本批新增用例暴露的真 bug）：26 页文档首页只有一行英文、整篇 407 字
  // → 被误判"文本量可观却渲染近乎全白"而拒绝落盘（首页 PNG 10116 字节、墨迹 0.039%、544 px）。
  // 主判据（整页几乎无墨，< 40 px）对所有文档仍然生效，单页文档行为逐字不变。
  const pageCount = typeof model === 'string'
    ? 1
    : (Array.isArray(model?.blocks) ? model.blocks.filter(b => b?.type === 'pagebreak').length + 1 : 1)
  const latin = Buffer.isBuffer(bytes) ? bytes.toString('latin1') : Buffer.from(bytes ?? '').toString('latin1')
  const embedded = /\/FontFile[23]?\s/.test(latin)
  // 插图账（第十二轮需求 1c）：嵌入了几张、**哪几张没嵌进去、为什么** —— 一律进 stats 与 notice。
  const imageStats = {
    images: Array.isArray(info.images) ? info.images.length : 0,
    ...(info.imagesSkipped?.length ? { imagesSkipped: info.imagesSkipped.map(x => ({ name: x.name, reason: x.reason })) } : {}),
  }
  const imageNotice = info.imagesSkipped?.length
    ? `${info.imagesSkipped.length} 张图片未能嵌入：${info.imagesSkipped.map(x => `${x.name}（${x.reason}）`).join('；')}`
    : ''
  const withImages = base => ({ ...base, ...imageStats, ...(imageNotice ? { notice: imageNotice } : {}) })
  const head = '【PDF 产出拒绝｜四要素】页码=第 1 页（产出抽检）；格式=pdf；根因='
  const tail = '；已删除临时半成品、未产出目标文件'
    + `；下一步=${PDF_ESCAPE}`
    + '（或设 DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1 跳过渲染级抽检）'

  // —— 1. 字节级：有 CJK 却没嵌字体 ——
  if (hasCjk && !embedded) {
    throw new Error(`${head}正文含 CJK 但产出未内嵌任何字体（/FontFile* 计数=0`
      + `${info.notes?.length ? `；字体链：${info.notes.join('；')}` : ''}）`
      + '——不嵌字体的中文 PDF 在 WinRT/Edge/Chrome 渲染空白或问号、提取丢字符，故拒绝落盘'
      + `${tail}；先确认 C:\\Windows\\Fonts\\simsun.ttc 存在（DSH_OFFICE_PDF_EMBED_CJK=0 的显式回退同样会被本门拦截）`)
  }

  // —— 2. 渲染级（可跳过）——
  if (process.env.DSH_OFFICE_PDF_SKIP_RENDER_CHECK === '1' || process.platform !== 'win32') {
    await saveBuffer(file, bytes)
    return withImages({ embedded, renderCheck: 'skipped', firstPageBytes: 0 })
  }

  // R18 任务 A：渲染副本落 tmpdir()（可被 DSH_OFFICE_PDF_GATE_DIR 覆盖），**不再写目标同目录**
  // —— WinRT 只能读 %TEMP% 下的 PDF。校验通过后才用公共两段式发布（同目录唯一名 + O_EXCL +
  // `.part` 后缀 + publishTemp），与 create / edit / convert / sidecar 同一套语义，**不叠加两层**；
  // 也绝不为了渲染往目标目录写任何副本。
  const gateTag = `${process.pid}-${Date.now()}-${Math.floor(Math.random() * 1e6)}`
  const pngDir = join(tmpdir(), `dsh-office-pdfgate-${gateTag}`)
  const stageDir = join(pdfGateStagingRoot(), gateTag)
  const stage = join(stageDir, 'check.pdf')
  let tmp
  try {
    mkdirSync(stageDir, { recursive: true })
    writeFileSync(stage, bytes)
    // 渲染重试一次：多会话并发时 WinRT 的异步 Wait 偶发聚合异常（瞬态竞态，
    // 实测单跑稳定成功、4 路并发会炸）——重试仍失败才是真渲染失败。
    let res = runRenderScript(stage, pngDir, ['1'])
    const png = join(pngDir, 'page-1.png')
    if (!existsSync(png)) res = runRenderScript(stage, pngDir, ['1'])
    if (!existsSync(png)) throw new Error(`${head}第 1 页渲染失败（${renderFailureDetail(res)}${renderLocationHint(stage)}）${tail}`)
    const firstPageBytes = statSync(png).size
    if (hasText) {
      let m = null
      try { m = pngInkCoverage(readFileSync(png)) } catch { m = null }
      const chars = text.length
      let inkTxt, blank, why
      if (m) {
        inkTxt = `${(m.ink * 100).toFixed(3)}%（${m.dark} px）`
        // 主判据：整页几乎无墨（< 40 个黑像素）＝ 文本没被渲染出来；
        // 辅助判据：文本量可观（≥24 字）却 PNG 很小且墨迹 <0.1% ＝ 大半页空白。
        // 短文本（标题/单行）即便墨迹占比低也不判空白，避免误杀。
        why = m.dark < 40 ? '整页几乎无墨' : '文本量可观却渲染近乎全白'
        blank = m.dark < 40
          || (pageCount === 1 && chars >= 24 && firstPageBytes < 30 * 1024 && m.ink < 0.001)
      } else {
        inkTxt = '未测'
        why = '文本量可观却渲染出极小页图'
        blank = chars >= 24 && firstPageBytes < 30 * 1024
      }
      if (blank) {
        throw new Error(`${head}渲染为空白（首页 PNG ${firstPageBytes} 字节、黑像素覆盖率 ${inkTxt}、`
          + `可见文本 ${chars} 字；判据：${why}）` + tail)
      }
      tmp = await writeTempBeside(file, bytes)
      await publishTemp(tmp, file)
      return withImages({ embedded, renderCheck: 'pass', firstPageBytes })
    }
    // 无可见文本（hr/空文档）：只要渲染管线能出图即过，不按"空白"误杀
    tmp = await writeTempBeside(file, bytes)
    await publishTemp(tmp, file)
    return withImages({ embedded, renderCheck: 'pass', firstPageBytes })
  } catch (e) {
    if (tmp) rmSync(tmp, { force: true })
    throw e
  } finally {
    rmSync(pngDir, { recursive: true, force: true })
    rmSync(stageDir, { recursive: true, force: true })
    if (tmp) rmSync(tmp, { force: true })
  }
}

function byteLength(bytes) {
  if (bytes == null) return 0
  if (bytes instanceof Uint8Array) return bytes.length
  return Buffer.byteLength(String(bytes))
}

function capText(s, limit) {
  if (s.length <= limit) return s
  return `${s.slice(0, limit)}\n\n…[已达 ${limit} 字符上限，已截断。可用 offset/limit 继续读取，或先转存文件后分段读取]`
}

/**
 * 截断 + 给出**精确**的下一个 offset。
 *
 * 截断协议（收敛后）：`content` 只放**纯前缀**，一个字节的说明文字都不掺。
 * 说明改走返回值里的 `note`（调用方挂到 `notice` / `stats.truncateNote`）。
 * 于是硬不变式对得上账：
 *
 *     offset + content.length === nextOffset      （截断时）
 *
 * 旧版把"…已截断，请用 offset:N 续读"这类后缀拼进 content，消费方只能写
 * `content.slice(0, nextOffset - offset)` 才能取到真正文 —— 拼接协议因此变得
 * 依赖调用方的自觉，写错就"又重又漏"。现在 content 就是正文，直接拼接即可。
 */
function capWithOffset(body, limit, offset) {
  const truncated = body.length > limit
  const content = truncated ? body.slice(0, limit) : body
  const nextOffset = truncated ? offset + content.length : undefined
  const note = truncated
    ? `正文超过单次读取上限 ${limit} 字符，已截断 ${body.length - content.length} 字符；`
      + `用 offset:${nextOffset} 继续读取（本段正文共 ${body.length} 字符）`
    : ''
  return { content, truncated, nextOffset, note }
}

// ---------------------------------------------------------------------------
// 返回边界：码点消毒 + lossless-JSON 收口
// ---------------------------------------------------------------------------
//
// 症状：任何**含正文**的 office_read 都返回
//
//   tool "office_read" returned invalid output: value is not lossless JSON
//
// 模型的正文全部丢失，只看到这一句；`as="meta"` 因为不含正文反而正常。
// 这条判定由 host 侧在工具体返回值离开 execute() 之前做（DSH 走
// @deepseek-ai/dsh-util-values 的 `snapshotJsonValue`；WorkBuddy 走 api.mjs
// 的 JSON.stringify，不会报错但同样在这条边界上丢码点）。它拒绝的东西按
// 出现频率排序是：
//
//   1. `undefined` 作为对象属性 —— **实测元凶**。`stats.ocrCovered: undefined`
//      这类"有则给、无则 undefined"的写法，JSON.stringify 会安静丢掉，
//      但判定器把它当成不可序列化值 → 整条结果作废（PDF 正文读取全军覆没）
//   2. NaN / ±Infinity / -0
//   3. 稀疏数组、数组上的非索引自有键（含 symbol 键）
//   4. 类实例 / Date / Map / Set / Buffer / 函数 —— 非"纯对象"
//
// 该判定**不检查字符串内容**（字符串只判 typeof），所以游离代理本身不会
// 触发这条报错；但游离代理与 C0/C1 控制符会在跨进程传输、写入磁盘、被
// 日志/Read 工具消费时被静默换成 `?` 或 U+FFFD，属于同一条边界上的隐患，
// 因此一并消毒。repro.mjs 是这条边界的常驻哨兵。
//
// 因此所有返回管线终点统一走 `finalizeToolValue()`：字符串逐码点消毒，
// 结构压成纯 JSON 树，全程记账（stats.sanitized / stats.sanitizeNotes + 正文脚注）。

const REPLACEMENT_CHAR = '\ufffd'
/** 唯一允许原样保留的控制符：制表、换行、回车。 */
const KEEP_CONTROL = new Set([0x09, 0x0a, 0x0d])

/**
 * C0/C1 控制符（\t \n \r 除外）。
 * U+2028/U+2029 是合法的行/段分隔符（Word、PowerPoint 都在用），不是控制符，
 * 保留不动 —— 消毒的原则是"宁留噪不删正文"。
 */
function isControlCodePoint(cp) {
  return (cp <= 0x1f && !KEEP_CONTROL.has(cp)) || (cp >= 0x7f && cp <= 0x9f)
}

/**
 * 单个字符串的边界消毒：游离代理 / U+FFFE / U+FFFF / C0+C1 控制符 → U+FFFD，
 * 然后 NFC 规范化。合法配对代理（补充平面字符，如 emoji、CJK 扩展 B）不动。
 * 只替换、不删除、不断句。
 *
 * @param {unknown} input 任意值（非字符串会被 String() 化）
 * @returns {{value: string, fixedCount: number, fixed: Array<{at: number, codepoint: number, kind: string}>, nfcChanged: boolean}}
 */
export function sanitizeTextForReturn(input) {
  const s = typeof input === 'string' ? input : String(input ?? '')
  let out = ''
  const fixed = []
  for (let i = 0; i < s.length; i++) {
    const cp = s.codePointAt(i)
    if (cp > 0xffff) { out += s.slice(i, i + 2); i++; continue }   // 合法配对 → 不动
    let kind = ''
    if (cp >= 0xd800 && cp <= 0xdfff) kind = 'lone-surrogate'
    else if (cp === 0xfffe || cp === 0xffff) kind = 'noncharacter'
    else if (isControlCodePoint(cp)) kind = 'control'
    if (kind) { fixed.push({ at: i, codepoint: cp, kind }); out += REPLACEMENT_CHAR; continue }
    out += s[i]
  }
  const nfc = out.normalize('NFC')
  return { value: nfc, fixedCount: fixed.length, fixed, nfcChanged: nfc !== out }
}

/**
 * 本插件自带的"无损 JSON"判定，语义与 host 侧对齐。
 * 用途有二：① 返回前自检，坏结构在插件内部就被修掉而不是被整条拒收；
 * ② 测试/repro.mjs 可以直接断言"这个返回值 host 一定收"。
 * @returns {string|null} 第一处违规的定位描述；合法时返回 null
 */
export function losslessJsonProblem(value, path = '$', seen = new Set()) {
  if (value === null) return null
  const t = typeof value
  if (t === 'string' || t === 'boolean') return null
  if (t === 'number') return Number.isFinite(value) && !Object.is(value, -0) ? null : `${path} 不是无损 JSON 数字（${String(value)}）`
  if (t === 'undefined') return `${path} 是 undefined（host 会整条拒收）`
  if (t === 'bigint') return `${path} 是 bigint`
  if (t === 'function' || t === 'symbol') return `${path} 是 ${t}`
  if (t !== 'object') return `${path} 类型 ${t} 不可序列化`
  if (seen.has(value)) return `${path} 存在循环引用`
  seen.add(value)
  try {
    if (Array.isArray(value)) {
      const keys = Reflect.ownKeys(value)
      if (keys.length !== value.length + 1) return `${path} 是稀疏数组或带额外自有键`
      for (let i = 0; i < value.length; i++) {
        if (!Object.prototype.hasOwnProperty.call(value, i)) return `${path}[${i}] 缺失（稀疏数组）`
        const bad = losslessJsonProblem(value[i], `${path}[${i}]`, seen)
        if (bad) return bad
      }
      return null
    }
    const proto = Object.getPrototypeOf(value)
    if (proto !== null && proto !== Object.prototype) return `${path} 不是纯对象（${proto?.constructor?.name || '未知'}）`
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== 'string' || !Object.prototype.propertyIsEnumerable.call(value, key)) return `${path} 含不可枚举/symbol 自有键`
      const bad = losslessJsonProblem(value[key], `${path}.${key}`, seen)
      if (bad) return bad
    }
    return null
  } finally {
    seen.delete(value)
  }
}

const OMIT = Symbol('dsh-office.omit')

/**
 * 深度收口：把任意返回值压成纯 JSON 树，同时消毒所有字符串。
 * 每一步降级都写进 ctx.notes —— 绝不静默丢弃。
 */
function losslessValueOf(input, ctx, depth = 0) {
  if (input === undefined) { ctx.droppedCount++; return OMIT }
  if (input === null) return null
  const t = typeof input
  if (t === 'string') {
    const r = sanitizeTextForReturn(input)
    if (r.fixedCount) {
      ctx.sanitized += r.fixedCount
      for (const f of r.fixed) ctx.fixed.push(f)
    }
    if (r.nfcChanged) ctx.notes.add('文本已做 NFC 规范化')
    return r.value
  }
  if (t === 'boolean') return input
  if (t === 'number') {
    if (Number.isFinite(input) && !Object.is(input, -0)) return input
    ctx.notes.add(`非有限数值 ${String(input)} 已替换为 null`)
    return null
  }
  if (t === 'bigint') { ctx.notes.add('bigint 已转为字符串'); return String(input) }
  if (t === 'function' || t === 'symbol') { ctx.notes.add(`${t} 值已丢弃`); ctx.droppedCount++; return OMIT }
  // object
  if (ctx.seen.has(input)) { ctx.notes.add('检测到循环引用，已置为 null'); return null }
  if (depth > 64) { ctx.notes.add('嵌套深度超过 64 层，已截断为 null'); return null }
  ctx.seen.add(input)
  try {
    if (Array.isArray(input)) {
      return input.map((item, i) => {
        if (!Object.prototype.hasOwnProperty.call(input, i)) { ctx.notes.add(`数组第 ${i} 项缺失（稀疏数组）已补 null`); return null }
        const r = losslessValueOf(item, ctx, depth + 1)
        return r === OMIT ? null : r          // 数组不能有洞：undefined → null
      })
    }
    if (input instanceof Date) { ctx.notes.add('Date 已转为 ISO 字符串'); return Number.isFinite(input.getTime()) ? input.toISOString() : null }
    if (ArrayBuffer.isView(input) || input instanceof ArrayBuffer) {
      const bytes = input.byteLength ?? input.length ?? 0
      ctx.notes.add(`二进制值（${bytes} 字节）已转为描述对象`)
      return { type: 'binary', bytes }
    }
    if (input instanceof Map) { ctx.notes.add(`Map（${input.size} 项）已转为对象`); return losslessValueOf(Object.fromEntries(input), ctx, depth + 1) }
    if (input instanceof Set) { ctx.notes.add(`Set（${input.size} 项）已转为数组`); return losslessValueOf([...input], ctx, depth + 1) }
    const proto = Object.getPrototypeOf(input)
    if (proto !== null && proto !== Object.prototype) ctx.notes.add(`非纯对象（${proto?.constructor?.name || '未知'}）已展开为普通对象`)
    const out = {}
    for (const key of Reflect.ownKeys(input)) {
      if (typeof key !== 'string') { ctx.notes.add('symbol 键已丢弃'); ctx.droppedCount++; continue }
      if (!Object.prototype.propertyIsEnumerable.call(input, key)) { ctx.notes.add(`不可枚举自有键 "${key}" 已丢弃`); ctx.droppedCount++; continue }
      const r = losslessValueOf(input[key], ctx, depth + 1)
      if (r === OMIT) { ctx.droppedCount++; continue }      // ← undefined 属性在这里被剔除
      out[key] = r
    }
    return out
  } finally {
    ctx.seen.delete(input)
  }
}

/** 消毒账本。 */
function newSanitizeContext() {
  return { seen: new Set(), notes: new Set(), fixed: [], sanitized: 0, droppedCount: 0 }
}

/** 正文尾部脚注：只在真的替换过码点时出现，绝不静默。 */
function sanitizeFootnote(n) {
  return `${n} 个非法码点已替换为 U+FFFD`
}

/** as=json / as=meta 的 content 本身是 JSON 文本，脚注只能进 notice 字段。 */
function injectJsonNotice(content, text) {
  try {
    const payload = JSON.parse(content)
    if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
      payload.notice = payload.notice ? `${payload.notice}；${text}` : text
      return JSON.stringify(payload, null, 2)
    }
  } catch { /* 不是 JSON 就退回调用方的默认处理 */ }
  return null
}

/** 内联体积护栏默认阈值；`DSH_OFFICE_MAX_INLINE_CHARS` 可调，0=关闭。 */
function maxInlineChars() {
  const raw = process.env.DSH_OFFICE_MAX_INLINE_CHARS
  if (raw !== undefined && String(raw).trim() !== '') {
    const n = Number(raw)
    if (n === 0) return Infinity
    if (Number.isFinite(n) && n >= 1000) return Math.floor(n)
  }
  return 120000
}

/** 在尽量靠后的段落边界切一刀，避免把一行/一句话劈成两半。 */
function cutAtBoundary(s, limit) {
  const head = s.slice(0, limit)
  const idx = Math.max(head.lastIndexOf('\n\n'), head.lastIndexOf('\n'))
  return idx > limit * 0.6 ? head.slice(0, idx) : head
}

/**
 * 工具返回值出站前的唯一终点：消毒 → 压平 → 记账 → 附注。
 * 覆盖 office_read / office_create / office_edit / office_convert 全部返回路径，
 * 以及 render 投影（host 侧同样会校验 render 的无损性）。
 *
 * @param {unknown} value execute()/render() 的原始返回值
 * @param {object} [args] 调用参数（用于判断 content 是否为 JSON 正文、是否显式给了 limit）
 */
export function finalizeToolValue(value, args = {}) {
  const ctx = newSanitizeContext()
  const safe = losslessValueOf(value, ctx)
  const out = safe === OMIT ? null : safe

  // ---- 内联体积护栏：超阈值不再交给 host 截断，而是自己给续读协议 ----
  // 与 capWithOffset 同一套协议：content 只放纯前缀，说明走 notice + stats.truncateNote，
  // 硬不变式 `offset + content.length === nextOffset` 由这里保证。
  let guardrail
  if (out && typeof out === 'object' && !Array.isArray(out) && typeof out.content === 'string'
    && args.as !== 'json' && args.as !== 'meta' && args.limit === undefined) {
    const cap = maxInlineChars()
    if (Number.isFinite(cap) && out.content.length > cap) {
      const kept = cutAtBoundary(out.content, cap)
      const droppedChars = out.content.length - kept.length
      const base = Math.max(0, Number(args.offset) || 0)
      const nextOffset = base + kept.length
      const note = `正文超过内联上限 ${cap} 字符，已截断 ${droppedChars} 字符；`
        + `用 offset:${nextOffset} 继续读取，或用 pages/sheet 分批取（DSH_OFFICE_MAX_INLINE_CHARS 可调，0=关闭护栏）`
      out.content = kept
      out.truncated = true
      out.nextOffset = nextOffset
      out.notice = typeof out.notice === 'string' && out.notice ? `${out.notice}；${note}` : note
      guardrail = { cap, kept: kept.length, dropped: droppedChars, note }
    }
  }

  if (!out || typeof out !== 'object' || Array.isArray(out)) return out

  // ---- 记账 ----
  const noteList = [...ctx.notes]
  if (guardrail) noteList.push(`内联护栏：正文保留 ${guardrail.kept} 字符，其余 ${guardrail.dropped} 字符可续读`)
  const stats = out.stats && typeof out.stats === 'object' && !Array.isArray(out.stats) ? out.stats : null
  if (stats) {
    stats.sanitized = ctx.sanitized
    if (guardrail) stats.truncateNote = guardrail.note
    if (noteList.length) stats.sanitizeNotes = noteList
  } else if (noteList.length) {
    out.sanitizeNotes = noteList
  }

  // ---- 正文脚注（JSON 正文改走 notice，避免破坏可解析性）----
  if (ctx.sanitized > 0) {
    const note = sanitizeFootnote(ctx.sanitized)
    if (typeof out.content === 'string') {
      if (args.as === 'json' || args.as === 'meta') {
        const patched = injectJsonNotice(out.content, note)
        if (patched !== null) out.content = patched
      } else {
        out.content += `\n\n> ${note}`
      }
    }
  }
  return out
}

/** 抛出去的错误信息同样要消毒：错误也会跨同一条边界。 */
function sanitizeThrown(error) {
  const raw = error instanceof Error ? error : new Error(String(error))
  const { value } = sanitizeTextForReturn(safeMessage(raw))
  if (raw.message !== value) {
    try { raw.message = value } catch { return new Error(value) }
  }
  return raw
}

function safeMessage(error) {
  try {
    if (error instanceof Error) return error.message
    if (error && typeof error === 'object' && 'message' in error && typeof error.message === 'string') return error.message
    return String(error)
  } catch {
    return '<无法打印的错误>'
  }
}

// ---------------------------------------------------------------------------
// 正文质量门 + 读取降级链（fail-safe）
// ---------------------------------------------------------------------------
//
// 降级链四级：文本层提取 → 质量门 → 自动本地 OCR → sidecar 兜底。
// 质量门只认"绝对不可信"的信号（替换字符 / 私用区 / 控制符），**不拿语言当判据**
// —— 英文文档的 CJK 覆盖率天然为 0，用它当闸门会把正常英文 PDF 判成乱码。
// CJK 覆盖率只作**相对**信号：同一文档内某页显著低于全书水平。

const PRIVATE_USE_START = 0xe000
const PRIVATE_USE_END = 0xf8ff

/** CJK 表意文字 + 全角形式（覆盖率统计用；不含标点，避免标点页虚高）。 */
function isCjkCodePoint(cp) {
  return (cp >= 0x4e00 && cp <= 0x9fff) || (cp >= 0x3400 && cp <= 0x4dbf)
    || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff01 && cp <= 0xff60)
}

/**
 * 单段文本的质量画像。
 * @param {string} text
 * @param {{cjkRatio: number}|null} baseline 同文档基线（整本一起算），用于相对判据
 * @param {{structural?: boolean}} [opts] structural=true 时启用**字符级/版面级启发检测**
 *   （需求 4：重复行率、单字行率、连续非词典字符率——"字符可提取但语义破碎"的页面）。
 *   只在 **PDF 逐页**质量门上开启（textLayerProfile / readPdf 的 suspect 循环）；
 *   整本文本（finishRead / 非 PDF meta / convert 探针）保持旧口径，避免整本级
 *   "重复行/短行"（合法的表格、清单）把误报抬上去——那是页面级启发，不是文档级判据。
 * @returns {{chars:number, visible:number, cjk:number, cjkRatio:number, asciiLetterRatio:number,
 *            replacement:number, replacementRatio:number, privateUse:number, privateUseRatio:number,
 *            control:number, controlRatio:number, oddCharRatio:number, noVowelRatio:number,
 *            repeatedLineRatio:number, shortLineRatio:number,
 *            garbled:boolean, reasons:string[]}}
 */
export function textQuality(text, baseline = null, opts = {}) {
  const structural = !!opts.structural
  const s = typeof text === 'string' ? text : String(text ?? '')
  // 装饰性"点前导"（目录页 `......`，连续 ≥3 个 ASCII 句点）**不计入可见字符**。
  // 它是排版填充而非文字内容；计入会让目录页撞上"相对 CJK 覆盖率"判据的**假阳性**
  // —— 病灶样本①第 4 页（目录）：12.2% vs 全书 73.8% → 整本被 convert 拒绝、目录页白跑一次 OCR。
  // 这**不是放宽闸门**：真缺 ToUnicode 的文档表现为 **PUA / 替换字符成片**，不会长成"一片点号"。
  // 下面的私用区 / 替换字符 / 控制符判据一字未动，仍各自独立命中（见第三轮"点前导"用例）。
  //
  // R18 任务 E：这里原先是 `for (const m of s.matchAll(/\.{3,}/g))` 先把下标塞进一个 Set。
  // 两个问题：① "一整篇 8 MiB 的点号"（合法输入）会让 `RegExpStringIterator` 抛
  // `RangeError: Maximum call stack size exceeded`；② 巨大输入下 Set 会涨到 N 条。
  // 现在改成在主扫描循环里**就地识别点号连排**（零分配、零正则）。
  let cjk = 0, replacement = 0, privateUse = 0, control = 0, asciiLetter = 0, visible = 0
  // 需求 4"连续非词典字符率"的字符集成分：非词典区 = 正常办公文本不会成片出现的码位。
  //   oddLatin —— Latin-1 扩展字母带（0xC0–0xFF，×÷ 除外）：UTF-8 被按 Latin-1 解的
  //               mojibake（"ä¸­æ–‡"）几乎全落在这里；也含 CID 错映射。
  //   oddScript —— 希腊/西里尔（0x370–0x4FF）：中文/英文文档里的俄文整段多半是错映射。
  //   oddBox —— 制表/方块/几何符号（0x2500–0x25FF）：CID 字体错映射的常见产物。
  // 希腊/西里尔**只参与相对判据**（俄文原文档 baseline 本身就高，绝不误伤）；
  // oddLatin/oddBox 另有绝对阈值（合法西语文本的带重音字母占比远够不到阈值）。
  let oddLatin = 0, oddScript = 0, oddBox = 0
  for (let i = 0; i < s.length; i++) {
    const cp = s.codePointAt(i)
    if (cp > 0xffff) i++
    if (cp === 0x2e) {
      // 点前导：本段连续 ≥3 个 ASCII 句点时整段跳过（逐字符语义与旧版 Set 完全一致）
      let j = i
      while (j < s.length && s.charCodeAt(j) === 0x2e) j++
      if (j - i >= 3) { i = j - 1; continue }
    }
    if (cp === 0xfffd) { replacement++; visible++; continue }
    if (cp >= PRIVATE_USE_START && cp <= PRIVATE_USE_END) { privateUse++; visible++; continue }
    if (isControlCodePoint(cp)) { control++; visible++; continue }
    if (cp === 0xfffe || cp === 0xffff || (cp >= 0xd800 && cp <= 0xdfff)) { control++; visible++; continue }
    if (cp === 0x20 || cp === 0x09 || cp === 0x0a || cp === 0x0d) continue
    visible++
    if (isCjkCodePoint(cp)) cjk++
    else if ((cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a)) asciiLetter++
    else if (cp >= 0xc0 && cp <= 0xff && cp !== 0xd7 && cp !== 0xf7) oddLatin++
    else if (cp >= 0x370 && cp <= 0x4ff) oddScript++
    else if (cp >= 0x2500 && cp <= 0x25ff) oddBox++
  }
  const ratio = n => (visible ? n / visible : 0)
  // "连续非词典字符率"的词法成分：无元音字母串（≥8 个字母连排且不含 aeiou）。
  // 真词（含拼音）元音密度高，几乎不会出现 ≥8 的无元音连排；错映射/混淆字母常整串无元音。
  let noVowelChars = 0, letterRunChars = 0
  // 任务三（R16）把 `s.match(/[A-Za-z]{6,}/g)` 换成 `matchAll` **惰性迭代**，理由是旧写法在
  // "单个 8 MiB 连续字母串"上抛 `RangeError: Maximum call stack size exceeded`。
  // R18 任务 E 实测复核：**换成 matchAll 并没有修好** —— 单个巨大匹配时
  // `RegExpStringIterator.next` 自己就炸（8 MiB 连续 `a`、以及 8 MiB 连续 `.` 都能复现），
  // 而那是完全合法的输入（大 XML 里一段无分隔内容 / 一整篇点号）。
  // 现在彻底不用正则、也不构造中间字符串：一遍扫描数出"≥6 个 ASCII 字母的连排"总长，
  // 以及其中"无元音段（≥8 个非 aeiou 字母连排）"的总长 —— 与旧口径逐字等价，O(n) 且零分配。
  const isAsciiLetterCp = cp => (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a)
  const isVowelCp = cp => cp === 0x61 || cp === 0x65 || cp === 0x69 || cp === 0x6f || cp === 0x75
    || cp === 0x41 || cp === 0x45 || cp === 0x49 || cp === 0x4f || cp === 0x55
  for (let i = 0; i < s.length;) {
    if (!isAsciiLetterCp(s.charCodeAt(i))) { i++; continue }
    let j = i
    while (j < s.length && isAsciiLetterCp(s.charCodeAt(j))) j++
    const len = j - i
    if (len >= 6) {
      letterRunChars += len
      let seg = 0
      for (let k = i; k < j; k++) {
        if (isVowelCp(s.charCodeAt(k))) { if (seg >= 8) noVowelChars += seg; seg = 0 } else seg++
      }
      if (seg >= 8) noVowelChars += seg
    }
    i = j
  }
  const q = {
    chars: s.length,
    visible,
    cjk,
    cjkRatio: ratio(cjk),
    asciiLetterRatio: ratio(asciiLetter),
    replacement,
    replacementRatio: ratio(replacement),
    privateUse,
    privateUseRatio: ratio(privateUse),
    control,
    controlRatio: ratio(control),
    oddCharRatio: ratio(oddLatin + oddScript + oddBox),
    noVowelRatio: letterRunChars ? noVowelChars / letterRunChars : 0,
    repeatedLineRatio: 0,
    shortLineRatio: 0,
    garbled: false,
    reasons: [],
  }
  if (visible >= 40) {
    if (q.replacementRatio >= 0.02) q.reasons.push(`替换字符占 ${(q.replacementRatio * 100).toFixed(1)}%`)
    if (q.privateUseRatio >= 0.02) q.reasons.push(`私用区码点占 ${(q.privateUseRatio * 100).toFixed(1)}%（CID 字体缺 ToUnicode 的典型产物）`)
    if (q.controlRatio >= 0.05) q.reasons.push(`控制字符占 ${(q.controlRatio * 100).toFixed(1)}%`)
    // 相对判据：中文书里某页 CJK 覆盖率极低。必须**额外**排除"正常英文/代码页"——
    // 英文页 ASCII 字母占比高（~0.8），CID 乱码被映射成 Latin-1/符号后 ASCII 字母占比低。
    // 少了这个护栏，中文书里的英文段落会被误判成乱码。
    if (baseline && baseline.cjkRatio > 0.3 && visible >= 200
      && q.cjkRatio < baseline.cjkRatio / 3 && q.asciiLetterRatio < 0.5) {
      q.reasons.push(`CJK 覆盖率 ${(q.cjkRatio * 100).toFixed(1)}% 远低于全书 ${(baseline.cjkRatio * 100).toFixed(1)}%（且非英文/代码页）`)
    }
    // —— 需求 4：字符级启发（绝对阈值的两条）——
    if (visible && oddLatin / visible >= 0.5) {
      q.reasons.push(`Latin-1 扩展字符占 ${((oddLatin / visible) * 100).toFixed(1)}%（疑似 UTF-8/字节编码错位或 CID 错映射）`)
    }
    if (visible && oddBox / visible >= 0.3) {
      q.reasons.push(`制表/几何符号占 ${((oddBox / visible) * 100).toFixed(1)}%（疑似 CID 错映射）`)
    }
    // —— 需求 4：连续非词典字符率 ——
    if (asciiLetter >= 60 && q.noVowelRatio >= 0.25) {
      q.reasons.push(`无元音字母串占 ${((q.noVowelRatio) * 100).toFixed(0)}%（连续非词典字符，疑似字符错映射）`)
    }
    // —— 需求 4：非词典区的**相对**判据（希腊/西里尔只在这里出现，俄文原文档不误伤）——
    const baseOdd = baseline?.oddCharRatio ?? 0
    if (baseline && baseOdd < 0.03 && visible >= 200 && q.oddCharRatio >= 0.12) {
      q.reasons.push(`非词典字符占 ${(q.oddCharRatio * 100).toFixed(1)}% 远高于全书 ${(baseOdd * 100).toFixed(1)}%（疑似编码错位）`)
    }
    q.garbled = q.reasons.length > 0
  }
  // —— 需求 4：版面级启发（只看行结构，自带的行数门槛就是闸门）——
  if (structural) {
    const lines = s.split(/\r?\n/).map(l => l.replace(/^[\s.·…‥_−-]+|[\s.·…‥_−-]+$/g, '').trim())
    // 重复行率：只统计"像正文"的行（≥4 字符且含字母/数字/CJK），避免页码/装饰线凑数
    const meaningful = lines.filter(l => l.length >= 4 && /[0-9A-Za-z\u00c0-\u024f\u3400-\u9fff\uf900-\ufaff]/.test(l))
    if (meaningful.length >= 10) {
      const counts = new Map()
      for (const l of meaningful) counts.set(l, (counts.get(l) || 0) + 1)
      let dup = 0
      for (const n of counts.values()) if (n > 1) dup += n
      q.repeatedLineRatio = dup / meaningful.length
      if (q.repeatedLineRatio >= 0.5) {
        q.reasons.push(`重复行占 ${Math.round(q.repeatedLineRatio * 100)}%（${dup}/${meaningful.length} 行逐字重复，疑似解析破碎）`)
        q.garbled = true
      }
    }
    // 单字/双字行率：逐字断行是"字符可提取但语义破碎"的典型版面（页码/纯数字行不算）
    const nonEmpty = lines.filter(l => l.length > 0)
    if (nonEmpty.length >= 12) {
      const short = nonEmpty.filter(l => l.length <= 2 && !/^[0-9]+$/.test(l)).length
      q.shortLineRatio = short / nonEmpty.length
      if (q.shortLineRatio >= 0.75) {
        q.reasons.push(`单字/双字行占 ${Math.round(q.shortLineRatio * 100)}%（${short}/${nonEmpty.length} 行，疑似逐字断行）`)
        q.garbled = true
      }
    }
  }
  return q
}

/**
 * 缓存落盘目录（opt-in）：设了 `DSH_OFFICE_CACHE_DIR` 就把 sidecar（`<name>.ocr.md` /
 * `<file>.read.md`）统一落到该目录 —— 只读场景（附件目录、网络盘）也能缓存，用户目录
 * 不被污染。未设置 → 返回空 dir，调用方走原来的"源文件同目录优先、不可写回退 %TEMP%"。
 * 目录不可写 → 回退默认位置，并把**回退原因**交回调用方进 stats/notice（绝不静默）。
 */
let cacheDirMemo = null

function cacheDirState() {
  const raw = process.env.DSH_OFFICE_CACHE_DIR
  if (raw === undefined || String(raw).trim() === '') return { want: '', dir: '', note: '' }
  const want = resolvePath(String(raw).trim())
  if (cacheDirMemo && cacheDirMemo.want === want) return cacheDirMemo
  let dir = ''
  let note = ''
  try {
    mkdirSync(want, { recursive: true })
    const probe = join(want, `.dsh-office-write-${process.pid}-${Date.now()}`)
    writeFileSync(probe, 'ok')
    rmSync(probe, { force: true })
    dir = want
  } catch (e) {
    note = `DSH_OFFICE_CACHE_DIR 不可用（${e?.code || e?.message || String(e)}），已回退到默认缓存位置`
  }
  cacheDirMemo = { want, dir, note }
  return cacheDirMemo
}

// ---------------------------------------------------------------------------
// 缓存身份（任务一）：sidecar 必须能证明"这份识别成果属于这个源的**这份内容**"
// ---------------------------------------------------------------------------
//
// 病灶（两条，都是静默的）：
//  ① 旧 sidecar 只校验 `parser:` 版本，**不绑定源内容**。同一路径的 PDF 被换成另一份
//     （另存为、重新扫描、覆盖下载）之后，只要 parser 没变，旧 OCR 文本就照旧被贴回来。
//  ② 集中缓存目录与临时回退目录按 `basename` 命名 —— `D:\a\第1章.pdf` 与 `D:\b\第1章.pdf`
//     会共用同一个 `第1章.ocr.md`：轻则 A 的识别成果被 B 当自己的命中，重则互相覆盖。
//     `.read.md`（兜底转存的整篇正文）与 `renderDirFor()`（渲染出的页面 PNG）同病。
//
// 方案：身份 = **规范化源路径指纹**（区分路径）+ **源内容 SHA-256**（区分内容）。
//   - 路径指纹纯字符串运算、零 IO，直接进缓存文件名 → 同名不同路径不再撞车；
//   - 内容哈希优先复用解析时已经读进内存的 `buf`（loadModel 的 pdf 分支），
//     没有 buf 时按 1MB 分块流式同步哈希，绝不把大文件再整本读一遍；
//   - 两者都写进 sidecar manifest，读取时逐项核对：**缺字段、路径不符、内容不符一律整份作废**，
//     绝不静默复用旧文本（`parser:` 版本规则原样保留，三者是"与"的关系）。
const SOURCE_IDENTITY_MEMO = new Map()
const SOURCE_IDENTITY_MEMO_MAX = 8

/** 规范化路径指纹（sha256 全 64 位十六进制）。纯字符串运算，零 IO。 */
export function pathHashOf(file) {
  return createHash('sha256').update(resolvePath(String(file)).toLowerCase()).digest('hex')
}

/** 路径指纹前 8 位：只用于给缓存文件名去重（`<name>-<key8>.ocr.md`）。 */
export function pathKeyOf(file) {
  return pathHashOf(file).slice(0, 8)
}

/** 同步分块 SHA-256：不把整本读进内存（大 PDF 的兜底路径）。 */
function hashFileSync(file) {
  const h = createHash('sha256')
  const fd = openSync(file, 'r')
  try {
    const chunk = Buffer.allocUnsafe(1 << 20)
    for (;;) {
      const n = readSync(fd, chunk, 0, chunk.length, null)
      if (n <= 0) break
      h.update(chunk.subarray(0, n))
    }
  } finally { closeSync(fd) }
  return h.digest('hex')
}

function rememberIdentity(key, id) {
  SOURCE_IDENTITY_MEMO.delete(key)
  SOURCE_IDENTITY_MEMO.set(key, id)
  while (SOURCE_IDENTITY_MEMO.size > SOURCE_IDENTITY_MEMO_MAX) {
    SOURCE_IDENTITY_MEMO.delete(SOURCE_IDENTITY_MEMO.keys().next().value)
  }
  return id
}

/**
 * 源文件身份（同步）：`{ pathHash, contentHash, size, mtimeMs }`；读不到 → `null`。
 *
 * memo 以 `路径|size|mtime` 为键：同一进程内对同一份文件只哈希一次；文件一被改动
 * （size/mtime 变）键就变，等于自动失效 —— 与 `readPdfMemo` 同款口径。
 */
export function sourceIdentityOf(file) {
  let st
  try { st = statSync(file) } catch { return null }
  if (!st.isFile()) return null
  const key = `${file}|${st.size}|${st.mtimeMs}`
  const hit = SOURCE_IDENTITY_MEMO.get(key)
  if (hit) return rememberIdentity(key, hit)
  let contentHash
  try { contentHash = hashFileSync(file) } catch { return null }
  return rememberIdentity(key, { pathHash: pathHashOf(file), contentHash, size: st.size, mtimeMs: st.mtimeMs })
}

/**
 * 解析路径专用：字节已经读进内存了，直接哈希并登记，避免随后再读一次盘。
 * `buf` 与盘上长度不一致（解析期间文件被换掉）→ 不信 buf，回落到同步哈希盘上内容。
 */
export function registerSourceIdentity(file, buf) {
  try {
    const st = statSync(file)
    if (!st.isFile() || st.size !== buf.length) return sourceIdentityOf(file)
    const contentHash = createHash('sha256').update(buf).digest('hex')
    return rememberIdentity(`${file}|${st.size}|${st.mtimeMs}`,
      { pathHash: pathHashOf(file), contentHash, size: st.size, mtimeMs: st.mtimeMs })
  } catch { return sourceIdentityOf(file) }
}

/**
 * sidecar 兜底文件：与源文件同名 + `.read.md`（设了 DSH_OFFICE_CACHE_DIR 则落到该目录）。
 * 集中目录 / 临时回退目录带**路径指纹后缀** —— 不同目录下的同名文件不再互相覆盖
 * （同目录 sidecar 路径本身已含完整路径，保持原有命名不动，逐字兼容旧调用方）。
 */
function readSidecarPath(file) {
  const st = cacheDirState()
  if (st.dir) return join(st.dir, `${basename(file)}-${pathKeyOf(file)}.read.md`)
  const side = `${file}.read.md`
  try {
    if (existsSync(dirname(side))) return side
  } catch { /* fall through */ }
  return join(tmpdir(), 'dsh-office-read', `${basename(file).replace(/[^\w.-]/g, '_')}-${pathKeyOf(file)}.read.md`)
}

/**
 * 把正文转存到 sidecar，返回一个**本身保证可无损序列化**的窄响应：
 * 路径 + stats + 首部摘录 + 续读提示。这是"序列化仍异常 / OCR 不可用"的兜底，
 * 绝不再把裸错误或整篇乱码抛给模型。
 */
function writeReadSidecar(file, content, meta = {}) {
  const p = readSidecarPath(file)
  try {
    mkdirSync(dirname(p), { recursive: true })
    const manifest = [`src: ${basename(file)}`, `fallback: ${meta.fallback || 'sidecar'}`]
    if (meta.pages) manifest.push(`pages: ${meta.pages}`)
    // 任务一：兜底转存件也记身份 —— 文件名已带路径指纹，身份行让"这份正文属于哪份内容"
    // 可被人工与后续工具核对（源文件读不到就不写，绝不编造）。
    const id = meta.identity !== undefined ? meta.identity : sourceIdentityOf(file)
    if (id) manifest.push(`srcpath: ${id.pathHash}`, `srcsha256: ${id.contentHash}`)
    const cd = cacheDirState()
    if (cd.dir) manifest.push(`dir: ${cd.dir}（DSH_OFFICE_CACHE_DIR）`)
    if (meta.reasons?.length) manifest.push(`why: ${meta.reasons.join('；')}`)
    // 任务二：sidecar 也走原子写 —— 半截 `.read.md` 会被当成"整篇正文"读给模型，
    // 半截 `.ocr.md`（有 OCR_MARK 但 parser/身份行被截掉）会把**整份花钱识别成果**判为作废。
    writeFileAtomicSync(p, `<!-- dsh-office read sidecar | ${manifest.join(' | ')} -->\n`
      + `# ${basename(file)}（正文由 dsh-office 转存，可安全删除）\n\n${content}\n`)
    return p
  } catch {
    return undefined
  }
}

/**
 * office_read 的收尾门：正文仍不可信时走 sidecar 兜底。
 * 只在"真的读不出可用正文"时触发，干净文件逐字不变（不写 quality 字段）。
 *
 * DSH 侧修正（第二轮）：`out.content` 可能已被内联护栏截成**前缀**，直接拿它写 sidecar
 * 会让 sidecar 成为"截断的截断"，而 notice 里承诺的是"整篇正文"——那就成了假话。
 * 所以 readPdf 把**截断前的完整正文**挂在临时字段 `__fullBody` 上带进来，这里用完立刻删除
 * （绝不让它出现在返回值里），sidecar 落的始终是全文。
 */
function finishRead(out, args, file) {
  if (!out || typeof out.content !== 'string') return out
  const fullBody = typeof out.__fullBody === 'string' ? out.__fullBody : ''
  if (fullBody) delete out.__fullBody
  if (args.as === 'json' || args.as === 'meta') return out
  const q = textQuality(out.content)
  const stats = out.stats && typeof out.stats === 'object' && !Array.isArray(out.stats) ? out.stats : {}
  // 缓存目录回退必须记账（"绝不静默"）：只在 DSH_OFFICE_CACHE_DIR 设了却不可写时出现
  const cd = cacheDirState()
  if (cd.note) stats.cacheDirNote = cd.note
  const noisy = q.garbled || q.replacement || q.privateUse || q.control
  if (noisy) {
    stats.quality = {
      chars: q.chars,
      cjkRatio: Number(q.cjkRatio.toFixed(4)),
      replacementRatio: Number(q.replacementRatio.toFixed(4)),
      privateUseRatio: Number(q.privateUseRatio.toFixed(4)),
      controlRatio: Number(q.controlRatio.toFixed(4)),
      // 需求 4 的增量字段：非词典字符占比（mojibake/CID 错映射的字符集信号）
      oddCharRatio: Number((q.oddCharRatio ?? 0).toFixed(4)),
    }
    if (q.reasons.length) stats.quality.reasons = q.reasons
  }
  if (!q.garbled) {
    out.stats = stats
    return out
  }

  // —— sidecar 兜底 ——
  // sidecar 必须是**全文**：内联护栏截没截过，落盘的转存件都不受它影响（见函数头注释）。
  const sidecarBody = fullBody.length > out.content.length ? fullBody : out.content
  const side = writeReadSidecar(file, sidecarBody, { fallback: 'sidecar', reasons: q.reasons, pages: args.pages })
  const excerpt = cutAtBoundary(out.content, 800)
  const howto = `office_read path="${file}" as="markdown" ocr="always" ocrEngine="local" pages="1-5"`
  stats.fallback = 'sidecar'
  if (side) { stats.sidecar = side; stats.sidecarChars = sidecarBody.length }
  else stats.sidecarNote = 'sidecar 写入失败（目录不可写）'
  const why = q.reasons.length ? q.reasons.join('；') : '正文质量门判定为不可读'
  // R9-3：降级正文只保留**头部摘录**（上面的 cutAtBoundary(out.content, 800)），而正常路径
  // 把两条"上限/预览"脚注挂在正文**尾部** —— 一截断，脚注就随尾部一起蒸发，调用方只看到
  // sidecar 说明，看不到"哪几页被上限砍掉"。这里从 stats 记账重建同一批脚注
  // （文案唯一源 ocrCapNotes()，与正常路径共用），接在 sidecar 说明**之后**、同样用 `；` 分隔。
  // 去重：正文极短时脚注可能整条落在摘录里，已出现过的就不再追加（否则同一句出现两次）。
  const capNotes = ocrCapNotes(stats.ocrPagesCapped, stats.ocrPreview, file)
    .filter(n => !excerpt.includes(n))
  const msg = `正文质量门判定为不可读（${why}）。已把整篇正文转存到 sidecar${side ? `：${side}` : ''}——`
    + `直接用 read 工具读该文件即可拿到真实换行的全文（不受内联上限影响）。`
    + `若想就地重识别，改用：${howto}`
  out.stats = stats
  out.fallback = 'sidecar'
  if (side) out.sidecar = side
  out.truncated = true
  // 截断说明（任务五）可能已占着 notice 字段：合并而不是覆盖
  out.notice = typeof out.notice === 'string' && out.notice ? `${out.notice}；${msg}` : msg
  if (capNotes.length) out.notice += `；${capNotes.join('；')}`
  out.content = `${excerpt}\n\n> ${out.notice}`
  return out
}

/**
 * 该文件 sidecar（OCR 缓存）的现状 —— **任务六-1（DSH 补充，对应宿主实测 F2/F3）**。
 *
 * 病灶：DSH 宿主里 sidecar 是**照常落盘**的（连"会话只读"的附件目录都写得进去），
 * 所以"调用失败但 OCR 成果已经在盘上"是常态。上一轮会话只能靠自己去 grep
 * `covered:` 行才发现进度 —— 这条把它变成官方输出：路径 + covered + 还缺哪些页 + grep 提示。
 *
 * 纯只读探测、绝不抛：非 PDF / sidecar 不存在 / 读不出都返回空串。
 */
function sidecarCoverage(file, identity) {
  try {
    const store = readOcrCache(file, identity)
    // 新命名路径没有、旧命名有（或反之）时，以**实际读到的那份**为准 ——
    // 提示里给出的路径必须是盘上真有的那个，否则用户 grep 半天什么也找不到。
    const path = store.path || ocrCachePath(file)
    if (!existsSync(path)) return ''
    // 旧解析器 / 身份不符的 sidecar：盘上有文件，但页面成果不可信 → 报"已作废"，不报 covered
    if (store.stale) return staleCacheNote({ path, parser: store.parser, reason: store.staleReason })
    const have = [...store.pages.keys()].sort((a, b) => a - b)
    if (!have.length) return ''
    const bits = [`covered: ${pageRanges(have)}`]
    if (store.total) {
      bits.push(`total: ${store.total}`)
      const miss = []
      for (let n = 1; n <= store.total; n++) if (!store.pages.has(n)) miss.push(n)
      if (miss.length) bits.push(`缺页=${pageRanges(miss)}`)
    }
    return `sidecar=${path}（${bits.join(' | ')}）—— OCR 成果已在盘上，可直接 grep 'covered:' 该文件头部看进度`
  } catch { return '' }
}

/**
 * office_read 错误分支的四要素：页码、格式、根因、可直接复制的下一步参数。
 * 目标是让模型**永远不再看到裸的 `invalid output`**。
 * DSH 补充（F2）：若这次失败**已产生副作用**（sidecar 已写 / 已部分写），
 * 错误文本必须一并给出 sidecar 路径与 covered / 缺页区间 —— 用户有权立刻知道它在哪。
 */
function readErrorHint(error, args, kind, file) {
  const pageRange = args.pages || '全部'
  const root = safeMessage(error)
  let next
  if (/无文本层|未提取到文本层/.test(root)) next = `office_read path="${file}" ocr="always" ocrEngine="local" pages="1-5"`
  else if (/encrypt|加密|密码/i.test(root)) next = 'office_convert --source=<源> --target=<新>.pdf（去掉打开密码后重试），或先用原程序另存一份'
  else if (/页码范围/.test(root)) next = `office_read path="${file}" as="meta"   # 先看总页数，再按实际页数传 pages`
  else if (/文件不存在|不是普通文件/.test(root)) next = '确认路径（支持绝对路径或相对会话工作目录）后重试'
  else if (/文件过大/.test(root)) next = '先用 office_convert 拆分成多份，或改读 as="meta" 只看结构'
  // 任务三：解压体积护栏命中时给"是护栏还是文件坏了"的明确指引（文案刻意不含"文件过大"）
  else if (/zip 条目解压后过大|zip 累计解压体积超限|zip 炸弹防护/.test(root)) {
    next = '该文件解压后体积异常（疑高压缩比包）；确认来源可信后可用 '
      + 'DSH_OFFICE_ZIP_MAX_ENTRY_BYTES / DSH_OFFICE_ZIP_MAX_TOTAL_BYTES 临时放宽（0=关闭该上限），'
      + '或先用原程序另存一份再读'
  }
  else next = `office_read path="${file}" as="meta"   # 先看结构，再决定 pages/sheet/offset`
  const cov = sidecarCoverage(file)
  return `【读取失败｜四要素】`
    + `页码=${pageRange}；格式=${kind || extOf(file) || '未知'}；根因=${root}；下一步=${next}`
    + (cov ? `\n${cov}` : '')
}

/** 大文档的分批建议（给 as="meta" 用）。 */
function readSuggestion(kind, model, extra, args, totalChars) {
  const batchOf = (list, per) => {
    const out = []
    for (let i = 0; i < list.length; i += per) out.push(list.slice(i, i + per))
    return out
  }
  if (kind === 'pdf' && extra?.pages > 12) {
    const per = 15
    const est = extra.pages ? Math.round(totalChars / extra.pages) : 0
    const pages = Array.from({ length: extra.pages }, (_, i) => i + 1)
    return {
      note: `${extra.pages} 页 / ${totalChars} 字符：一次整本读会撞内联上限（默认 ${maxInlineChars() === Infinity ? '已关闭' : maxInlineChars()} 字符），建议分批`,
      pageBatch: per,
      estCharsPerBatch: est * per,
      plan: batchOf(pages, per).map(g => (g.length === 1 ? `pages:"${g[0]}"` : `pages:"${g[0]}-${g[g.length - 1]}"`)),
      copy: `office_read path="${args?.path || ''}" as="markdown" pages="1-${per}"`,
    }
  }
  if (kind === 'workbook' && model) {
    try {
      const w = normalizeWorkbook(model)
      const big = w.sheets.filter(s => s.rows.length > 1000)
      if (big.length) {
        return {
          note: `有 ${big.length} 个工作表行数超过 1000，建议用 sheet + offset/limit 续读`,
          sheets: big.map(s => ({ name: s.name, rows: s.rows.length })),
          copy: `office_read path="${args?.path || ''}" as="markdown" sheet="${big[0].name}" offset=0 limit=20000`,
        }
      }
    } catch { /* best-effort */ }
    return undefined
  }
  if (kind === 'document' && totalChars > 120000) {
    return {
      note: `${totalChars} 字符：建议用 offset/limit 分段读`,
      copy: `office_read path="${args?.path || ''}" as="markdown" offset=0 limit=20000`,
    }
  }
  return undefined
}

/**
 * 文本层质量画像（**纯 CPU**：只消费已解析文本，不渲染 PNG、不触发 OCR）。
 * 基线与 readPdf 正文路径同一套（本批全文），CJK 覆盖率只作相对判据 —— 这样
 * "meta 说的" 与 "read 实际遇到的" 不会出现两套口径。
 * 本就无文本层的页**不计入** testedPages：那是 scannedPages 的语义，不是质量门命中。
 *
 * @returns {{testedPages:number, garbled:number[], reasons:string[]}}
 */
function textLayerProfile(sections) {
  const baseline = textQuality(sections.map(s => String(s.text || '')).join('\n'))
  const tested = []
  const garbled = []
  const reasons = []
  for (const s of sections) {
    const raw = String(s.text || '')
    if (!raw.replace(/\s/g, '').length) continue
    tested.push(s.page)
    // 需求 4：页面级质量门启用**字符级/版面级启发**（structural）——
    // "字符可提取但语义破碎"的页在旧口径下是漏检的。基线仍是整本同算，相对判据不吃亏。
    const q = textQuality(raw, baseline, { structural: true })
    if (!q.garbled) continue
    garbled.push(s.page)
    if (reasons.length < 3) reasons.push(`第 ${s.page} 页：${q.reasons.join('、')}`)
  }
  return { testedPages: tested.length, garbled, reasons }
}

/** 乱码页清单的外显形状：非空 → `"a-b,c"`；空 → `[]`。顶层与 qualityGate 共用同一口径。 */
function garbledSpec(pages) {
  return pages.length ? pageRanges(pages) : []
}

/**
 * `office_convert` 的出站质量门（G1：消灭"静默坏输出"）。
 *
 * 病灶：文字层整本 CID 乱码时，旧版 convert 会把 5 万字节的二进制垃圾（含 NUL 与
 * C0 控制符）写进目标文件，而返回 JSON 只有 `{source,target,bytes}` —— 零告警。
 * 这里在**写盘之前**判一次：乱码页有完整 OCR 缓存就替换后继续，没有就拒绝落盘。
 *
 * DSH 补充（任务一-2 / F2 F3）：除了"缺不缺缓存"，还要把 **sidecar 路径 + 当前 covered /
 * 缺页区间**一并带出来 —— 宿主里 sidecar 是会落盘的，用户有权立刻知道它在哪、跑到第几页。
 *
 * @returns {{garbled:number[], reasons:string[], fromCache:boolean, cache:Map|null,
 *            missing:number[], covered:number[], total:number|undefined, cachePath:string}}
 */
function convertSourceGuard(file, kind, extra) {
  const none = { garbled: [], reasons: [], fromCache: false, cache: null, missing: [], covered: [], total: undefined, cachePath: '', staleCache: null }
  if (kind !== 'pdf' || !Array.isArray(extra?.sections)) return none
  const prof = textLayerProfile(extra.sections)
  if (!prof.garbled.length) return none
  const store = readOcrCache(file, extra.identity)
  const cache = store.pages
  const missing = prof.garbled.filter(p => !String(cache.get(p) || '').trim())
  return {
    garbled: prof.garbled,
    reasons: prof.reasons,
    fromCache: missing.length === 0,
    cache,
    missing,
    covered: [...cache.keys()].sort((a, b) => a - b),
    total: extra.pages,
    cachePath: ocrCachePath(file),
    staleCache: store.stale ? { path: store.path, parser: store.parser, reason: store.staleReason } : null,
  }
}

/** 转换被质量门拦下时的四要素错误（页码 / 格式 / 根因 / 可复制的下一步 + sidecar 现状）。 */
function convertRefusalError(file, guard, kind = 'pdf') {
  const pages = guard.garbled.length ? pageRanges(guard.garbled) : '全文'
  const why = guard.reasons.length ? guard.reasons.join('；') : '渲染出的正文未通过质量门'
  // 下一步只针对**还缺的页**给参数串（缓存已覆盖的不用重跑）；单批 ≤20 页与工具层上限一致。
  const want = (guard.missing?.length ? guard.missing : guard.garbled).slice(0, 20)
  const nextPages = want.length ? pageRanges(want) : '1-20'
  const cov = guard.cachePath
    ? `；sidecar=${guard.cachePath}（covered: ${guard.covered?.length ? pageRanges(guard.covered) : '无'}`
      + `${guard.total ? ` / total: ${guard.total}` : ''}）`
      + (guard.missing?.length ? `，仍缺第 ${pageRanges(guard.missing)} 页` : '')
      + ` —— 可直接 grep 'covered:' 该文件头部看进度；`
    : ''
  // 缓存"凭空消失"必须解释：旧解析器写下的 sidecar 已被整份作废（任务 D 的迁移协议）
  const stale = guard.staleCache ? `；${staleCacheNote(guard.staleCache)}` : ''
  return '【转换拒绝｜四要素】'
    + `页码=${pages}；格式=${kind}；根因=文字层不可信（${why}），且没有完整 OCR 缓存 —— `
    + '继续转换只会把不可读正文写进目标文件，故已拒绝且未产出任何文件'
    + cov
    + stale
    + `下一步=先识别再转：office_read path="${file}" ocr="always" ocrEngine="local" pages="${nextPages}"`
    + `（结果缓存为同名 .ocr.md；分批跑到覆盖全部乱码页），然后重跑本次转换`
}

/** 模型里的正文字符数（分批建议用，尽力而为）。 */
function modelCharCount(model) {
  if (typeof model === 'string') return model.length
  try {
    if (model?.kind === 'document') {
      return normalizeDocument(model).blocks.reduce((a, b) => a
        + String(b.text || '').length
        + (b.runs || []).reduce((x, r) => x + String(r.text || '').length, 0)
        + (b.rows || []).reduce((x, row) => x + row.join('').length, 0), 0)
    }
    if (model?.kind === 'workbook') {
      return normalizeWorkbook(model).sheets.reduce((a, s) => a + s.rows.reduce((x, r) => x + r.join('').length, 0), 0)
    }
    if (model?.kind === 'slides') {
      return normalizeSlides(model).slides.reduce((a, s) => a
        + String(s.title || '').length + String(s.subtitle || '').length + String(s.notes || '').length
        + (s.bullets || []).reduce((x, b) => x + String(b.text || b || '').length, 0), 0)
    }
  } catch { /* best-effort */ }
  return 0
}

// ---------------------------------------------------------------------------
// readers: file → unified content model
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// PDF 全量解析 memo
// ---------------------------------------------------------------------------

/**
 * `readPdfText` 是大 PDF 最贵的一步（整本解析），而 `as="meta"`、`pages` 局部读、
 * 连续分批续读都会把它重跑一遍。这里做 LRU（≤3）：key 带 mtime+size，
 * 文件一改 key 自动失效，绝不会把旧内容当新内容返回。
 * 缓存对象只有只读消费方（pdfToDocument 复制出新的 blocks，sections 之后不再变异），共享是安全的。
 */
const PDF_MEMO_MAX = 3
const pdfMemo = new Map()
let pdfMemoHits = 0
let pdfMemoMisses = 0

/** 测试用账本：`{ size, hits, misses }`。 */
export function pdfMemoStats() {
  return { size: pdfMemo.size, hits: pdfMemoHits, misses: pdfMemoMisses }
}

/** 命中免解析；文件 mtime/size 一变 key 就变，等于自动失效。 */
async function readPdfMemo(file, buf) {
  let key = file
  try {
    const st = await stat(file)
    key = `${file}|${st.mtimeMs}|${st.size}`
  } catch { /* stat 失败就退化成按路径缓存 */ }
  const hit = pdfMemo.get(key)
  if (hit) {
    pdfMemoHits++
    pdfMemo.delete(key)          // LRU：命中即移到队尾
    pdfMemo.set(key, hit)
    return hit
  }
  const pdf = readPdfText(buf)
  pdfMemoMisses++
  pdfMemo.set(key, pdf)
  while (pdfMemo.size > PDF_MEMO_MAX) pdfMemo.delete(pdfMemo.keys().next().value)
  return pdf
}

async function loadModel(file, exec) {
  const buf = await readBuffer(file)
  throwIfAborted(exec)
  const ext = extOf(file)
  // 需求 1：.html/.htm 走零依赖 DOM 解析（html.js），产出与其他来源同构的 document 模型 ——
  // meta/markdown/json/text 四种 as 与 office_convert 的全部目标格式都自动继承。
  // 纯容错：即便文件没有任何标签（其实是纯文本），也会按段落读出，绝不因此失败。
  if (ext === 'html' || ext === 'htm') {
    const text = decodeTextBytes(buf)
    return { kind: 'html', model: htmlToDocument(text), extra: { sourceChars: text.length } }
  }
  const container = sniff(buf, ext)
  const zipObj = container === 'zip' ? openZip(buf) : null
  const kind = resolveKind(buf, ext, zipObj ? zipObj.names : null)
  switch (kind) {
    case 'docx':
      return { kind, model: readDocx(zipObj) }
    case 'xlsx':
      return { kind, model: readXlsx(zipObj) }
    case 'pptx':
      return { kind, model: readPptx(zipObj) }
    case 'odt':
      return { kind, model: readOdt(zipObj) }
    case 'ods':
      return { kind, model: readOds(zipObj) }
    case 'odp':
      return { kind, model: readOdp(zipObj) }
    case 'pdf': {
      const pdf = await readPdfMemo(file, buf)
      // 任务一：解析时源字节已经在这只 buf 里了 —— 顺手把身份（路径指纹 + 内容 SHA-256）
      // 算出来登记，下游的 readOcrCache / writeOcrCache 直接用，**不再为大文件重复读盘**。
      const identity = registerSourceIdentity(file, buf)
      return { kind, model: pdfToDocument(pdf), extra: { pages: pdf.pages, sections: pdf.sections, textFound: pdf.textFound, identity } }
    }
    case 'doc':
      return { kind, model: readLegacyDoc(parseOle2(buf)) }
    case 'xls':
      return { kind, model: readLegacyXls(parseOle2(buf)) }
    case 'ppt':
      return { kind, model: readLegacyPpt(parseOle2(buf)) }
    case 'csv':
    case 'tsv': {
      const text = decodeTextBytes(buf)
      const rows = parseDelimited(text, ext === 'tsv' ? '\t' : undefined)
      return {
        kind,
        model: { kind: 'workbook', meta: {}, sheets: [{ name: 'Sheet1', rows: rows.map(r => r.map(v => ({ v, t: typeofVal(v) }))) }] },
        extra: { rows },
      }
    }
    case 'rtf':
      return { kind, model: { kind: 'document', meta: {}, blocks: textToBlocks(rtfToText(decodeTextBytes(buf))) } }
    case 'md':
      return { kind, model: markdownToDocument(decodeTextBytes(buf)) }
    case 'text':
    case 'txt':
    case 'json':
    case 'jsonl':
      return { kind: 'text', model: decodeTextBytes(buf) }
    case 'ole2': {
      const ole = parseOle2(buf)
      if (ole.has('WordDocument')) return { kind: 'doc', model: readLegacyDoc(ole) }
      if (ole.has('Workbook') || ole.has('Book')) return { kind: 'xls', model: readLegacyXls(ole) }
      if (ole.has('PowerPoint Document')) return { kind: 'ppt', model: readLegacyPpt(ole) }
      throw new Error(`无法识别的 OLE2 复合文档（内部流: ${ole.names.slice(0, 8).join(', ')}）`)
    }
    default:
      if (kind === 'ooxml-unknown') throw new Error('zip 内未发现 word/xl/ppt 目录，可能不是有效的 OOXML 文件')
      throw new Error(`暂不支持该格式: ${ext ? '.' + ext : '(无扩展名)'}`)
  }
}

function pdfToDocument(pdf) {
  const blocks = []
  const multi = pdf.sections.length > 1
  for (const sec of pdf.sections) {
    if (multi) blocks.push({ type: 'heading', level: 2, text: `第 ${sec.page} 页` })
    for (const line of sec.text.split('\n')) {
      if (!line.trim()) { blocks.push({ type: 'paragraph', runs: [{ text: '' }] }); continue }
      if (line.startsWith('# ')) blocks.push({ type: 'heading', level: 3, text: line.slice(2) })
      else blocks.push({ type: 'paragraph', runs: [{ text: line }] })
    }
  }
  return { kind: 'document', meta: pdf.meta || {}, blocks }
}

function typeofVal(v) {
  return /^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(String(v).trim()) ? 'n' : 's'
}

function rtfToText(rtf) {
  let out = ''
  let i = 0
  const n = rtf.length
  let skip = 0
  let ucSkip = 1
  let codepage = 1252
  const decoders = new Map()
  const decoderFor = codePage => {
    if (!decoders.has(codePage)) {
      // WHATWG labels only: map the numeric code page to a supported label.
      const LABELS = {
        65001: 'utf-8', 936: 'gbk', 950: 'big5', 932: 'shift_jis', 949: 'euc-kr',
        874: 'windows-874', 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252',
        1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256',
        1257: 'windows-1257', 1258: 'windows-1258', 10000: 'macintosh', 20866: 'koi8-r', 21866: 'koi8-u',
      }
      let d
      try { d = new TextDecoder(LABELS[codePage] ?? 'windows-1252') } catch { d = new TextDecoder('windows-1252') }
      decoders.set(codePage, d)
    }
    return decoders.get(codePage)
  }
  const SKIP_DESTS = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'generator', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'listtable', 'listoverridetable', 'revtbl', 'nonshppict', 'xmlnstbl', 'rsidtbl', 'mmathPr'])
  while (i < n) {
    const c = rtf[i]
    if (c === '{') { i++; continue }
    if (c === '}') { if (skip > 0) skip--; i++; continue }
    if (c === '\\') {
      const next = rtf[i + 1]
      if (next === "'") {
        const hex = rtf.slice(i + 2, i + 4)
        if (/^[0-9a-fA-F]{2}$/.test(hex)) {
          if (!skip) out += decoderFor(codepage).decode(Uint8Array.from([parseInt(hex, 16)]))
          i += 4
          continue
        }
      }
      if (next === '\\' || next === '{' || next === '}') {
        if (!skip) out += next
        i += 2
        continue
      }
      const m = /^\\([a-zA-Z]+)(-?\d{1,10})? ?/.exec(rtf.slice(i))
      if (m) {
        const word = m[1]
        const num = m[2] !== undefined ? Number(m[2]) : undefined
        const advance = m[0].length
        if (word === 'ansicpg' && num) codepage = num
        if (word === 'uc' && num !== undefined) ucSkip = num
        if (word === 'u' && num !== undefined) {
          if (!skip) out += String.fromCharCode(num < 0 ? num + 65536 : num)
          i += advance
          let skipped = 0
          while (skipped < ucSkip && i < n) {
            if (rtf[i] === '\\' && rtf[i + 1] === "'") { i += 4; skipped++; continue }
            if (/[^{}\\\r\n]/.test(rtf[i])) { i++; skipped++; continue }
            break
          }
          continue
        }
        if (!skip && SKIP_DESTS.has(word)) {
          i += advance
          let depth = 1
          while (i < n && depth > 0) {
            if (rtf[i] === '{') depth++
            else if (rtf[i] === '}') depth--
            i++
          }
          continue
        }
        if (!skip) {
          if (word === 'par' || word === 'line' || word === 'sect') out += '\n'
          else if (word === 'tab' || word === 'cell' || word === 'row') out += '\t'
          else if (word === 'space') out += ' '
          else if (word === 'emdash') out += '—'
          else if (word === 'endash') out += '–'
        }
        i += advance
        continue
      }
      i += 2
      continue
    }
    if (!skip && c !== '\r' && c !== '\n') out += c
    i++
  }
  return out.replace(/\n{3,}/g, '\n\n')
}

// ---------------------------------------------------------------------------
// model → markdown / text
// ---------------------------------------------------------------------------

function modelToMarkdown(model) {
  if (model && model.kind === 'document') return documentToMarkdown(model)
  if (model && model.kind === 'workbook') return workbookToMarkdown(model)
  if (model && model.kind === 'slides') return slidesToMarkdown(model)
  return String(model ?? '')
}

function modelToText(model) {
  return modelToMarkdown(model)
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/^\|[\s:|-]+\|$/gm, '')
    .replace(/\\\|/g, '|')
}

function workbookToMarkdown(wb) {
  const w = normalizeWorkbook(wb)
  const out = []
  for (const sh of w.sheets) {
    if (w.sheets.length > 1) out.push(`## 工作表: ${sh.name}`)
    if (!sh.rows.length) { out.push('_(空表)_'); continue }
    out.push(tableMd(sh.rows.map(r => r.map(cellText)), false))
  }
  return out.join('\n')
}

function slidesToMarkdown(sl) {
  const s = normalizeSlides(sl)
  const out = []
  s.slides.forEach((slide, i) => {
    out.push(`## 幻灯片 ${i + 1}${slide.title ? `: ${slide.title}` : ''}`)
    if (slide.subtitle) out.push(`*${slide.subtitle}*`)
    for (const b of slide.bullets || []) out.push(`${'  '.repeat(Math.min(8, b.level || 0) + 1)}- ${b.text}`)
    if (slide.table?.rows) out.push(tableMd(slide.table.rows.map(r => r.map(cellText)), false))
    if (slide.notes) out.push(`> 备注: ${slide.notes}`)
    if (slide.images) out.push(`_含 ${slide.images} 张图片_`)
    out.push('')
  })
  return out.join('\n')
}

function cellText(c) {
  const v = c && typeof c === 'object' ? c.v : c
  return String(v ?? '').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')
}

// ---------------------------------------------------------------------------
// model → target file bytes
// ---------------------------------------------------------------------------
//
// 需求 5（CSV/文本写出的编码收敛）：文本类写出统一走 `textOut()` ——
//   csv     默认 **UTF-8 with BOM**（Excel 双击直接正确识别中文；旧行为已是如此，本轮参数化）
//   tsv/md/txt 默认无 BOM（旧行为不变），`encoding: "utf-8-sig"` 可显式加 BOM
// `encoding: "utf-8"` 显式表示无 BOM。gb18030 等旧代码页的**写出**不是零依赖可达的
// （Node Buffer 只内建 utf8/utf16le/latin1 系），维持不提供；**读取**侧本来就会自动
// 识别 GBK/Big5（decodeTextBytes），不受影响。

const textOut = (s, encoding, defaultBom = false) => {
  const bom = defaultBom ? encoding !== 'utf-8' : encoding === 'utf-8-sig'
  return Buffer.from((bom ? '\ufeff' : '') + s, 'utf8')
}

const WRITERS = {
  docx: (m, _enc, info) => writeDocx(m, info && typeof info === 'object' ? { info } : {}),
  odt: m => writeOdt(m),
  pdf: (m, _enc, info) => writePdf(m, info && typeof info === 'object' ? { info } : {}),
  xlsx: m => writeXlsx(m),
  ods: m => writeOds(m),
  pptx: m => writePptx(m),
  odp: m => writeOdp(m),
  md: (m, enc) => textOut(modelToMarkdown(m), enc),
  txt: (m, enc) => textOut(modelToText(m), enc),
  csv: (m, enc) => textOut(renderDelimited(csvRowsOf(m), ','), enc, true),
  tsv: (m, enc) => textOut(renderDelimited(csvRowsOf(m), '\t'), enc),
  json: m => Buffer.from(JSON.stringify(plainJson(m), null, 2), 'utf8'),
  html: m => Buffer.from(documentToHtml(m), 'utf8'),
}

function csvRowsOf(model) {
  const w = normalizeWorkbook(model && model.kind === 'workbook' ? model : { sheets: [] })
  const rows = []
  for (const sh of w.sheets) {
    if (w.sheets.length > 1) rows.push([`# ${sh.name}`])
    for (const r of sh.rows) rows.push(r.map(c => String((c && typeof c === 'object' ? c.v : c) ?? '')))
  }
  return rows
}

function plainJson(model) {
  return JSON.parse(JSON.stringify(model ?? null))
}

const KIND_TARGET = {
  docx: 'document', odt: 'document', pdf: 'document', md: 'document', txt: 'document', html: 'document',
  xlsx: 'workbook', ods: 'workbook', csv: 'workbook', tsv: 'workbook', json: 'workbook',
  pptx: 'slides', odp: 'slides',
}

function adaptModel(model, targetExt) {
  const want = KIND_TARGET[targetExt]
  if (!want) return model
  if (model && model.kind === want) return model
  if (want === 'document') {
    if (model && model.kind === 'workbook') return workbookToDocumentModel(model)
    if (model && model.kind === 'slides') return slidesToDocumentModel(model)
    if (typeof model === 'string') return { kind: 'document', meta: {}, blocks: textToBlocks(model) }
    return { kind: 'document', meta: {}, blocks: [] }
  }
  if (want === 'workbook') {
    if (model && model.kind === 'document') return documentToWorkbookModel(model)
    if (model && model.kind === 'slides') return slidesToWorkbookModel(model)
    if (typeof model === 'string') return normalizeWorkbook({ sheets: [{ name: 'Sheet1', rows: model.split('\n').map(l => [l]) }] })
    return normalizeWorkbook({ sheets: [{ name: 'Sheet1', rows: [] }] })
  }
  if (want === 'slides') {
    if (model && model.kind === 'document') return documentToSlidesModel(model)
    if (model && model.kind === 'workbook') return workbookToSlidesModel(model)
    if (typeof model === 'string') {
      return normalizeSlides({ slides: [{ layout: 'title', title: '文档', bullets: model.split('\n').slice(0, 40).map(t => ({ text: t, level: 0 })) }] })
    }
    return normalizeSlides({ slides: [{ layout: 'title', title: '演示文稿', bullets: [] }] })
  }
  return model
}

/**
 * 插图兜底（第十二轮需求 1a 的「不静默」面；第二轮需求 1b 把 `.docx` 移出降级名单）。
 *
 * 输入侧会把 `![alt](path)` 解析成 `image` 块。写出端里 **PDF / HTML / docx** 现在都会
 * 真正内嵌图片（docx 见 `docx.js::writeDocx` 的 image 分支）；其余目标（`.odt` / `.pptx` /
 * `.md` 等）没有 image 分支，若把 image 块原样交出去会落进 `blockToXml` 的 default 分支
 * → **什么都不输出 = 静默丢图**，直接违反"不静默"红线。
 *
 * 所以这里显式降级成字面 markdown 文本，并逐张记账；调用方把 `degraded` 写进
 * `stats.imageFallback` 与 `notice`。图片语义不丢（路径与 alt 都在文本里），只是不内嵌。
 */
const IMAGE_EMBED_TARGETS = new Set(['pdf', 'html', 'docx'])
function degradeImageBlocks(model, targetExt) {
  if (!model || model.kind !== 'document' || IMAGE_EMBED_TARGETS.has(targetExt)) return { model, degraded: [] }
  const degraded = []
  const blocks = (model.blocks || []).map(b => {
    if (b.type !== 'image') return b
    const alt = b.alt || '图片'
    const name = b.name || ''
    degraded.push({ alt, name, reason: `.${targetExt} 写出端不内嵌图片，已降级为字面文本 ![${alt}](${name})` })
    return { type: 'paragraph', runs: [{ text: `![${alt}](${name})` }] }
  })
  if (!degraded.length) return { model, degraded: [] }
  return { model: { ...model, blocks }, degraded }
}

/**
 * 把写出端（`winfo`）的图片记账并进返回的 `stats` / `notice`（第二轮需求 1b / 4c / 4e）。
 * 无图片记账时**不动** `out`（干净文件的返回值逐字不变）。
 */
function applyImageWriteInfo(out, winfo, ext) {
  const skipped = Array.isArray(winfo?.imagesSkipped) ? winfo.imagesSkipped : []
  const sizing = Array.isArray(winfo?.imageSizing) ? winfo.imageSizing : []
  const reused = Array.isArray(winfo?.imageReused) ? winfo.imageReused : []
  const media = Array.isArray(winfo?.imageMedia) ? winfo.imageMedia : []
  const notices = []
  if (media.length) {
    out.stats = { ...(out.stats || {}), imageMedia: media }
    if (sizing.length) {
      out.stats.imageSizing = sizing
      notices.push(`内嵌 ${media.length} 张图片：`
        + sizing.map(s => `${s.name || '(无路径)'} 原始 ${s.px} → ${s.pt}pt（${s.rule}${s.capped ? '，超可用宽已等比缩小' : ''}）`).join('；'))
    }
    if (reused.length) {
      out.stats.imageReused = reused
      notices.push(`${reused.length} 张图片内容与已有媒体部件相同，已复用（不重复写部件）`)
    }
  }
  if (skipped.length) {
    out.stats = { ...(out.stats || {}), imagesSkipped: skipped }
    notices.push(`${skipped.length} 张图片未能内嵌到 .${ext}（已退化为字面文本，不静默丢图）：`
      + skipped.map(s => `${s.name || '(无路径)'}（${s.reason}）`).join('；'))
  }
  if (notices.length) out.notice = [out.notice, ...notices].filter(Boolean).join('；')
  return out
}

function workbookToDocumentModel(wb) {
  const w = normalizeWorkbook(wb)
  const blocks = []
  for (const sh of w.sheets) {
    if (w.sheets.length > 1 || sh.name) blocks.push({ type: 'heading', level: 2, text: sh.name })
    if (sh.rows.length) blocks.push({ type: 'table', header: true, rows: sh.rows.map(r => r.map(c => String((c && typeof c === 'object' ? c.v : c) ?? ''))) })
  }
  return { kind: 'document', meta: w.meta || {}, blocks }
}

function slidesToDocumentModel(sl) {
  const s = normalizeSlides(sl)
  const blocks = []
  s.slides.forEach((slide, i) => {
    blocks.push({ type: 'heading', level: 2, text: slide.title || `幻灯片 ${i + 1}` })
    if (slide.subtitle) blocks.push({ type: 'paragraph', runs: [{ text: slide.subtitle }] })
    if (slide.bullets?.length) blocks.push({ type: 'list', ordered: false, items: slide.bullets })
    if (slide.table?.rows) blocks.push({ type: 'table', header: true, rows: slide.table.rows.map(r => r.map(cellText)) })
    if (slide.notes) blocks.push({ type: 'quote', text: `备注: ${slide.notes}` })
  })
  return { kind: 'document', meta: s.meta || {}, blocks }
}

function documentToWorkbookModel(doc) {
  const d = normalizeDocument(doc)
  const tables = d.blocks.filter(b => b.type === 'table')
  if (tables.length) {
    return normalizeWorkbook({
      meta: d.meta,
      sheets: tables.map((t, i) => ({ name: `表${i + 1}`, rows: t.rows.map(r => r.map(c => ({ v: String(c ?? ''), t: typeofVal(c) }))) })),
    })
  }
  const rows = []
  for (const b of d.blocks) {
    if (b.type === 'heading') rows.push([b.text])
    else if (b.type === 'paragraph') rows.push([plainOf(b.runs)])
    else if (b.type === 'list') for (const it of b.items) rows.push([it.text])
    else if (b.type === 'quote') rows.push([String(b.text ?? '')])
  }
  return normalizeWorkbook({ meta: d.meta, sheets: [{ name: 'Sheet1', rows }] })
}

function slidesToWorkbookModel(sl) {
  const s = normalizeSlides(sl)
  const rows = [['序号', '标题', '要点', '备注']]
  s.slides.forEach((slide, i) => {
    const bullets = (slide.bullets || []).map(b => '· ' + b.text).join('\n')
    const table = slide.table?.rows ? slide.table.rows.map(r => r.map(cellText).join(' | ')).join('\n') : ''
    rows.push([String(i + 1), slide.title || '', [bullets, table].filter(Boolean).join('\n'), slide.notes || ''])
  })
  return normalizeWorkbook({ meta: s.meta, sheets: [{ name: 'Slides', rows: rows.map(r => r.map(v => ({ v, t: typeofVal(v) }))) }] })
}

function documentToSlidesModel(doc) {
  const d = normalizeDocument(doc)
  const slides = []
  let current = null
  for (const b of d.blocks) {
    if (b.type === 'heading' && b.level <= 2) {
      current = { layout: slides.length === 0 ? 'title' : 'content', title: b.text, bullets: [] }
      slides.push(current)
      continue
    }
    if (b.type === 'hr') { current = null; continue }
    if (!current) {
      current = { layout: slides.length === 0 ? 'title' : 'content', title: d.meta?.title || '内容', bullets: [] }
      slides.push(current)
    }
    if (b.type === 'paragraph') current.bullets.push({ text: plainOf(b.runs), level: 0 })
    else if (b.type === 'list') for (const it of b.items) current.bullets.push(it)
    else if (b.type === 'table') current.table = { rows: b.rows }
    else if (b.type === 'quote') current.bullets.push({ text: String(b.text ?? ''), level: 0 })
  }
  const cleaned = slides.map(s => ({ ...s, bullets: (s.bullets || []).filter(b => String(b.text || '').trim()) }))
  if (!cleaned.length) cleaned.push({ layout: 'title', title: d.meta?.title || '演示文稿', bullets: [] })
  return normalizeSlides({ meta: d.meta, slides: cleaned })
}

function workbookToSlidesModel(wb) {
  const w = normalizeWorkbook(wb)
  const slides = []
  for (const sh of w.sheets) {
    const rows = sh.rows.map(r => r.map(c => String((c && typeof c === 'object' ? c.v : c) ?? '')))
    if (!rows.length) continue
    slides.push({
      layout: 'content',
      title: sh.name,
      bullets: [],
      table: { rows: rows.slice(0, 20) },
      notes: rows.length > 20 ? `完整数据共 ${rows.length} 行，此处仅显示前 20 行` : undefined,
    })
  }
  if (!slides.length) slides.push({ layout: 'title', title: '数据为空', bullets: [] })
  return normalizeSlides({ meta: w.meta, slides })
}

function plainOf(runs) {
  return (runs || []).map(r => (typeof r === 'string' ? r : r.text)).join('')
}

// ---------------------------------------------------------------------------
// zip part editing
// ---------------------------------------------------------------------------

function zipToMap(zipObj) {
  const map = new Map()
  for (const partName of zipObj.names) map.set(partName, Buffer.from(zipObj.get(partName)))
  return map
}

function mapToZip(map, order = []) {
  const seen = new Set()
  const entries = []
  const push = partName => {
    if (seen.has(partName) || !map.has(partName)) return
    seen.add(partName)
    entries.push({ name: partName, data: map.get(partName), store: partName === 'mimetype' })
  }
  for (const n of order) push(n)
  for (const n of map.keys()) push(n)
  return makeZip(entries)
}

function utf8(s) {
  return Buffer.from(s, 'utf8')
}

function coreXmlMeta(coreXmlText, update) {
  const NS = {
    'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
    'xmlns:dcterms': 'http://purl.org/dc/terms/',
    'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
  }
  const doc = parseXML(coreXmlText || '')
  let body = doc.children.find(c => typeof c !== 'string' && c.name === 'cp:coreProperties')
  if (!body) {
    body = { name: 'cp:coreProperties', attrs: { ...NS }, children: [] }
    doc.children.push(body)
  }
  for (const [key, tag] of Object.entries({ title: 'dc:title', author: 'dc:creator', subject: 'dc:subject', keywords: 'cp:keywords', description: 'dc:description' })) {
    if (update[key] === undefined) continue
    const existing = findAll(body, tag)[0]
    if (existing) existing.children = [encodeEntities(String(update[key]))]
    else body.children.push({ name: tag, attrs: {}, children: [encodeEntities(String(update[key]))] })
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serializeXML(body)}`
}

// ---------------------------------------------------------------------------
// editors
// ---------------------------------------------------------------------------

async function editDocx(file, ops, exec) {
  const buf = await readBuffer(file)
  const zipObj = openZip(buf)
  const parts = zipToMap(zipObj)
  let docXml = zipObj.getText('word/document.xml')
  if (docXml === undefined) throw new Error('docx: 缺少 word/document.xml')
  const summary = []
  const hyperlinks = []
  const imageRels = []          // {id, target, kind}
  const ctKey = '[Content_Types].xml'
  // 现有关系 id：新关系必须接着往后编，否则会与既有 hyperlink/图片撞号（Word 会判包损坏）
  const existingRids = new Set()
  {
    const relsText = zipObj.getText('word/_rels/document.xml.rels') || ''
    for (const m of relsText.matchAll(/Id="([^"]+)"/g)) existingRids.add(m[1])
  }
  // 媒体去重表（第二轮需求 4c）：内容 SHA-256 → {part, rid}，覆盖**既有** word/media 部件；
  // 同时记既有媒体序号上界，本次新增部件接着往后编（绝不覆盖既有图片）。
  const mediaByHash = new Map()
  let mediaSeqMax = 0
  {
    const relsText = zipObj.getText('word/_rels/document.xml.rels') || ''
    const target2rid = new Map()
    for (const m of relsText.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1]
      const target = /Target="([^"]+)"/.exec(m[0])?.[1]
      const type = /Type="([^"]+)"/.exec(m[0])?.[1] || ''
      if (id && target && /\/image$/.test(type)) target2rid.set(target.replace(/^\.\//, ''), id)
    }
    for (const [name, data] of parts) {
      const mm = /^word\/media\/image(\d+)\./.exec(name)
      if (mm) mediaSeqMax = Math.max(mediaSeqMax, Number(mm[1]))
      if (!/^word\/media\//.test(name)) continue
      const hash = createHash('sha256').update(Buffer.from(data)).digest('hex')
      if (!mediaByHash.has(hash)) mediaByHash.set(hash, { part: name, rid: target2rid.get(name.replace(/^word\//, '')) || null })
    }
  }
  const docxImageSkipped = []
  const docxImageSizing = []
  const docxImageReused = []
  const docxImageMedia = []
  let imageGraphSeq = 1000
  let imageAnchorMiss = ''
  for (const op of ops) {
    throwIfAborted(exec)
    if (op.op === 'replace_text') {
      const find = String(op.find ?? '')
      if (!find && !op.regex) throw new Error('replace_text 需要 find')
      const r = replaceTextInDocument(docXml, find, String(op.replace ?? ''), { regex: !!op.regex })
      docXml = r.xml
      summary.push(`replace_text: 替换 ${r.count} 处`)
    } else if (op.op === 'append_markdown') {
      const md = String(op.markdown ?? op.text ?? '')
      if (!md.trim()) throw new Error('append_markdown 需要 markdown 内容')
      const doc = markdownToDocument(md)
      const r = appendBlocksToDocument(docXml, doc.blocks, { mediaSeqStart: mediaSeqMax, byHash: mediaByHash })
      docXml = r.xml
      hyperlinks.push(...r.hyperlinks)
      for (const m of r.media || []) {
        parts.set(m.name, m.data)
        docxImageMedia.push(m.name)
        const mm = /^word\/media\/image(\d+)\./.exec(m.name)
        if (mm) mediaSeqMax = Math.max(mediaSeqMax, Number(mm[1]))
      }
      imageRels.push(...(r.imageRels || []))
      if (r.imagesSkipped?.length) docxImageSkipped.push(...r.imagesSkipped)
      if (r.imageSizing?.length) docxImageSizing.push(...r.imageSizing)
      if (r.imageReused?.length) docxImageReused.push(...r.imageReused)
      summary.push(`append_markdown: 追加 ${doc.blocks.length} 个内容块`
        + (r.media?.length ? `（内含 ${r.media.length} 张图片）` : '')
        + (r.imagesSkipped?.length ? `；${r.imagesSkipped.length} 张图片读不到，已退化为字面文本` : ''))
    } else if (op.op === 'set_meta') {
      parts.set('docProps/core.xml', utf8(coreXmlMeta(zipObj.getText('docProps/core.xml'), op)))
      summary.push(`set_meta: 更新文档属性 (${Object.keys(op).filter(k => k !== 'op').join(', ')})`)
    } else if (op.op === 'insert_image' || op.op === 'append_image') {
      // ---- 插图（第十二轮 需求 1b）----
      // zip 级新增 `word/media/imageN.*` + image 关系 + 一段 `<w:drawing>`；
      // 既有文档根元素一个字不动（命名空间就地声明在插入的那段 XML 上）。
      const src = op.path ?? op.base64 ?? op.image ?? op.dataUrl
      if (!src) throw new Error(`${op.op} 需要 path 或 base64/base64:… / data:image/…;base64,…`)
      const img = readImageBytes(src)
      const info = sniffImage(img.buf)
      if (!info || !info.width || !info.height) {
        throw new Error(`【插图】不认识的图片格式（只支持 PNG/JPEG/GIF/BMP）；`
          + `下一步=把图片另存为 PNG 或 JPEG 再试；来源=${img.from === 'base64' ? 'base64' : img.path}`)
      }
      // 内容哈希去重（需求 4c）：同图复用同一媒体部件与同一关系 id，不重复写部件
      const imgHash = createHash('sha256').update(img.buf).digest('hex')
      const dup = mediaByHash.get(imgHash)
      let mediaPart = ''
      let rid = ''
      let reusedNote = ''
      if (dup) {
        mediaPart = dup.part
        rid = dup.rid || ''
        reusedNote = `（内容哈希相同，复用 ${mediaPart}）`
        docxImageReused.push({ name: src, part: mediaPart, reason: '内容 SHA-256 与已有媒体部件相同，复用同一部件与关系 id' })
      } else {
        // media 部件序号：接着现有 word/media/ 往后编，绝不覆盖既有图片
        mediaSeqMax += 1
        mediaPart = `word/media/image${mediaSeqMax}.${imageExt(info)}`
        parts.set(mediaPart, img.buf)
        mediaByHash.set(imgHash, { part: mediaPart, rid: null })
      }
      docxImageMedia.push(mediaPart)
      if (!rid) {
        // 关系 id：接着现有 rId 往后编
        let maxRid = 0
        for (const r of existingRids) {
          const m = /^rId(\d+)$/.exec(r)
          if (m) maxRid = Math.max(maxRid, Number(m[1]))
        }
        rid = `rId${maxRid + 1}`
        existingRids.add(rid)          // 立刻记账：同一次 edit 里插多张图时 id 才会递增（否则撞号 → Word 判包损坏）
        imageRels.push({ id: rid, target: mediaPart.replace(/^word\//, ''), kind: imageExt(info) })
        const rec = mediaByHash.get(imgHash)
        if (rec) rec.rid = rid
      }
      // 尺寸（需求 4e）：省略 width → 原图像素 × 72/96；超 A4 可用宽则等比缩到可用宽；
      // 换算过程进 docxImageSizing，收尾时汇总成 notice（不静默改尺寸）。
      const size = imageSizeFor(info.width, info.height, op.width)
      const cx = Math.round(size.widthPt * 12700)
      const cy = Math.round(size.heightPt * 12700)
      const para = imageParagraphXml(rid, cx, cy, ++imageGraphSeq, String(op.alt ?? op.descr ?? ''))
      const placed = insertImageParagraph(docXml, para, { after: op.after, at: op.at })
      docXml = placed.xml
      docxImageSizing.push({ name: src, part: mediaPart, px: `${info.width}×${info.height}`, pt: `${size.widthPt}×${size.heightPt}`, rule: size.rule, capped: size.capped })
      summary.push(`${op.op}: 插入 ${info.kind} ${info.width}×${info.height} → ${size.widthPt}×${size.heightPt}pt`
        + `（${size.rule}${size.capped ? '，超可用宽已等比缩小' : ''}），${mediaPart}${reusedNote}`
        + (op.after && !placed.inserted ? '；⚠ 锚点文本未命中，已退化为插到文末' : ''))
      if (op.after && !placed.inserted) imageAnchorMiss = String(op.after)
    } else {
      throw new Error(`docx 不支持操作 "${op.op}"（可用: replace_text / append_markdown / set_meta / insert_image / append_image）`)
    }
  }
  parts.set('word/document.xml', utf8(docXml))
  // [Content_Types].xml：新插入的图片扩展名要有 Default，否则包不算"类型完备"
  if (imageRels.length) {
    const ct = parts.get(ctKey)?.toString('utf8') || ''
    let next = ct
    for (const kind of new Set(imageRels.map(r => r.kind))) next = ensureContentTypeDefault(next, kind, imageMime(kind))
    if (next !== ct) parts.set(ctKey, utf8(next))
  }
  if (hyperlinks.length || imageRels.length) {
    const relsText = zipObj.getText('word/_rels/document.xml.rels')
      || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    const relDoc = parseXML(relsText)
    let relBody = relDoc.children.find(c => typeof c !== 'string' && c.name === 'Relationships')
    if (!relBody) {
      relBody = { name: 'Relationships', attrs: { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }, children: [] }
      relDoc.children.push(relBody)
    }
    for (const h of hyperlinks) {
      relBody.children.push({
        name: 'Relationship',
        attrs: {
          Id: h.id,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
          Target: h.target,
          TargetMode: 'External',
        },
        children: [],
      })
    }
    for (const r of imageRels) {
      relBody.children.push({
        name: 'Relationship',
        attrs: {
          Id: r.id,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          Target: r.target,
        },
        children: [],
      })
    }
    parts.set('word/_rels/document.xml.rels', utf8(serializeXML(relDoc)))
  }
  const out = { bytes: mapToZip(parts, ['[Content_Types].xml', '_rels/.rels']), summary }
  const notices = []
  if (imageAnchorMiss) notices.push(`insert_image 的 after 锚点"${imageAnchorMiss}"未命中，图片已退化为插到文末`)
  if (docxImageSizing.length) {
    notices.push('图片尺寸换算：' + docxImageSizing.map(s =>
      `原始 ${s.px} → ${s.pt}pt（${s.rule}${s.capped ? '，超可用宽已等比缩小' : ''}）`).join('；'))
  }
  if (docxImageReused.length) {
    out.stats = { ...(out.stats || {}), imageReused: docxImageReused }
    notices.push(`${docxImageReused.length} 张图片内容与已有媒体部件相同，已复用（不重复写部件）`)
  }
  if (docxImageSkipped.length) {
    out.stats = { ...(out.stats || {}), imagesSkipped: docxImageSkipped }
    notices.push(`${docxImageSkipped.length} 张图片未能内嵌（已退化为字面文本，不静默丢图）：`
      + docxImageSkipped.map(s => `${s.name || '(无路径)'}（${s.reason}）`).join('；'))
  }
  if (notices.length) out.notice = notices.join('；')
  const mediaList = [...new Set(docxImageMedia)]
  if (mediaList.length) out.stats = { ...(out.stats || {}), imageMedia: mediaList, imageSizing: docxImageSizing }
  return out
}

async function editOdf(file, ops, exec) {
  const zipObj = openZip(await readBuffer(file))
  const parts = zipToMap(zipObj)
  let content = zipObj.getText('content.xml')
  if (content === undefined) throw new Error('odf: 缺少 content.xml')
  const summary = []
  for (const op of ops) {
    throwIfAborted(exec)
    if (op.op === 'replace_text') {
      const find = String(op.find ?? '')
      if (!find && !op.regex) throw new Error('replace_text 需要 find')
      const r = replaceTextInOdfContent(content, find, String(op.replace ?? ''), { regex: !!op.regex })
      if (r.xml) content = r.xml
      summary.push(`replace_text: 替换 ${r.count} 处`)
    } else if (op.op === 'append_markdown') {
      const md = String(op.markdown ?? op.text ?? '')
      if (!md.trim()) throw new Error('append_markdown 需要 markdown 内容')
      const doc = markdownToDocument(md)
      const xml = odfBlocksXml(doc.blocks)
      const idx = content.lastIndexOf('</office:text>')
      if (idx < 0) throw new Error('odf: 非文本文档，append_markdown 仅支持 .odt')
      content = content.slice(0, idx) + xml + content.slice(idx)
      summary.push(`append_markdown: 追加 ${doc.blocks.length} 个内容块`)
    } else {
      throw new Error(`odt/ods/odp 不支持操作 "${op.op}"（可用: replace_text / append_markdown(.odt)）`)
    }
  }
  parts.set('content.xml', utf8(content))
  return { bytes: mapToZip(parts, ['mimetype', 'META-INF/manifest.xml', 'content.xml']), summary }
}

function replaceTextInOdfContent(contentXml, find, replace, { regex = false } = {}) {
  const doc = parseXML(contentXml)
  let count = 0
  let re = null
  if (regex) {
    try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`正则表达式无效: ${e.message}`) }
  }
  const decode = s => s
  const plainText = node => node.children.map(c => (typeof c === 'string' ? decode(c) : c.cdata || '')).join('')
  for (const t of findAll(doc, ['text:p', 'text:h'])) {
    const text = plainText(t)
    if (!text) continue
    if (re) {
      const hits = text.match(re)
      if (hits) { count += hits.length; t.children = [encodeEntities(text.replace(re, () => replace))] }
    } else if (text.includes(find)) {
      count += text.split(find).length - 1
      t.children = [encodeEntities(text.split(find).join(replace))]
    }
  }
  return { xml: count ? serializeXML(doc) : null, count }
}

function odfBlocksXml(blocks) {
  const bytes = writeOdt({ blocks })
  const xml = openZip(bytes).getText('content.xml') || ''
  const m = /<office:text>([\s\S]*)<\/office:text>/.exec(xml)
  return m ? m[1] : ''
}

async function editXlsx(file, ops, exec) {
  const zipObj = openZip(await readBuffer(file))
  const parts = zipToMap(zipObj)
  let model = null
  const ensureModel = () => (model ??= readXlsx(zipObj))
  const summary = []
  const sheetXmlPath = i => `xl/worksheets/sheet${i + 1}.xml`
  const currentSheetXml = i => {
    const p = sheetXmlPath(i)
    return parts.has(p) ? Buffer.from(parts.get(p)).toString('utf8') : undefined
  }
  const sheetIndex = name => {
    const m = ensureModel()
    if (!name) return 0
    const i = m.sheets.findIndex(s => s.name === String(name))
    if (i < 0) throw new Error(`找不到工作表 "${name}"（现有: ${m.sheets.map(s => s.name).join(', ')}）`)
    return i
  }
  for (const op of ops) {
    throwIfAborted(exec)
    if (op.op === 'set_cell') {
      const i = sheetIndex(op.sheet)
      const ref = String(op.cell || '').toUpperCase().replace(/\$/g, '')
      if (!/^[A-Z]{1,3}\d{1,7}$/.test(ref)) throw new Error('set_cell 需要 cell 形如 B3')
      const xml = currentSheetXml(i)
      if (xml === undefined) throw new Error(`xlsx: 缺少 ${sheetXmlPath(i)}`)
      parts.set(sheetXmlPath(i), utf8(setCellInSheetXml(xml, ref, op.value)))
      summary.push(`set_cell ${ref} = ${JSON.stringify(op.value ?? '').slice(0, 40)}`)
    } else if (op.op === 'append_rows') {
      const i = sheetIndex(op.sheet)
      const rows = Array.isArray(op.rows) ? op.rows : []
      if (!rows.length) throw new Error('append_rows 需要 rows 数组')
      const xml = currentSheetXml(i)
      if (xml === undefined) throw new Error(`xlsx: 缺少 ${sheetXmlPath(i)}`)
      parts.set(sheetXmlPath(i), utf8(appendRowsToSheetXml(xml, rows)))
      summary.push(`append_rows: 追加 ${rows.length} 行`)
    } else if (op.op === 'replace_value' || op.op === 'replace_text') {
      const find = String(op.find ?? '')
      if (!find && !op.regex) throw new Error('replace_value 需要 find')
      const whole = op.op === 'replace_value' ? op.whole !== false : false
      const m = ensureModel()
      let total = 0
      const sstIndexes = new Set()
      for (let i = 0; i < m.sheets.length; i++) {
        const xml = currentSheetXml(i)
        if (xml === undefined) continue
        const res = replaceInSheetXml(xml, find, String(op.replace ?? ''), { regex: !!op.regex, whole, sstIndexes })
        total += res.count
        if (res.count) parts.set(sheetXmlPath(i), utf8(res.xml))
      }
      const sstText = zipObj.getText('xl/sharedStrings.xml')
      if (sstIndexes.size && sstText) {
        const res = replaceSharedStrings(sstText, sstIndexes, find, String(op.replace ?? ''), { regex: !!op.regex, whole })
        total += res.count
        parts.set('xl/sharedStrings.xml', utf8(res.xml))
      }
      summary.push(`${op.op}: 替换 ${total} 处`)
    } else if (op.op === 'add_sheet') {
      const m = ensureModel()
      const sheetName = String(op.name || `Sheet${m.sheets.length + 1}`).slice(0, 31)
      const idx = m.sheets.length
      const rows = (Array.isArray(op.rows) ? op.rows : []).map(r => (Array.isArray(r) ? r : [r]))
      parts.set(sheetXmlPath(idx), utf8(freshSheetXml(rows)))
      const wbXml = Buffer.from(parts.get('xl/workbook.xml')).toString('utf8')
      parts.set('xl/workbook.xml', utf8(wbXml.replace('</sheets>', `<sheet name="${encodeEntities(sheetName, true)}" sheetId="${idx + 1}" r:id="dshSheet${idx + 1}"/></sheets>`)))
      const relsXml = Buffer.from(parts.get('xl/_rels/workbook.xml.rels')).toString('utf8')
      parts.set('xl/_rels/workbook.xml.rels', utf8(relsXml.replace('</Relationships>', `<Relationship Id="dshSheet${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/></Relationships>`)))
      const ct = Buffer.from(parts.get('[Content_Types].xml')).toString('utf8')
      parts.set('[Content_Types].xml', utf8(ct.replace('</Types>', `<Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`)))
      summary.push(`add_sheet: 新增工作表 "${sheetName}"`)
    } else if (op.op === 'rename_sheet') {
      ensureModel()
      const wbXml = Buffer.from(parts.get('xl/workbook.xml')).toString('utf8')
      const re = new RegExp(`(<sheet[^>]*\\sname=")${escapeRe(String(op.name))}("[^>]*/?>)`)
      if (!re.test(wbXml)) throw new Error(`找不到工作表 "${op.name}"`)
      parts.set('xl/workbook.xml', utf8(wbXml.replace(re, `$1${encodeEntities(String(op.newName), true)}$2`)))
      summary.push(`rename_sheet: ${op.name} → ${op.newName}`)
    } else if (op.op === 'delete_sheet') {
      const m = ensureModel()
      if (m.sheets.length <= 1) throw new Error('不允许删除唯一的工作表')
      const i = m.sheets.findIndex(s => s.name === String(op.name))
      if (i < 0) throw new Error(`找不到工作表 "${op.name}"`)
      const wbDoc = parseXML(Buffer.from(parts.get('xl/workbook.xml')).toString('utf8'))
      const relsText = Buffer.from(parts.get('xl/_rels/workbook.xml.rels')).toString('utf8')
      const rels = new Map(findAll(parseXML(relsText), 'Relationship').map(r => [r.attrs.Id, r.attrs.Target]))
      const sheetNodes = findAll(wbDoc, 'sheet')
      const node = sheetNodes[i]
      const rid = node?.attrs['r:id'] || node?.attrs.Id
      const target = String(rels.get(rid) || '').replace(/^\.\.\//, '').replace(/^\//, '')
      const partPath = target.startsWith('xl/') ? target : `xl/${target}`
      const sheetsNode = findAll(wbDoc, 'sheets')[0]
      if (sheetsNode) sheetsNode.children = sheetsNode.children.filter(c => c !== node)
      const root = wbDoc.children.find(c => typeof c !== 'string')
      parts.set('xl/workbook.xml', utf8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serializeXML(root)}`))
      parts.set('xl/_rels/workbook.xml.rels', utf8(relsText.replace(new RegExp(`<Relationship[^>]*Id="${escapeRe(rid)}"[^>]*/>`), '')))
      const ct = Buffer.from(parts.get('[Content_Types].xml')).toString('utf8')
      parts.set('[Content_Types].xml', utf8(ct.replace(new RegExp(`<Override PartName="/${escapeRe(partPath)}"[^>]*/>`), '')))
      parts.delete(partPath)
      summary.push(`delete_sheet: 删除 "${op.name}"`)
    } else {
      throw new Error(`xlsx 不支持操作 "${op.op}"（可用: set_cell / append_rows / replace_value / add_sheet / rename_sheet / delete_sheet）`)
    }
  }
  return { bytes: mapToZip(parts, ['[Content_Types].xml', '_rels/.rels']), summary }
}

function escapeRe(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function cellToXml(ref, value) {
  const v = value && typeof value === 'object' ? value : { v: value }
  if (v.f) return `<c r="${ref}"><f>${encodeEntities(String(v.f).replace(/^=/, ''))}</f></c>`
  const raw = v.v ?? ''
  if (typeof raw === 'number' || (typeof raw === 'string' && raw.trim() !== '' && /^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(raw.trim()) && !/^0\d/.test(raw.trim()))) {
    return `<c r="${ref}"><v>${Number(raw)}</v></c>`
  }
  if (raw === true || raw === false || raw === 'TRUE' || raw === 'FALSE') {
    return `<c r="${ref}" t="b"><v>${raw === true || raw === 'TRUE' ? 1 : 0}</v></c>`
  }
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${encodeEntities(String(raw))}</t></is></c>`
}

function parseCellRefLite(ref) {
  const m = /^([A-Z]{1,3})(\d{1,7})$/.exec(ref)
  if (!m) return null
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { col: col - 1, row: Number(m[2]) - 1 }
}

function setCellInSheetXml(sheetXml, ref, value) {
  const p = parseCellRefLite(ref)
  if (!p) throw new Error(`非法单元格引用 "${ref}"`)
  const rowRe = new RegExp(`<row r="${p.row + 1}"([^>]*)>([\\s\\S]*?)</row>`)
  const cellXml = cellToXml(ref, value)
  const selfClosing = new RegExp(`<c r="${escapeRe(ref)}"(\\s[^>]*)?/>`)
  const paired = new RegExp(`<c r="${escapeRe(ref)}"(\\s[^>]*)?>[\\s\\S]*?</c>`)
  if (rowRe.test(sheetXml)) {
    const m = rowRe.exec(sheetXml)
    let body = m[2]
    if (paired.test(body)) body = body.replace(paired, cellXml)
    else if (selfClosing.test(body)) body = body.replace(selfClosing, cellXml)
    else body = insertCellOrdered(body, p.col, cellXml)
    return sheetXml.replace(rowRe, `<row r="${p.row + 1}"$1>${body}</row>`)
  }
  const newRow = `<row r="${p.row + 1}">${cellXml}</row>`
  if (!/<\/sheetData>/.test(sheetXml)) throw new Error('xlsx: 工作表缺少 sheetData')
  const rows = [...sheetXml.matchAll(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g)]
  for (const m of rows) {
    if (Number(m[1]) > p.row + 1) return sheetXml.slice(0, m.index) + newRow + sheetXml.slice(m.index)
  }
  return sheetXml.replace('</sheetData>', newRow + '</sheetData>')
}

function insertCellOrdered(rowBody, col, cellXml) {
  const re = /<c r="([A-Z]{1,3})\d+"/g
  let m
  let insertAt = rowBody.length
  while ((m = re.exec(rowBody))) {
    const c = parseCellRefLite(`${m[1]}1`)
    if (c && c.col > col) { insertAt = m.index; break }
  }
  return rowBody.slice(0, insertAt) + cellXml + rowBody.slice(insertAt)
}

function appendRowsToSheetXml(sheetXml, rows) {
  let maxRow = 0
  for (const m of sheetXml.matchAll(/<row r="(\d+)"/g)) maxRow = Math.max(maxRow, Number(m[1]))
  let add = ''
  rows.forEach((r, ri) => {
    const rowNum = maxRow + ri + 1
    const cells = (Array.isArray(r) ? r : [r]).map((c, ci) => cellToXml(colName(ci) + rowNum, c)).join('')
    add += `<row r="${rowNum}">${cells}</row>`
  })
  if (!add) return sheetXml
  if (!/<sheetData/.test(sheetXml)) throw new Error('xlsx: 工作表缺少 sheetData')
  return sheetXml.replace('</sheetData>', add + '</sheetData>')
}

function freshSheetXml(rows) {
  let body = ''
  rows.forEach((r, ri) => {
    const cells = (Array.isArray(r) ? r : [r]).map((c, ci) => cellToXml(colName(ci) + (ri + 1), c)).join('')
    body += `<row r="${ri + 1}">${cells}</row>`
  })
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData>${body}</sheetData></worksheet>`
}

/** Replace values inside one worksheet part; shared-string hits go to sstIndexes. */
function replaceInSheetXml(sheetXml, find, replace, { regex = false, whole = true, sstIndexes = new Set() } = {}) {
  const doc = parseXML(sheetXml)
  let count = 0
  let re = null
  if (regex) {
    try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`正则表达式无效: ${e.message}`) }
  }
  const decode = s => s
  for (const c of findAll(doc, 'c')) {
    const is = findAll(c, 'is')[0]
    const vNode = findAll(c, 'v')[0]
    if (c.attrs.t === 's') {
      if (!vNode) continue
      const idx = Number(vNode.children.map(x => (typeof x === 'string' ? decode(x) : '')).join(''))
      if (Number.isFinite(idx)) sstIndexes.add(idx)
      continue
    }
    let text = ''
    if (c.attrs.t === 'inlineStr' && is) {
      text = findAll(is, 't').map(t => t.children.map(x => (typeof x === 'string' ? decode(x) : '')).join('')).join('')
    } else if (vNode) {
      text = vNode.children.map(x => (typeof x === 'string' ? decode(x) : '')).join('')
    } else continue
    const res = applyReplacement(text, find, replace, { regex, whole }, re)
    if (!res) continue
    count += res.count
    if (is) {
      is.children = [{ name: 't', attrs: { 'xml:space': 'preserve' }, children: [encodeEntities(res.text)] }]
    } else if (vNode) {
      c.attrs.t = 'inlineStr'
      const at = c.children.indexOf(vNode)
      c.children.splice(at, 1, { name: 'is', attrs: {}, children: [{ name: 't', attrs: { 'xml:space': 'preserve' }, children: [encodeEntities(res.text)] }] })
    }
  }
  return { xml: serializeXML(doc), count }
}

function applyReplacement(text, find, replace, { regex, whole }, re) {
  if (regex) {
    const hits = text.match(re ?? new RegExp(find, 'g'))
    if (!hits) return null
    return { text: text.replace(re ?? new RegExp(find, 'g'), () => replace), count: hits.length }
  }
  if (whole) return text === find ? { text: replace, count: 1 } : null
  if (!text.includes(find)) return null
  return { text: text.split(find).join(replace), count: text.split(find).length - 1 }
}

function replaceSharedStrings(sstXml, indexes, find, replace, { regex = false, whole = true } = {}) {
  const doc = parseXML(sstXml)
  const sis = findAll(doc, 'si')
  let count = 0
  const re = regex ? new RegExp(find, 'g') : null
  for (const idx of indexes) {
    const si = sis[idx]
    if (!si) continue
    const tNodes = findAll(si, 't')
    if (!tNodes.length) continue
    const decode = s => s
    const text = tNodes.map(t => t.children.map(x => (typeof x === 'string' ? decode(x) : '')).join('')).join('')
    const res = applyReplacement(text, find, replace, { regex, whole }, re)
    if (!res) continue
    count += res.count
    si.children = [{ name: 't', attrs: { 'xml:space': 'preserve' }, children: [encodeEntities(res.text)] }]
  }
  return { xml: serializeXML(doc), count }
}

async function editPptx(file, ops, exec) {
  const zipObj = openZip(await readBuffer(file))
  const parts = zipToMap(zipObj)
  let model = null
  const ensureModel = () => (model ??= readPptx(zipObj))
  const summary = []
  for (const op of ops) {
    throwIfAborted(exec)
    if (op.op === 'replace_text') {
      const find = String(op.find ?? '')
      if (!find && !op.regex) throw new Error('replace_text 需要 find')
      let total = 0
      for (const partName of zipObj.names) {
        if (!/^ppt\/(slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/.test(partName)) continue
        const xml = zipObj.getText(partName)
        const r = replaceTextInPptxPart(xml, find, String(op.replace ?? ''), { regex: !!op.regex })
        if (r.count) { parts.set(partName, utf8(r.xml)); total += r.count }
      }
      summary.push(`replace_text: 替换 ${total} 处`)
    } else if (op.op === 'add_slide') {
      const m = ensureModel()
      const spec = buildSlideSpec(op)
      const num = m.slides.length + 1
      parts.set(`ppt/slides/slide${num}.xml`, utf8(serializeSlideXmlForEdit(spec)))
      parts.set(`ppt/slides/_rels/slide${num}.xml.rels`, utf8(slideRelsForEdit(num, spec)))
      const pres = Buffer.from(parts.get('ppt/presentation.xml')).toString('utf8')
      const ids = [...pres.matchAll(/<p:sldId id="(\d+)"/g)].map(x => Number(x[1]))
      const newId = Math.max(255, ...ids) + 1
      const rid = `dshSlide${num}`
      parts.set('ppt/presentation.xml', utf8(pres.replace('</p:sldIdLst>', `<p:sldId id="${newId}" r:id="${rid}"/></p:sldIdLst>`)))
      const presRels = Buffer.from(parts.get('ppt/_rels/presentation.xml.rels')).toString('utf8')
      parts.set('ppt/_rels/presentation.xml.rels', utf8(presRels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${num}.xml"/></Relationships>`)))
      addContentTypeOverride(parts, `/ppt/slides/slide${num}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
      if (spec.notes) {
        parts.set(`ppt/notesSlides/notesSlide${num}.xml`, utf8(notesXmlForEdit(spec, num)))
        parts.set(`ppt/notesSlides/_rels/notesSlide${num}.xml.rels`, utf8(notesRelsForEdit(num)))
        addContentTypeOverride(parts, `/ppt/notesSlides/notesSlide${num}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
      }
      summary.push(`add_slide: 新增第 ${num} 页${spec.title ? `（${spec.title}）` : ''}`)
    } else if (op.op === 'update_slide') {
      const m = ensureModel()
      const i = Number(op.index)
      if (!Number.isInteger(i) || i < 1 || i > m.slides.length) throw new Error(`update_slide 的 index 超出范围 1..${m.slides.length}`)
      const old = m.slides[i - 1]
      const spec = {
        layout: op.layout || old.layout,
        title: op.title !== undefined ? String(op.title) : old.title,
        subtitle: op.subtitle !== undefined ? String(op.subtitle) : old.subtitle,
        bullets: op.bullets !== undefined ? normalizeBullets(op.bullets) : old.bullets,
        table: op.table !== undefined ? op.table : old.table,
        notes: op.notes !== undefined ? String(op.notes) : old.notes,
      }
      parts.set(`ppt/slides/slide${i}.xml`, utf8(serializeSlideXmlForEdit(spec)))
      summary.push(`update_slide: 重写第 ${i} 页（应用新的标题/要点/备注）`)
    } else if (op.op === 'delete_slide') {
      const m = ensureModel()
      if (m.slides.length <= 1) throw new Error('不允许删除唯一的幻灯片')
      const i = Number(op.index)
      if (!Number.isInteger(i) || i < 1 || i > m.slides.length) throw new Error(`delete_slide 的 index 超出范围 1..${m.slides.length}`)
      const pres = Buffer.from(parts.get('ppt/presentation.xml')).toString('utf8')
      const presRels = Buffer.from(parts.get('ppt/_rels/presentation.xml.rels')).toString('utf8')
      const rels = new Map(findAll(parseXML(presRels), 'Relationship').map(r => [r.attrs.Id, r.attrs.Target]))
      const sldIds = findAll(parseXML(pres), 'p:sldId')
      const rid = sldIds[i - 1]?.attrs['r:id']
      const target = String(rels.get(rid) || '').replace(/^\.\.\//, '').replace(/^\//, '')
      const partPath = target.startsWith('ppt/') ? target : `ppt/${target}`
      parts.set('ppt/presentation.xml', utf8(pres.replace(new RegExp(`<p:sldId[^>]*r:id="${escapeRe(rid)}"[^>]*/>`), '')))
      parts.set('ppt/_rels/presentation.xml.rels', utf8(presRels.replace(new RegExp(`<Relationship[^>]*Id="${escapeRe(rid)}"[^>]*/>`), '')))
      removeContentTypeOverride(parts, `/${partPath}`)
      parts.delete(partPath)
      parts.delete(partPath.replace(/\/([^/]+)$/, '/_rels/$1.rels'))
      summary.push(`delete_slide: 删除第 ${i} 页`)
    } else {
      throw new Error(`pptx 不支持操作 "${op.op}"（可用: replace_text / add_slide / update_slide / delete_slide）`)
    }
  }
  return { bytes: mapToZip(parts, ['[Content_Types].xml', '_rels/.rels']), summary }
}

function buildSlideSpec(op) {
  if (op.slide && typeof op.slide === 'object') {
    return {
      layout: op.slide.layout || (op.slide.subtitle !== undefined ? 'title' : 'content'),
      title: op.slide.title !== undefined ? String(op.slide.title) : '',
      subtitle: op.slide.subtitle !== undefined ? String(op.slide.subtitle) : undefined,
      bullets: normalizeBullets(op.slide.bullets),
      table: op.slide.table,
      notes: op.slide.notes !== undefined ? String(op.slide.notes) : undefined,
    }
  }
  return {
    layout: op.subtitle !== undefined ? 'title' : 'content',
    title: op.title !== undefined ? String(op.title) : '',
    subtitle: op.subtitle !== undefined ? String(op.subtitle) : undefined,
    bullets: normalizeBullets(op.bullets),
    table: op.table,
    notes: op.notes !== undefined ? String(op.notes) : undefined,
  }
}

function normalizeBullets(bullets) {
  if (!Array.isArray(bullets)) return []
  return bullets.map(b => (typeof b === 'string' ? { text: b, level: 0 } : { level: 0, ...b, text: String(b?.text ?? '') }))
}

function addContentTypeOverride(parts, partName, contentType) {
  const ct = Buffer.from(parts.get('[Content_Types].xml')).toString('utf8')
  if (ct.includes(`PartName="${partName}"`)) return
  parts.set('[Content_Types].xml', utf8(ct.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)))
}

function removeContentTypeOverride(parts, partName) {
  const ct = Buffer.from(parts.get('[Content_Types].xml')).toString('utf8')
  parts.set('[Content_Types].xml', utf8(ct.replace(new RegExp(`<Override PartName="${escapeRe(partName)}"[^>]*/>`), '')))
}

async function editTextFile(file, ops, exec) {
  const ext = extOf(file)
  const buf = await readBuffer(file)
  // 需求 5：编辑前记住源文件有没有 BOM —— 旧版 decode→写回会把 BOM 静默剥掉，
  // "编辑一次、Excel 再打开就乱码"。有则原样写回（csv/tsv/md/txt 全适用）。
  const hadBom = buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf
  let text = decodeTextBytes(buf)
  const summary = []
  if (ext === 'csv' || ext === 'tsv') {
    const delim = ext === 'tsv' ? '\t' : ','
    for (const op of ops) {
      throwIfAborted(exec)
      const rows = parseDelimited(text, delim)
      if (op.op === 'replace_value' || op.op === 'replace_text') {
        const find = String(op.find ?? '')
        if (!find) throw new Error(`${op.op} 需要 find`)
        const replace = String(op.replace ?? '')
        const whole = op.op === 'replace_value'
        let count = 0
        for (const row of rows) {
          for (let i = 0; i < row.length; i++) {
            if (whole ? row[i] === find : row[i].includes(find)) {
              count += whole ? 1 : row[i].split(find).length - 1
              row[i] = whole ? replace : row[i].split(find).join(replace)
            }
          }
        }
        text = renderDelimited(rows, delim)
        summary.push(`${op.op}: 修改 ${count} 个单元格`)
      } else if (op.op === 'append_rows') {
        const add = Array.isArray(op.rows) ? op.rows : []
        if (!add.length) throw new Error('append_rows 需要 rows 数组')
        rows.push(...add.map(r => (Array.isArray(r) ? r.map(String) : [String(r)])))
        text = renderDelimited(rows, delim)
        summary.push(`append_rows: 追加 ${add.length} 行`)
      } else {
        throw new Error(`csv/tsv 不支持操作 "${op.op}"（可用: replace_value / replace_text / append_rows）`)
      }
    }
  } else {
    for (const op of ops) {
      throwIfAborted(exec)
      if (op.op === 'replace_text') {
        const find = String(op.find ?? '')
        if (!find && !op.regex) throw new Error('replace_text 需要 find')
        if (op.regex) {
          let re
          try { re = new RegExp(find, 'g') } catch (e) { throw new Error(`正则表达式无效: ${e.message}`) }
          const hits = text.match(re)
          text = text.replace(re, () => String(op.replace ?? ''))
          summary.push(`replace_text: 替换 ${hits ? hits.length : 0} 处`)
        } else {
          const count = text.split(find).length - 1
          text = text.split(find).join(String(op.replace ?? ''))
          summary.push(`replace_text: 替换 ${count} 处`)
        }
      } else if (op.op === 'append_text') {
        text = `${text.replace(/\s*$/, '')}\n${String(op.text ?? '')}\n`
        summary.push('append_text: 追加文本')
      } else if (op.op === 'prepend_text') {
        text = `${String(op.text ?? '')}\n${text}`
        summary.push('prepend_text: 前置文本')
      } else {
        throw new Error(`文本文件不支持操作 "${op.op}"（可用: replace_text / append_text / prepend_text）`)
      }
    }
  }
  return { bytes: Buffer.from((hadBom ? '\ufeff' : '') + text, 'utf8'), summary }
}

// ---------------------------------------------------------------------------
// scanned-PDF OCR: rasterize a page with Windows' built-in PDF renderer, read it
// with the bundled local engine (RapidOCR-json) and fall back to the vision
// bridge only for pages the local engine got wrong or judged unreliable
// ---------------------------------------------------------------------------

const OCR_MARK = '<!-- dsh-office OCR cache -->'

/**
 * **文字层解析器版本戳**（sidecar manifest 的 `parser:` 段）—— 任务 D 的缓存迁移凭据。
 *
 * 为什么必须有：`.ocr.md` 里的成果是按**页码**记的，而页码有两个互不相干的来源 ——
 * 文字层走插件自己的 `pages()`（PDF 页面树的枚举顺序），OCR 渲染走 `pdf-render.ps1`
 * 的 WinRT `GetPage(n-1)`（文件里声明的物理页序）。两者一旦不一致，`ocrMap.get(s.page)`
 * 就会把**另一张纸**的识别结果贴到这一页上，而且是静默的。
 *
 * 实测（本轮）：修 `expandObjStms` 的 off-by-`First` 之前，病灶样本①的 `pages()` 顺序
 * 从文档声明的 `/Kids` 顺序退化成"对象号升序"（封面 `436` 从第 1 位跑到第 34 位），
 * 于是第 1–33 页的 OCR 成果整体错一位。**旧 sidecar 一律不可信**。
 *
 * 因此：`parser` 缺失或 ≠ 当前值 → 该 sidecar 视为**未覆盖**（丢弃页缓存、要求重识别），
 * 并在 stats/notice 里明说原因 —— 绝不静默复用错页文本。任何改变"页码 → 页对象"
 * 映射的改动都必须让这个数字 +1。
 */
export const PDF_PARSER_VERSION = 2

/** 页码列表 → 紧凑区间串（1、2、3、7、9 → `1-3,7,9`），manifest 与脚注共用。 */
function pageRanges(nums) {
  const ps = [...new Set(nums)].map(Number).filter(Number.isFinite).sort((a, b) => a - b)
  const out = []
  for (let i = 0; i < ps.length;) {
    let j = i
    while (j + 1 < ps.length && ps[j + 1] === ps[j] + 1) j++
    out.push(j > i ? `${ps[i]}-${ps[j]}` : String(ps[i]))
    i = j + 1
  }
  return out.join(',')
}

/** `1-3,7` → [1,2,3,7]（只用于回读 manifest）。 */
function expandRanges(spec) {
  const out = []
  for (const chunk of String(spec).split(/[,、\s]+/)) {
    if (!chunk) continue
    const m = /^(\d+)\s*[-~]\s*(\d+)$/.exec(chunk)
    if (m) { for (let i = Number(m[1]); i <= Number(m[2]); i++) out.push(i) } else if (/^\d+$/.test(chunk)) out.push(Number(chunk))
  }
  return out
}

/**
 * 两条"静默少做"脚注的**唯一文案源** —— 正常路径（readPdf 的 notes 组装）与
 * 降级路径（finishRead 的 sidecar 兜底，R9-3）共用同一份模板，杜绝两处文案漂移。
 *
 * 顺序固定：capped（单次上限截断）先、preview（未指定 pages 的前 N 页预览）后
 * —— 第八轮 P3 对账 #1 定的顺序。
 * `skipped` 形状归一：内部记账是**页码数组**（ocrPagesCapped / ocrPreviewLimited），
 * 而外显的 `stats.ocrPreview.skipped` 已是**区间字符串**（见 stats 组装处）——两种都要吃。
 */
function ocrCapNotes(capped, preview, file) {
  const ranges = v => Array.isArray(v) ? pageRanges(v) : String(v ?? '')
  const notes = []
  if (capped) {
    const skipped = ranges(capped.skipped)
    notes.push(`单次读取最多识别 ${capped.limit} 页：本次要求 ${capped.requested} 页，`
      + `第 ${skipped} 页未做 OCR（仍用原文本层）。`
      + `续读：office_read path="${file}" ocr="always" ocrEngine="local" pages="${skipped}"`)
  }
  if (preview) {
    const skipped = ranges(preview.skipped)
    notes.push(`未指定 pages 按预览只 OCR 前 ${preview.preview} 页（待识别共 ${preview.total} 页，`
      + `第 ${skipped} 页未做，仍用原文本层）。`
      + `续读：office_read path="${file}" ocr="always" ocrEngine="local" pages="${skipped}"`)
  }
  return notes
}

/**
 * Path of the sidecar OCR cache next to the PDF (falls back to temp).
 *
 * 任务一：集中缓存目录与临时回退目录带**路径指纹后缀**。旧命名 `<basename>.ocr.md` 会让
 * `D:\a\第1章.pdf` 与 `D:\b\第1章.pdf` 共用同一份识别成果（互相命中 / 互相覆盖）。
 * 同目录 sidecar（`<file>.ocr.md`）的路径本身已唯一，命名**逐字保持不变**（旧调用方零影响）。
 */
function ocrCachePath(file) {
  const key = pathKeyOf(file)
  const base = basename(file).replace(/\.pdf$/i, '')
  const st = cacheDirState()
  if (st.dir) return join(st.dir, `${base}-${key}.ocr.md`)
  const side = file.replace(/\.pdf$/i, '') + '.ocr.md'
  try {
    const dir = dirname(side)
    if (existsSync(dir)) return side
  } catch { /* fall through */ }
  return join(tmpdir(), 'dsh-office-ocr', `${base.replace(/[^\w.-]/g, '_')}-${key}.ocr.md`)
}

/**
 * 旧命名规则（无路径指纹）—— 只用来**发现并明确作废**历史缓存：
 * 旧的集中目录缓存没有身份字段，页文本一律不回收，但要让用户看到"它为什么没了"。
 * 同目录 sidecar 的新旧命名是同一个路径，调用方用 `===` 判重后跳过。
 */
function legacyOcrCachePath(file) {
  const base = basename(file).replace(/\.pdf$/i, '')
  const st = cacheDirState()
  if (st.dir) return join(st.dir, `${base}.ocr.md`)
  const side = file.replace(/\.pdf$/i, '') + '.ocr.md'
  try {
    if (existsSync(dirname(side))) return side
  } catch { /* fall through */ }
  return join(tmpdir(), 'dsh-office-ocr', `${base.replace(/[^\w.-]/g, '_')}.ocr.md`)
}

/**
 * Read the sidecar: page bodies plus the optional header manifest.
 * `pages` is always authoritative (derived from the `## 第 N 页` sections, so it
 * cannot disagree with the text); `total` / `src` only exist in sidecars written
 * by a version that records them — an old sidecar simply reads as unknown.
 * `retry` 是历史换倍率记录（`page → scale`），读回来是为了**跨批累积**而不是每批重写时丢掉。
 *
 * 任务一：**缓存身份必须先核对再复用**。`identity` 由调用方传入（来自 loadModel 的
 * `extra.identity`，零额外读盘）；缺省时本函数同步算一次（memo 命中则免费）。
 * 作废原因写进 `staleReason`，供 staleCacheNote 说清"为什么没了"：
 *   `parser` / `parser-missing` / `identity-missing` / `path-mismatch` / `content-mismatch` / `unverifiable`
 */
function readOcrCache(file, identity) {
  const map = new Map()
  const src = new Map()
  const retry = new Map()
  let total
  let parser
  const fresh = { pages: map, src, total, retry, parser, stale: false, staleReason: null, path: ocrCachePath(file), identity: null }
  const p = fresh.path
  try {
    let readPath = p
    if (!existsSync(readPath)) {
      // 旧命名（无路径指纹）的历史缓存：读进来只为**明确作废并说清原因**，页文本一律不回收。
      const legacy = legacyOcrCachePath(file)
      if (legacy === p || !existsSync(legacy)) return fresh
      readPath = legacy
    }
    const text = readFileSync(readPath, 'utf8')
    if (!text.startsWith(OCR_MARK)) return { ...fresh, path: readPath }
    // 只在页正文之前的头部找 manifest，免得正文里的 "total:"/"src:" 被误读
    const head = text.split(/\r?\n## 第 /)[0].replace(/-->/g, '')
    const pm = /\bparser:\s*(\d+)/.exec(head)
    parser = pm ? Number(pm[1]) : undefined
    // ---- 身份核对（任务一）----
    // 与 `parser` 规则是"与"的关系：任一不符 → 整份作废，绝不静默复用错页文本。
    const phm = /\bsrcpath:\s*([0-9a-f]{64})\b/i.exec(head)
    const shm = /\bsrcsha256:\s*([0-9a-f]{64})\b/i.exec(head)
    const id = identity !== undefined ? identity : sourceIdentityOf(file)
    let reason = null
    if (parser !== PDF_PARSER_VERSION) reason = parser === undefined ? 'parser-missing' : 'parser'
    else if (!phm || !shm) reason = 'identity-missing'
    else if (!id) reason = 'unverifiable'
    else if (phm[1].toLowerCase() !== id.pathHash) reason = 'path-mismatch'
    else if (shm[1].toLowerCase() !== id.contentHash) reason = 'content-mismatch'
    if (reason) {
      // 页号可能指的不是同一张纸、正文也可能是另一份源的 → 页缓存整份不回收。
      return { pages: map, src, total, retry, parser, stale: true, staleReason: reason, path: readPath, identity: id || null }
    }
    const tm = /\btotal:\s*(\d+)/.exec(head)
    if (tm) total = Number(tm[1])
    const sm = /\bsrc:\s*([^|\n]+)/.exec(head)
    if (sm) {
      // 引擎之间用 `;` 分隔，区间之间才用 `,`（否则 `rapidocr=1-3,7` 会被拆坏）
      for (const part of sm[1].split(';')) {
        const m = /^\s*(rapidocr|vision)\s*=\s*([\d\s\-~、,]+?)\s*$/.exec(part)
        if (!m) continue
        for (const n of expandRanges(m[2])) src.set(n, m[1])
      }
    }
    // **DSH 侧修正**：旧版这里用 `[^\n]+` 一路吃到行尾。manifest 是一个对象数组的
    // `key: value` 用 ` | ` 拼在同一行上 —— 于是第二轮新加的 `retry:` 段被当成 src 值的一部分，
    // 逐段正则不匹配 → **整行 src 解析不出来**。后果：下一次写 sidecar 时 `srcOf` 是空的，
    // 于是头部只剩本批页的 src（跨批合并的历史来源被静默丢弃），retry 记录同样被抹掉。
    // 收紧到 `[^|\n]+` 并把 retry 也读回来，两条账都能跨批累积。
    const rm = /\bretry:\s*([^|\n]+)/.exec(head)
    if (rm) {
      for (const part of rm[1].split(/[,\s;]+/)) {
        const m = /^(\d+)=(\d+(?:\.\d+)?)$/.exec(part)
        if (m) retry.set(Number(m[1]), Number(m[2]))
      }
    }
    // 逐行切段：注意不能用 `[\s\S]*?(?=^## …|$)` 配 /m —— 多行模式下 $ 会命中每个行尾，
    // 于是每页只能取回第一行（旧版缓存回读被静默截断，就是这个坑）。
    let page = 0
    let buf = null
    for (const line of text.split(/\r?\n/)) {
      const m = /^## 第 (\d+) 页/.exec(line)
      if (m) {
        if (page) map.set(page, buf.join('\n').trim())
        page = Number(m[1]); buf = []
        continue
      }
      if (page) buf.push(line)
    }
    if (page) map.set(page, buf.join('\n').trim())
    return { pages: map, src, total, retry, parser, stale: false, staleReason: null, path: readPath, identity: id || null }
  } catch { /* cache is best-effort */ }
  return fresh
}

/**
 * 旧解析器 sidecar 的提示文案 —— read / convert / sidecar 三处共用同一措辞。
 * 核心是**别让"缓存没了"变成一个没有解释的现象**：要说清是哪个文件、为什么作废、
 * 以及下一步怎么办。任务一：作废原因从"只有 parser 版本"扩到**源文件路径 / 内容身份**，
 * 每一类都要能被用户对上号（否则"缓存不命中了"照样是个没有解释的现象）。
 */
function staleCacheNote(info) {
  const parserTxt = info?.parser === undefined ? '未记录解析器版本' : `记录的是 parser: ${info.parser}`
  // 调用点有的传 `{ path, parser, reason }`，有的直接传 readOcrCache 的 store（字段是 `staleReason`）
  const reason = info?.reason ?? info?.staleReason
  const why =
    reason === 'parser' ? `${parserTxt}，当前 parser: ${PDF_PARSER_VERSION}`
      : reason === 'parser-missing' ? `未记录解析器版本（当前 parser: ${PDF_PARSER_VERSION}）`
        : reason === 'path-mismatch' ? '记录的 srcpath 与当前源文件路径不符（集中缓存目录里同名文件写下的缓存）'
          : reason === 'content-mismatch' ? '记录的 srcsha256 与当前源文件内容不符（同一路径的 PDF 已被替换或改动）'
            : reason === 'identity-missing' ? '未记录源文件身份（缺 srcpath / srcsha256 段）'
              : reason === 'unverifiable' ? '读不到源文件，无法核实缓存身份'
                : `${parserTxt}，当前 parser: ${PDF_PARSER_VERSION}`
  return `已作废旧 OCR 缓存（${info.path}：${why}）—— `
    + '页号与正文是按**另一份源**（或旧解析器口径）记的成果，整份不回收（绝不静默复用错页文本）；'
    + `请重新识别需要的页（office_read ocr="always"），旧文件可直接删除`
}

/**
 * Write the sidecar. The header carries a manifest so the cache states its own
 * coverage — `covered`/`total` tell a caller how much of the document is done
 * (the page cache is per-page, so the union can be expressed exactly), and
 * `src` records which engine produced each page, which is what lets a later
 * cache hit say whose result it is handing back. One line, constant size.
 *
 * 任务一：再补一行**身份**（`srcpath` / `srcsha256` / `srcsize`），让下一次读取能证明
 * "这些页确实是这份源文件的这份内容"；身份行独立成行，covered 行保持逐字不变。
 * `meta.identity` 由调用方传入（ocrPdfPages 从 readOcrCache 拿回，避免重复哈希）。
 */
function writeOcrCache(file, map, meta = {}) {
  const p = ocrCachePath(file)
  try {
    mkdirSync(dirname(p), { recursive: true })
    const pages = [...map.keys()].sort((a, b) => a - b)
    const body = pages.map(n => `## 第 ${n} 页（OCR）\n${map.get(n)}`).join('\n\n')
    const manifest = [`covered: ${pageRanges(pages) || '无'}`]
    if (meta.total) manifest.push(`total: ${meta.total}`)
    // 版本戳必须写：它是下一次读取判断"这些页号还指同一张纸吗"的唯一凭据。
    manifest.push(`parser: ${PDF_PARSER_VERSION}`)
    const bySrc = new Map()
    for (const n of pages) {
      const s = meta.src?.get?.(n)
      if (!s) continue
      if (!bySrc.has(s)) bySrc.set(s, [])
      bySrc.get(s).push(n)
    }
    if (bySrc.size) manifest.push(`src: ${[...bySrc].map(([s, ns]) => `${s}=${pageRanges(ns)}`).join(';')}`)
    // 换倍率才识别出来的页：src 仍是 rapidocr（引擎没变），倍率单独记一笔
    if (meta.retry?.size) {
      manifest.push(`retry: ${[...meta.retry].sort((a, b) => a[0] - b[0]).map(([n, s]) => `${n}=${s}`).join(',')}`)
    }
    // 身份行：**独立一行**（covered 行逐字不变，旧调用方 `split('\n')[1]` 的断言不受影响）。
    // 身份读不到（源文件已不在）就不写，绝不编造 —— 下一次读会按 `identity-missing` 明确作废，
    // 而不是把这份成果静默贴到别的源上。
    const id = meta.identity !== undefined ? meta.identity : sourceIdentityOf(file)
    const identityLine = id
      ? `<!-- srcpath: ${id.pathHash} | srcsha256: ${id.contentHash} | srcsize: ${id.size} -->\n`
      : ''
    // 任务二：原子写 —— 写一半的 `.ocr.md` 会被 readOcrCache 判"旧格式/身份缺失"而整份作废，
    // 等于把已经花掉的识别成果丢掉；这里保证"要么完整、要么没有"。
    writeFileAtomicSync(p, `${OCR_MARK}\n<!-- ${manifest.join(' | ')} -->\n${identityLine}`
      + `# OCR 文本（由 dsh-office 生成，可安全删除以重新识别）\n\n${body}\n`)
    return p
  } catch {
    return undefined
  }
}

/**
 * Temp dir holding this PDF's rendered PNGs (keyed by mtime → auto-invalidates).
 *
 * `tag` 区分同一页在不同渲染倍率下的产物（质量门重试用）：默认（原生分辨率）仍是
 * 老目录名 `<safe>-<mtime>`，逐字不变；带倍率时加 `-s<倍率>` 后缀 —— 必须分目录，
 * 否则 `renderedPng()` 会命中上一次渲染的 PNG，重试就成了空转。
 *
 * 任务一：目录名再嵌**路径指纹**（`<safe>-<key8>-<mtime>`）。旧命名只看 basename + mtime，
 * 两份不同目录下的同名 PDF 若 mtime 相同（复制文件很常见）就会共用页面 PNG ——
 * 那等于把**另一份文档的页面图**喂给 OCR，比缓存碰撞更危险。
 */
function renderDirFor(file, tag = '') {
  const safe = basename(file).replace(/[^\w.-]/g, '_')
  const suffix = tag ? `-s${String(tag).replace(/[^\w.]/g, '')}` : ''
  return join(tmpdir(), 'dsh-office-ocr', `${safe}-${pathKeyOf(file)}-${statSync(file).mtimeMs | 0}${suffix}`)
}

function renderedPng(outDir, pageNo) {
  const png = join(outDir, `page-${pageNo}.png`)
  return existsSync(png) && statSync(png).size > 1000 ? png : null
}

/**
 * `.ocr.md` / `.ocr.json` / `.read.md` 是花钱识别或兜底转存的成果，任何清扫都不许碰。
 *
 * R18 任务 D3 复核结论：`.ocr.json` **本仓没有任何写入点**（全仓 grep：只在三处"受保护"判断
 * 里出现）。**决定：保留**。理由是保留的成本为一次字符串后缀比较，而删掉有两个真实下行风险：
 *  ① 历史版本 / 姊妹实现（WorkBuddy 侧）可能已经把 `.ocr.json` 写进共享缓存目录与附件目录，
 *     清扫器一旦不认它，就会删掉用户的识别成果；
 *  ② 目录盘点（`office_read paths=[目录]`）会把它当成"用户文档"列进清单 —— 那本来就是旧小坑。
 * `test.mjs` 的 `P2-5 .ocr.md/.ocr.json 与其所在目录永不碰` 是这条结论的守门用例。
 */
const isProtectedCacheName = n => {
  const l = String(n).toLowerCase()
  return l.endsWith('.ocr.md') || l.endsWith('.ocr.json') || l.endsWith('.read.md')
}

const RENDER_SWEEP_MEMO = new Set()

/**
 * 渲染缓存清理：`renderDirFor` 按 mtime 键控，`%TEMP%\dsh-office-ocr` 会随文件每次改动无限累积。
 * 只动两类：① 同一 basename 的旧 mtime 目录（>7 天）；② `rapid-*` 引擎暂存（>1 天）。
 * 其余一律不碰 —— 尤其是 `.ocr.md` / `.ocr.json`（识别成果）、任何**内含**缓存的目录，
 * 以及当前正在用的目录。全程 try/catch：清理是尽力而为，绝不影响识别本身。
 */
export function sweepRenderCache(file, { now = Date.now() } = {}) {
  const report = { root: '', scanned: 0, removed: [], kept: [] }
  try {
    const root = join(tmpdir(), 'dsh-office-ocr')
    report.root = root
    if (!existsSync(root)) return report
    const safe = basename(file).replace(/[^\w.-]/g, '_')
    let current = ''
    try { current = renderDirFor(file) } catch { /* 文件已不在就只按名字前缀清 */ }
    // 同一 basename 的旧 mtime 目录：默认目录 `<safe>-<mtime>`，重试目录多一个 `-s<倍率>` 后缀。
    // 任务一给默认目录加了路径指纹（`<safe>-<key8>-<mtime>`）—— 这里两种都认，
    // 免得历史目录永远清不掉（渲染产物可重建，误清只是多渲染一次，绝不碰识别成果）。
    const sameFile = new RegExp(`^${safe.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(-[0-9a-f]{8})?-\\d+(-s[\\d.]+)?$`)
    for (const ent of readdirSync(root, { withFileTypes: true })) {
      report.scanned++
      const full = join(root, ent.name)
      if (full === current || isProtectedCacheName(ent.name)) { report.kept.push(full); continue }
      let st = null
      try { st = statSync(full) } catch { continue }
      const age = now - st.mtimeMs
      const staleDir = sameFile.test(ent.name) && age > 7 * 86400e3
      const staleRapid = ent.name.startsWith('rapid-') && age > 86400e3
      if (!staleDir && !staleRapid) { report.kept.push(full); continue }
      // 删之前再看一眼里面有没有花钱识别出来的缓存
      if (ent.isDirectory()) {
        let inner = []
        try { inner = readdirSync(full) } catch { inner = [] }
        if (inner.some(isProtectedCacheName)) { report.kept.push(full); continue }
      }
      try {
        rmSync(full, { recursive: true, force: true })
        report.removed.push(full)
      } catch { report.kept.push(full) }
    }
  } catch { /* 清理失败无所谓 */ }
  return report
}

/** 每进程每 file|mtime 只扫一次（每批 OCR 进渲染前都会走到这里）。 */
function sweepRenderCacheOnce(file) {
  let key = file
  try { key = `${file}|${statSync(file).mtimeMs | 0}` } catch { /* ignore */ }
  if (RENDER_SWEEP_MEMO.has(key)) return
  if (RENDER_SWEEP_MEMO.size > 256) RENDER_SWEEP_MEMO.clear()
  RENDER_SWEEP_MEMO.add(key)
  sweepRenderCache(file)
}

/**
 * Run pdf-render.ps1; returns the spawn result (the PNG on disk is the truth).
 * R18 起导出：测试里的 WinRT 判别实验必须与实现**同源**地调用渲染器
 * （受限沙箱下管道 stdio 会被拒成 EPERM，这里已经带"忽略输出重来一次"的兜底）。
 */
export function runRenderScript(file, outDir, extraArgs, scale) {
  const script = join(PLUGIN_DIR, 'pdf-render.ps1')
  if (!existsSync(script)) throw new Error(`缺少栅格化脚本 ${script}（重装插件可恢复）`)
  mkdirSync(outDir, { recursive: true })
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, file, outDir, ...extraArgs]
  const opts = { encoding: 'utf8' }
  // 重试倍率走**子进程 env**（不动 process.env）：pdf-render.ps1 只认环境变量，
  // 而引擎侧走显式参数；两边拿同一个值，且并发 read 之间不会互相污染。
  if (scale !== undefined) opts.env = { ...process.env, DSH_OFFICE_RENDER_SCALE: String(scale) }
  let res = spawnSync('powershell.exe', args, opts)
  // 受限沙箱里管道 stdio 会被拒（EPERM）→ 忽略输出重来一次，只认落地的 PNG
  if (res.error && /EPERM|EACCES/.test(String(res.error.code))) res = spawnSync('powershell.exe', args, { ...opts, stdio: 'ignore' })
  return res
}

function renderFailureDetail(res) {
  return String(res?.stderr || res?.error?.message || `exit=${res?.status}`).trim().split('\n')[0].slice(0, 200)
}

/**
 * 渲染失败时的**位置提示**（R18 任务 A 的连带诊断）：本机 WinRT 只读得到 `%TEMP%` 内的 PDF。
 * 只有当暂存件不在 `%TEMP%` 下时才追加这句话（正常路径一个字符都不变）——
 * 否则 `DSH_OFFICE_PDF_GATE_DIR` 指错地方时调用方只会看到一句无信息量的 `exit=1`。
 */
function renderLocationHint(stage) {
  try {
    const tmpRoot = resolvePath(tmpdir())
    const dir = resolvePath(dirname(stage))
    if (dir.toLowerCase().startsWith(tmpRoot.toLowerCase())) return ''
    return `；注意=渲染暂存目录 ${dir} 不在 %TEMP% 下，而本机 WinRT 只读得到 %TEMP% 内的 PDF`
      + `（把 DSH_OFFICE_PDF_GATE_DIR 指回 ${join(tmpRoot, 'dsh-office-pdfgate')} 或干脆不设它）`
  } catch { return '' }
}

/**
 * Rasterize one PDF page to PNG (cached in the temp render directory).
 * `tag` / `scale` 只在质量门重试时给（见 renderDirFor）。
 */
function renderPdfPage(file, pageNo, exec, tag = '', scale) {
  void exec
  const outDir = renderDirFor(file, tag)
  const hit = renderedPng(outDir, pageNo)
  if (hit) return hit
  const res = runRenderScript(file, outDir, [String(pageNo - 1), `page-${pageNo}.png`], scale)
  const png = renderedPng(outDir, pageNo)
  if (!png) throw new Error(`第 ${pageNo} 页栅格化失败：${renderFailureDetail(res)}`)
  return png
}

/**
 * Rasterize many pages with ONE document load. Loading a 30 MB scan costs seconds
 * per call, so per-page rendering used to dominate batch OCR.
 * Returns Map(page → png path); pages that failed to render are absent.
 */
function renderPdfPages(file, pageNumbers, exec, tag = '', scale) {
  void exec
  const outDir = renderDirFor(file, tag)
  const map = new Map()
  const want = []
  for (const p of pageNumbers) {
    const hit = renderedPng(outDir, p)
    if (hit) map.set(p, hit); else want.push(p)
  }
  if (want.length) {
    try { runRenderScript(file, outDir, [want.join(',')], scale) } catch { /* judged per page below */ }
    for (const p of want) {
      const png = renderedPng(outDir, p)
      if (png) map.set(p, png)
    }
  }
  return map
}

/** Cut a rendered page into one horizontal band (cached next to the page PNG). */
function bandPng(pagePng, index, count) {
  if (count <= 1) return pagePng
  const out = pagePng.replace(/\.png$/i, `-b${index + 1}of${count}.png`)
  if (existsSync(out) && statSync(out).size > 500) return out
  // 任务二：切片缓存也走原子写 —— 半截 PNG 只要 >500 字节就会被上面的判据当成"已渲染好"复用，
  // 视觉模型于是读到一张坏图（失败现象离根因很远）。
  writeFileAtomicSync(out, cropPngBand(readFileSync(pagePng), index, count))
  return out
}

function visionText(vision, args, value) {
  try {
    const blocks = vision.output?.render?.(args, value)
    if (Array.isArray(blocks)) {
      const text = blocks.filter(b => b && b.type === 'text').map(b => b.text).join('\n').trim()
      if (text) return text
    }
  } catch { /* fall through to raw value */ }
  if (typeof value === 'string') return value
  // 视觉桥常把证据做成结构化对象（modlens 的 ocr.full_text 等）。先把正文字段挑出来，
  // 否则一整坨 JSON 会被写进 .ocr.md 缓存，再也洗不干净。
  if (value && typeof value === 'object') {
    for (const cand of [value?.ocr?.full_text, value.full_text, value.text, value.content, value.result]) {
      if (typeof cand === 'string' && cand.trim()) return cand
    }
  }
  try { return JSON.stringify(value) } catch { return '' }
}

const VISION_LEADIN = [
  /^[\s>#*-]*(?:transcription|transcript|full\s*text|ocr\s*output|text\s*below|原文|正文|转录(?:结果)?|抄写(?:结果|内容)?)\s*[:：]\s*$/i,
  /^[\s>#*-]*(?:here(?:'|’)?s|here\s+is|the\s+following|below\s+is)[^\n]{0,40}[:：]\s*$/i,
  // 只在明确指向"图/页/识别结果"时才当客套话；"以下为具体安排："这种正文引导句必须保住
  /^[\s>#*-]*(?:以下|下列|下面)(?:是|为)?(?:这张|这份|该|此)?(?:图片|图像|页面|文档|图中|扫描件)(?:的)?[^\n]{0,24}[:：]\s*$/,
]
const VISION_DESC = /^(?:该|此|这|本|上)?[\s]?(?:图片|图像|扫描件|照片|截图|页面|页|文档|图)(?:的内容|内容|的局部|局部)?(?:为|是|显示|展示|呈现|描绘|中|看起来|表明)/
/** 视觉模型的自我提醒（"Uncertain: 底部被截断"），永远不是页面原文。 */
const VISION_UNCERTAIN = /^[\s>#*-]*(?:uncertain(?:ty)?|不确定(?:部分|区域)?|无法确认)\s*[:：]/i
/**
 * 图说句：整行在说"这是一张/该页图片如何如何"，且以句号/冒号收尾。
 * 只在后面跟着标记行时才用来判噪声，所以不会误伤正文（正文里不会出现 `Transcription:`）。
 */
const isVisionCaption = l => VISION_DESC.test(String(l).trim())
  || (/[。：:]\s*$/.test(String(l).trim()) && /(?:图片|图像|截图|照片|扫描件|页面|文档|表格|图)/.test(l))

/**
 * Vision models like to wrap a transcription in a lead-in ("该图片为…", "Transcription:")
 * or markdown fences. Since page text is written to the sidecar cache verbatim, that
 * noise would be reused forever — strip it at the source. Deliberately conservative:
 * without a clear marker line the text comes back untouched, so real content is never
 * mistaken for chatter.
 *
 * 多 tile 拼接（视觉模型输出上限 → 一页切 2/4/8 片，见 ocrPageText）会把模型对
 * **每一片**的图说都拼进同一页，所以标记行不止出现在开头：凡"其后（跳过空行）就是
 * 标记行"的图说句、以及所有 `Uncertain:` 行，在任意位置都要剪掉；命中旧缓存回写时
 * 走的也是这个函数，于是同样的脏 sidecar 会被洗掉。
 */
function cleanVisionText(raw) {
  const isLeadin = l => VISION_LEADIN.some(re => re.test(l))
  // 围栏行不是任何办公页面的正文，直接滤掉；只去围栏，不碰内容。
  let lines = String(raw ?? '').replace(/\r\n/g, '\n')
    .split('\n').filter(l => !/^[ \t]*```[a-z0-9]*[ \t]*$/i.test(l))
  // 1) 开头第一个标记行（`Transcription:` 之类）之前的内容算客套话 —— 但只在那些行确实
  //    都像客套话时才整段剪掉；前面已经是正文时宁可留着，不赌（正文优先于去噪）。
  const chatter = l => isLeadin(l) || VISION_UNCERTAIN.test(l) || isVisionCaption(l)
  for (let i = 0; i < Math.min(lines.length, 6); i++) {
    if (!isLeadin(lines[i])) continue
    if (lines.slice(0, i).filter(l => l.trim()).every(chatter)) lines = lines.slice(i + 1)
    break
  }
  // 2) 多 tile：第一片之后**每一片**自己的标记行、自我提醒、图说句，在任意位置都要剪掉
  const drop = new Set()
  for (let i = 0; i < lines.length; i++) {
    if (VISION_UNCERTAIN.test(lines[i]) || isLeadin(lines[i])) { drop.add(i); continue }
    if (!lines[i].trim()) continue
    let j = i + 1
    while (j < lines.length && !lines[j].trim()) j++
    if (j < lines.length && isLeadin(lines[j]) && isVisionCaption(lines[i])) drop.add(i)
  }
  lines = lines.filter((_, i) => !drop.has(i))
  // 3) 开头剩下的裸描述（没有标记行跟着，只能按句式认）
  while (lines.length > 1 && VISION_DESC.test(String(lines[0]).trim())) lines.shift()
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()
}

/**
 * 视觉升级并发度：默认 1，与旧版逐页串行**完全一致**。只有显式设成 2 或 3 才生效，
 * 其余任何值（"4"、0、""、垃圾）一律回落 1。并发只发生在**页与页之间**：页内的
 * 2/4/8 分带重试永远串行（见 ocrPageText），本地引擎也永远单进程。
 */
export function visionConcurrency() {
  const n = Number(process.env.DSH_OFFICE_VISION_CONCURRENCY)
  return [2, 3].includes(n) ? n : 1
}

/**
 * 单次 ocrPdfPages 的视觉调用预算。未设 DSH_OFFICE_VISION_MAX_CALLS 时无上限（=旧行为）；
 * 显式设成 0 表示一页都不许走视觉（全部记 skipped，而不是静默截断）。
 */
export function visionBudget() {
  const raw = process.env.DSH_OFFICE_VISION_MAX_CALLS
  const n = Number(raw)
  const ok = raw !== undefined && String(raw).trim() !== '' && Number.isFinite(n) && n >= 0
  const max = ok ? Math.floor(n) : Infinity
  return { max, used: 0, enabled: max !== Infinity }
}

const visionBudgetMsg = budget => `视觉调用预算已用尽（DSH_OFFICE_VISION_MAX_CALLS=${budget.max}）`

/** 预算耗尽的错误不可分带重试：否则会白跑 2/4/8 片，每片都再撞一次墙。 */
function visionBudgetError(budget) {
  const e = new Error(visionBudgetMsg(budget))
  e.code = 'VISION_BUDGET'
  return e
}

/**
 * 无依赖并发池：共享游标 + 至多 limit 个 worker。worker 内部自己保证"一页串行"
 * （ocrPageText 的分带是顺序 await），所以这里只并发页与页。
 */
async function runPool(items, limit, worker) {
  let cursor = 0
  const run = async () => {
    while (cursor < items.length) {
      const i = cursor++
      await worker(items[i], i)
    }
  }
  await Promise.all(Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, run))
}

async function ocrOneBand(file, pageNo, exec, vision, band, stat, budget) {
  const full = renderPdfPage(file, pageNo, exec)
  const png = bandPng(full, band.index, band.count)
  const where = band.count > 1 ? `第 ${pageNo} 页的第 ${band.index + 1}/${band.count} 段（横向分带）` : `第 ${pageNo} 页`
  const args = {
    path: png,
    prompt: `逐字抄写${where}的原文：输出语言必须与图片中的文字一致（中文就输出中文，禁止翻译成英文），保持原有段落、序号、标题与列表结构，不要总结、不要改写、不要省略任何文字，只输出抄写结果本身。`,
  }
  if (budget) {
    // 检查与自增之间没有 await → 单线程下是原子的，页间并发也不会超预算
    if (budget.used >= budget.max) throw visionBudgetError(budget)
    budget.used++
  }
  stat.calls++
  const value = await vision.execute(args, exec)
  return cleanVisionText(visionText(vision, args, value))
}

/**
 * Recognise one page. Vision models cap their output length (glm-4v-flash is
 * limited to 1024 tokens), so a dense page is retried in 2/4/8 horizontal
 * bands until every band fits — this is what makes long scanned pages come
 * back complete instead of truncated. `stat` accumulates the *real* accounting
 * (every attempt, including the ones abandoned by a retry) so the caller can
 * report how many slices the page was cut into and how many calls that cost.
 */
async function ocrPageText(file, pageNo, exec, vision, bandCount = 1, depth = 0, stat = { calls: 0 }, budget = null) {
  const texts = []
  try {
    for (let i = 0; i < bandCount; i++) {
      texts.push(await ocrOneBand(file, pageNo, exec, vision, { index: i, count: bandCount }, stat, budget))
    }
    // 逐片洗过之后再整页洗一遍：冷读写盘的内容与"命中缓存时回洗"的结果必须逐字一致
    return { text: cleanVisionText(texts.join('\n')), bands: bandCount, calls: stat.calls }
  } catch (e) {
    // 预算耗尽是终局，不是"这一片太长"——重试只会再撞一次墙
    if ((e && e.code === 'VISION_BUDGET') || depth >= 2 || bandCount >= 8) throw e
    return ocrPageText(file, pageNo, exec, vision, bandCount === 1 ? 2 : bandCount * 2, depth + 1, stat, budget)
  }
}

/**
 * 本地引擎的批大小纯函数：一页不落，按 `size` 串行切块（25 → [20,5]，40 → [20,20]，0 → []）。
 * 引擎只认批大小来限制占用（`--numThread`/OMP_* 在这个 build 上不生效），所以超过
 * LOCAL_MAX_IMAGES 时必须循环跑完，绝不能把第 21 页起静默丢给视觉桥。
 */
export function chunkLocalBatches(pages, size = LOCAL_MAX_IMAGES) {
  const n = Number.isInteger(size) && size > 0 ? size : LOCAL_MAX_IMAGES
  const out = []
  for (let i = 0; i < pages.length; i += n) out.push(pages.slice(i, i + n))
  return out
}

// ---------------------------------------------------------------------------
// 质量门失败页的自动换倍率重试（P1）
// ---------------------------------------------------------------------------
//
// 病灶形态：同一页在原生分辨率下被 `gateResult` 判"置信度低/版面复杂"，换个渲染
// 倍率就能过（实测：某样本第 17 页原生不过、scale=2/1.5 都过）。旧版把这类页直接
// 标 failed 丢给人工，等于"丢一页"。这里改成自动救回来。
//
// 红线：重试**串行**（逐倍率一批，不做页级并发），渲染次数计入现有渲染记账；
// 倍率与引擎边长上限必须同步（render 走子进程 env，引擎走显式参数）。

/** 备选倍率默认表；`DSH_OFFICE_OCR_RETRY_SCALES` 可调，空串 = 关闭重试。 */
const OCR_RETRY_SCALES_DEFAULT = '2,1.5,1'

/**
 * 解析重试倍率表：去重、保序、只留 0.5–4 的合法值。
 * 空串 / 全是垃圾值 → `[]`（关闭重试，回到旧行为）。
 */
export function ocrRetryScales(raw = process.env.DSH_OFFICE_OCR_RETRY_SCALES) {
  const spec = raw === undefined ? OCR_RETRY_SCALES_DEFAULT : String(raw)
  const out = []
  for (const chunk of spec.split(/[,;\s]+/)) {
    if (!chunk) continue
    const n = Number(chunk)
    if (!Number.isFinite(n) || n < 0.5 || n > 4) continue
    if (!out.includes(n)) out.push(n)
  }
  return out
}

/** 当前生效的渲染倍率（未设 → null = 原生分辨率）。 */
export function currentRenderScale(raw = process.env.DSH_OFFICE_RENDER_SCALE) {
  if (raw === undefined || String(raw).trim() === '') return null
  const n = Number(raw)
  return Number.isFinite(n) && n >= 0.5 && n <= 4 ? n : null
}

/** 倍率 → 目录/记账标签（2 → "2"，1.5 → "1.5"，原生 → ""）。 */
function scaleTag(scale) {
  return scale === null || scale === undefined ? '' : String(scale).replace(/[^\w.]/g, '')
}

/**
 * OCR the given 1-based PDF pages. Preference order: the sidecar cache, then the
 * **local** RapidOCR-json engine (offline, ~0.5 s/page, costs no model call), and
 * only the pages the local engine missed or distrusted escalate to the vision
 * bridge. `mode`: 'auto' (default) | 'local' (never call vision) | 'vision'.
 */
async function ocrPdfPages(file, pageNumbers, exec, tools, mode = 'auto', docTotal, identity) {
  // 任务一：identity 由 readPdf 从 `extra.identity` 带进来（loadModel 解析时已用源字节算好，
  // 零额外读盘）；缺省时 readOcrCache 自己同步算一次。写回时复用同一份身份，绝不二次哈希。
  const store = readOcrCache(file, identity)
  const cache = store.pages
  const srcOf = new Map(store.src)            // 每页由哪个引擎产出；旧 sidecar 没有这项 → 该页来源未知
  const retryOf = new Map(store.retry || [])  // 历史换倍率记录：跨批累积（stats 只报本批，见 info.retried）
  const total = docTotal || store.total
  const done = new Map()
  let cacheDirty = false
  for (const p of pageNumbers) {
    if (!cache.has(p)) continue
    const t = String(cache.get(p) || '')
    const healed = cleanVisionText(t)          // 修复之前生成的 sidecar 里可能还夹着视觉桥前缀，命中时顺手洗掉
    if (healed !== t) { cache.set(p, healed); cacheDirty = true }
    done.set(p, { page: p, text: healed, cached: true, engine: 'cache', blank: !healed.trim(), src: srcOf.get(p) })
  }
  const missing = pageNumbers.filter(p => !done.has(p))
  const vision = mode === 'local' ? null : tools?.get?.('modlens_read_image')
  let engine = null
  try { if (mode !== 'vision') engine = findEngine() } catch { engine = null }   // 引擎探测异常不能连累已缓存的页
  const info = {
    list: [], engineName: engine ? engineLabel(engine) : '', escalated: [], forcedVision: [],
    localError: '', total, visionCalls: 0, visionSkipped: [], visionPlanned: 0,
    retried: new Map(),        // page → 最终胜出的倍率（换倍率才救回来的页）
    retryFailed: new Map(),    // page → 全倍率都没过时的原因（含"已试过 scale=…"）
    cacheDirNote: cacheDirState().note,
    // 旧解析器 / 身份不符的 sidecar：页号或正文可能不是这份源的，已经整份作废
    // （见 PDF_PARSER_VERSION 与缓存身份段）；reason 让提示能说清是"哪一类"作废。
    staleCache: store.stale ? { path: store.path, parser: store.parser, reason: store.staleReason } : null,
  }
  if (missing.length && !engine && !vision) {
    throw new Error(`没有可用的识别引擎：本地 RapidOCR-json 未找到（把 RapidOCR-json_v0.2.0 放到 ${join(PLUGIN_DIR, 'vendor')}，或设 DSH_OFFICE_RAPIDOCR_DIR 指向引擎目录），当前 profile 也未装配视觉桥 modlens_read_image，而第 ${missing.join('、')} 页没有可复用的 OCR 缓存`)
  }

  // 1) local OCR: render every page once, then hand them to the engine in
  //    LOCAL_MAX_IMAGES-sized batches, strictly one batch at a time. The engine
  //    only limits its footprint via batch size, so a request bigger than the
  //    batch cap must loop — never silently push page 21+ to the vision bridge.
  if (engine && missing.length) {
    sweepRenderCacheOnce(file)                                     // 每 file|mtime 一次：先清旧渲染产物
    const baseScale = currentRenderScale()
    const baseTag = scaleTag(baseScale)
    const rendered = renderPdfPages(file, missing, exec, baseTag, baseScale)   // 一次文档加载出齐所有页
    const pages = missing.filter(p => rendered.get(p))
    for (const pageNo of missing) {
      if (!rendered.get(pageNo)) done.set(pageNo, { page: pageNo, text: '', cached: false, engine: LOCAL_ENGINE, failed: `第 ${pageNo} 页栅格化失败` })
    }
    const hard = []                                                // 本地没过质量门的页：先换倍率救，救不回才交视觉桥
    for (const batch of chunkLocalBatches(pages)) {
      let res = []
      try {
        const r = await ocrImages(batch.map(p => rendered.get(p)), engine, { scale: baseTag || undefined })
        res = r.pages
      } catch (e) {
        // 一批起不来（spawn 失败）：记下原因并停止后续批，剩余页走原有的视觉/失败路径
        info.localError = e.message
        break
      }
      for (let i = 0; i < batch.length; i++) {
        const r = res[i]
        const verdict = gateResult(r)
        // 没过门的页带上 retryable 标记，供下面重试分流（结构性失败不换倍率）。
        if (verdict.hard) {
          hard.push({ page: batch[i], reason: verdict.reason || '本地引擎未通过质量门', retryable: verdict.retryable !== false })
          continue
        }
        const text = verdict.blank ? '' : String(r.text || '').trim()
        done.set(batch[i], { page: batch[i], text, cached: false, engine: LOCAL_ENGINE, blank: verdict.blank, score: r.avg, boxes: r.boxes.length })
        cache.set(batch[i], text)
        srcOf.set(batch[i], LOCAL_ENGINE)
      }
    }

    // ---- 换倍率重试：逐倍率一批，串行；取"通过质量门且置信度最高"的一次 ----
    // P1：按失败原因分流——`短句占比高` 是版面**结构**特征（这页本来就是表格/数字），
    // 换倍率只让同一批碎句更清晰、占比几乎不动，纯属浪费时间；只有"置信度低 / 版面复杂 /
    // 无结果"这类才值得换倍率救。结构页直接短路（下面 structural 循环只记失败原因）。
    if (hard.length) {
      const retryable = hard.filter(h => h.retryable)
      const structural = hard.filter(h => !h.retryable)
      const candidates = ocrRetryScales().filter(s => s !== baseScale)
      const triedOf = new Map(retryable.map(h => [h.page, []]))
      const best = new Map()                                       // page → { r, scale }
      for (const s of candidates) {
        const tag = scaleTag(s)
        let again
        try { again = renderPdfPages(file, retryable.map(h => h.page), exec, tag, s) } catch { again = new Map() }
        for (const h of retryable) triedOf.get(h.page).push(s)
        const usable = retryable.map(h => h.page).filter(p => again.get(p))
        for (const batch of chunkLocalBatches(usable)) {
          let res2 = []
          try {
            const r = await ocrImages(batch.map(p => again.get(p)), engine, { scale: tag })
            res2 = r.pages
          } catch (e) {
            info.localError = info.localError || e.message
            break
          }
          for (let i = 0; i < batch.length; i++) {
            const rr = res2[i]
            const v = gateResult(rr)
            if (v.hard || v.blank) continue                        // 判空白保持原语义：不当"救回来"
            const text = String(rr.text || '').trim()
            if (!text) continue
            const prev = best.get(batch[i])
            if (!prev || (Number(rr.avg) || 0) > (Number(prev.r.avg) || 0)) best.set(batch[i], { r: rr, scale: s })
          }
        }
      }
      for (const h of retryable) {
        const win = best.get(h.page)
        const tried = [baseScale, ...triedOf.get(h.page)].filter(v => v !== null)
        if (win) {
          const text = String(win.r.text || '').trim()
          done.set(h.page, { page: h.page, text, cached: false, engine: LOCAL_ENGINE, blank: false, score: win.r.avg, boxes: win.r.boxes.length, retryScale: win.scale })
          cache.set(h.page, text)
          srcOf.set(h.page, LOCAL_ENGINE)                           // src 仍记 rapidocr：引擎没变，只是换了倍率
          info.retried.set(h.page, win.scale)
        } else {
          info.retryFailed.set(h.page, `${h.reason}${tried.length ? `；已试过 scale=${tried.join('/')}` : ''}`)
        }
      }
      // 结构性失败（短句占比高）：不换倍率，直接记原因——别在"注定失败的重试"上白花时间。
      // 这些页仍会照常交视觉桥（done 里没有它们），只是不再烧渲染+引擎时间。
      for (const h of structural) {
        info.retryFailed.set(h.page, `${h.reason}（结构性失败，换倍率无效）`)
      }
    }
  }

  // 2) escalate what the local engine could not settle
  //    页与页之间可选并发（DSH_OFFICE_VISION_CONCURRENCY，默认 1=与旧版逐页串行一致）；
  //    视觉调用总预算由 budget 把关（DSH_OFFICE_VISION_MAX_CALLS，默认无上限）。
  const budget = visionBudget()
  const todo = []
  for (const pageNo of missing) {
    if (done.has(pageNo)) continue                                   // 本地已定案（含"无文字页"）或已带失败原因
    if (!vision) {
      done.set(pageNo, {
        page: pageNo, text: '', cached: false, engine: null,
        // 优先报"换过哪些倍率"的证据（该页被本地重试过），否则才退回泛化的质量门话术
        failed: info.localError || info.retryFailed.get(pageNo) || (mode === 'local'
          ? '本地引擎未通过质量门，且 ocrEngine:"local" 禁止回落视觉桥'
          : '本地引擎未通过质量门，且当前 profile 未装配视觉桥 modlens_read_image'),
      })
      continue
    }
    todo.push(pageNo)
  }
  info.visionPlanned = todo.length
  if (todo.length) {
    const worker = async pageNo => {
      // 预算在此之前已用尽 → 这一页一次都没开始过，记 skipped（区别于"试到一半撞墙"的 failed）
      if (budget.used >= budget.max) {
        done.set(pageNo, { page: pageNo, text: '', cached: false, engine: null, skipped: true, failed: visionBudgetMsg(budget) })
        info.visionSkipped.push(pageNo)
        return
      }
      try {
        const vp = await ocrPageText(file, pageNo, exec, vision, 1, 0, { calls: 0 }, budget)
        // 谁让它走视觉的：用户指定 / 没装本地引擎 / 本地没过质量门。三者的话术不同，
        // 不能把用户指定说成"本地没把握"（那会诱导去调 DSH_OFFICE_OCR_MIN_SCORE）。
        const via = mode === 'vision' ? 'user' : engine ? 'review' : 'no-local'
        done.set(pageNo, { page: pageNo, text: vp.text, cached: false, engine: 'vision', bands: vp.bands, calls: vp.calls, via })
        cache.set(pageNo, vp.text)
        srcOf.set(pageNo, 'vision')
        info.visionCalls += vp.calls
        if (via === 'review') info.escalated.push(pageNo)
        else info.forcedVision.push(pageNo)
      } catch (e) {
        done.set(pageNo, { page: pageNo, text: '', cached: false, engine: 'vision', failed: e.message })
      }
    }
    await runPool(todo, visionConcurrency(), worker)
    // 并发完成后页码必须有序：pageRanges / 脚注文案要与串行版逐字一致
    info.escalated.sort((a, b) => a - b)
    info.forcedVision.sort((a, b) => a - b)
    info.visionSkipped.sort((a, b) => a - b)
  }

  info.list = pageNumbers.map(p => done.get(p) ?? { page: p, text: '', cached: false, engine: null, failed: '未识别' })
  info.covered = [...cache.keys()].sort((a, b) => a - b)
  if (cacheDirty || info.list.some(r => !r.cached)) {
    // manifest 的 retry: 要**累积**：本批救回来的页叠在历史之上，否则下一次写会把上一批的倍率账抹掉
    const retryMerged = new Map(retryOf)
    for (const [p, s] of info.retried) retryMerged.set(p, s)
    writeOcrCache(file, cache, { src: srcOf, total, retry: retryMerged, identity: store.identity || identity })
  }
  return info
}

/** Human-readable provenance for one OCR'd page. */
function ocrSourceLabel(r, requested = 'auto') {
  if (r.cached) {
    // 命中缓存必须说清"这是谁的结果"：用户换引擎排查质量问题时，静默回旧结果
    // 会让他误判"换引擎无效"。auto 不问引擎，维持原来的短标注不加负担。
    if (requested === 'auto') return 'OCR 缓存'
    const got = r.src === LOCAL_ENGINE ? 'local' : r.src === 'vision' ? 'vision' : ''
    if (!got) return `OCR 缓存 · 来源未记录（指定的 ${requested} 未执行）`
    if (got === requested) return `OCR 缓存 · ${got}`
    return `OCR 缓存 · ${got}（指定的 ${requested} 未执行）`
  }
  if (r.engine === LOCAL_ENGINE) return `本地 OCR · RapidOCR${r.boxes != null ? `（${r.boxes} 框，置信 ${r.score}）` : ''}`
  if (r.engine === 'vision') {
    // 切片/调用次数是记账口径：一页切了几片、真实调了几次，不能只说"1 页"
    const cuts = r.bands > 1 ? ` · 切 ${r.bands} 片（${r.calls} 次调用）` : ''
    if (r.via === 'user') return `视觉模型识别 · 用户指定 ocrEngine:"vision"${cuts}`
    if (r.via === 'no-local') return `视觉模型识别 · 本地引擎不可用${cuts}`
    return `视觉模型识别 · 本地复核${cuts}`
  }
  return 'OCR 识别'
}

/**
 * `stats.ocrEngine` 的口径：本次**真正执行**的是谁。
 * 全命中缓存就说 `cache`，不要拿"可用引擎"冒充"跑过的引擎"（本地引擎装了但一次没跑的情况很容易被误读）。
 */
function ocrEngineUsed(info, pages) {
  if (!info) return pages.length ? 'unavailable' : 'not-needed'
  if (!pages.length) return 'not-needed'
  const ran = info.list.filter(r => !r.cached && !r.failed)
  const local = ran.some(r => r.engine === LOCAL_ENGINE)
  const vision = ran.some(r => r.engine === 'vision')
  if (local && info.engineName) return vision ? `${info.engineName} + 视觉桥` : info.engineName
  if (vision) return 'vision'
  if (ran.length) return 'unavailable'
  return info.list.some(r => r.cached) ? 'cache' : 'unavailable'
}

/**
 * The provenance marker written after every OCR'd page body. Kept as one pattern
 * next to the emitter, and matched greedily to line end because the label itself
 * nests full-width parens（像 `（9 框，置信 0.988）`）.
 */
const OCR_PAGE_MARK_RE = /^_（第 (\d+) 页：.*）_[ \t]*$/gm

/**
 * Split OCR'd read output back into per-page bodies, keyed by 1-based page number.
 * This is the only content-agnostic way to compare a cold read with a cache hit —
 * the cache must reproduce every page verbatim, nothing before/after the markers.
 */
function ocrPageBodies(text) {
  const src = String(text || '')
  const re = new RegExp(OCR_PAGE_MARK_RE.source, 'gm')
  const map = new Map()
  let last = 0
  let m
  while ((m = re.exec(src))) {
    map.set(Number(m[1]), src.slice(last, m.index).trim())
    last = re.lastIndex
  }
  return map
}

/**
 * 任务 H（opt-in，**只改呈现**）：`DSH_OFFICE_BLANKS=1` 时把 PDF 正文里的填空下划线
 * `_{3,}` 渲成可读的占位 `<span class="blank">＿＿＿＿</span>`。
 *
 * 红线：它**不 gate 任何正确性修复**（关掉它，文本层提取 / 质量门 / OCR / 缓存全部逐字不变），
 * 生效边界也严格限定在 `office_read` 的 **PDF 分支** + `as="markdown"` / `as="json"`：
 * `as="meta"` 不经过这里；`as="text"` 按规格**不**生效（要原文）。对 OCR 路线天然无效
 * —— 识别产物里本来就没有 `_`（那是另一件事，需要单独对 OCR 文本做后处理）。
 */
function blanksPresentation(raw = process.env.DSH_OFFICE_BLANKS) {
  const v = String(raw ?? '').trim().toLowerCase()
  return v === '1' || v === 'true' || v === 'on' || v === 'yes'
}

/** `_{3,}` → 等宽占位。全角下划线保证在中文正文里宽度可读。 */
function renderBlanks(text) {
  return String(text).replace(/_{3,}/g, m => `<span class="blank">${'＿'.repeat(Math.max(3, Math.round(m.length * 0.6)))}</span>`)
}

/** 只遍历字符串叶子，绝不碰结构（JSON 仍然可解析）。 */
function renderBlanksDeep(v) {
  if (typeof v === 'string') return renderBlanks(v)
  if (Array.isArray(v)) return v.map(renderBlanksDeep)
  if (v && typeof v === 'object') {
    const out = {}
    for (const k of Object.keys(v)) out[k] = renderBlanksDeep(v[k])
    return out
  }
  return v
}

/**
 * 把 PDF 的 document 模型按页切组（`pdfToDocument` 在每页正文前插了一个二级标题 `第 N 页`）。
 * 单页 PDF 没有页标题 → 返回唯一一组、`page === null`。
 */
function splitDocByPage(doc) {
  const groups = []
  let cur = { page: null, blocks: [] }
  for (const b of doc?.blocks || []) {
    const m = b.type === 'heading' && Number(b.level) === 2 ? /^第 (\d+) 页$/.exec(String(b.text || '')) : null
    if (m) {
      if (cur.page !== null || cur.blocks.length) groups.push(cur)
      cur = { page: Number(m[1]), blocks: [b] }
    } else cur.blocks.push(b)
  }
  if (cur.page !== null || cur.blocks.length) groups.push(cur)
  return groups
}

/**
 * PDF 的 `as="json"` 出站收口（第三轮连带修复）。
 *
 * 旧行为两处都会坏事：① 忽略 `pages`，永远给整本模型；② 超 `READ_CAP` 时 `capText` 会在
 * **字符串中间**切断 → 返回**无法 parse 的半截 JSON**。页序 bug 修好之前 ② 碰不到（每页塌成
 * 一行、整本模型才几十 KB），修好之后 43 页中文 PDF 的模型立刻 > 200000 字符。
 *
 * 现在：honor `pages`（按 `第 N 页` 分段筛）；仍超限就**按页丢尾部**并回报丢了哪几页 ——
 * 保证"返回的永远是合法 JSON"，且丢弃显式记账，绝不静默。
 */
function capJsonPayload(doc, args, selected) {
  const base = doc && typeof doc === 'object' && Array.isArray(doc.blocks) ? doc : null
  if (!base) {
    const s = JSON.stringify(doc ?? null, null, 2)
    return { content: capText(s, READ_CAP), truncated: s.length > READ_CAP, droppedPages: [], kept: 0 }
  }
  let groups = splitDocByPage(base)
  if (args?.pages && selected?.length) {
    const want = new Set(selected.map(s => s.page))
    groups = groups.filter(g => (g.page === null ? want.size > 0 : want.has(g.page)))
  }
  const build = gs => JSON.stringify({ ...base, blocks: gs.flatMap(g => g.blocks) }, null, 2)
  if (build(groups).length <= READ_CAP) return { content: build(groups), truncated: false, droppedPages: [], kept: groups.length }
  // 按页丢尾部，找到能装下的最大前缀（至少保留 1 页，否则退化成空模型）
  let keep = groups.length
  while (keep > 1 && build(groups.slice(0, keep)).length > READ_CAP) keep--
  // 单页仍超限：连这一页也放不下，就只留元数据（仍然是合法 JSON）
  if (build(groups.slice(0, keep)).length > READ_CAP) {
    const bare = { ...base, blocks: [], truncateNote: `单页 JSON 超过内联上限 ${READ_CAP} 字符，已省略正文` }
    return { content: JSON.stringify(bare, null, 2), truncated: true, droppedPages: groups.map(g => g.page).filter(Boolean), kept: 0 }
  }
  const dropped = groups.slice(keep).map(g => g.page).filter(Boolean)
  return { content: build(groups.slice(0, keep)), truncated: true, droppedPages: dropped, kept: keep }
}

/**
 * PDF read path: page selection, scanned-page detection, optional OCR (local
 * engine first, vision bridge for the doubtful pages), and per-page annotations
 * for pages without a text layer.
 */
async function readPdf(args, exec, file, model, extra, ocrMode, tools, ocrEngine = 'auto') {
  const noText = list => list.filter(s => !String(s.text || '').replace(/\s/g, '').length)
  // —— 需求 3：页级续读语义 —— pageFrom/pageTo（含端点）= 从第 N 页读到第 M 页。
  // 字符级续读（offset/limit）已有；页级续读更符合 PDF 直觉：每批边界就是页面边界，
  // 调用方不再需要知道每页正文落在字符流的哪个偏移上。
  if (args.pageFrom !== undefined || args.pageTo !== undefined) {
    if (args.pages) throw new Error('pages 与 pageFrom/pageTo 不能同时给：二选一（任意页码集合用 pages，页级区间续读用 pageFrom/pageTo）')
    const from = Math.max(1, Number(args.pageFrom) || 1)
    const to = Math.min(extra.pages || from, Number(args.pageTo) || extra.pages || from)
    if (from > to) throw new Error(`页码范围无效：pageFrom=${from} > pageTo=${to}（全书 ${extra.pages} 页）`)
    if (from > (extra.pages || 1)) throw new Error(`页码范围 "${from}-${to}" 未命中（该 PDF 共 ${extra.pages} 页）`)
    args = { ...args, pages: `${from}-${to}` }
  }
  let selected = extra.sections
  if (args.pages) {
    const wanted = parsePageSpec(args.pages, extra.pages)
    selected = extra.sections.filter(s => wanted.has(s.page))
    if (!selected.length) throw new Error(`页码范围 "${args.pages}" 未命中（该 PDF 共 ${extra.pages} 页）`)
  }
  if (args.as === 'meta') {
    const stats = statsOf(model, extra)
    stats.pagesWithText = extra.sections.length - noText(extra.sections).length
    stats.scannedPages = noText(extra.sections).map(s => s.page)
    stats.fallback = 'none'
    // —— 质量画像：文字层"有没有"（pagesWithText/scannedPages）与"能不能用"（这里）是两件事 ——
    // CID 乱码样本的 textFound=true、scannedPages=[]，旧版 meta 因此看起来完全健康。
    const prof = textLayerProfile(extra.sections)
    stats.textLayerUsable = Boolean(extra.textFound) && prof.garbled.length === 0
    stats.garbledPages = garbledSpec(prof.garbled)
    stats.qualityGate = { testedPages: prof.testedPages, garbledPages: garbledSpec(prof.garbled), reasons: prof.reasons }
    // 缓存迁移 + 缓存身份（任务一）：旧解析器 / 身份不符的 .ocr.md 现状要能从 meta 一眼看到，
    // 且**不触发**渲染/OCR（identity 已在 loadModel 解析时算好，这里零额外读盘）
    const cacheState = readOcrCache(file, extra.identity)
    if (cacheState.stale) {
      stats.ocrCacheStale = {
        path: cacheState.path, parser: cacheState.parser ?? null,
        current: PDF_PARSER_VERSION, reason: cacheState.staleReason || null,
      }
      stats.ocrCacheNote = staleCacheNote(cacheState)
    }
    if (prof.garbled.length) {
      const per = Math.min(20, Math.max(1, extra.pages || 20))
      stats.suggestion = {
        note: `文字层不可信（${prof.garbled.length} 页命中乱码），整本需 OCR`,
        copy: `read --path="${file}" --ocr=always --ocrEngine=local --pages="1-${per}"`,
      }
    } else {
      const suggestion = readSuggestion('pdf', model, extra, args, extra.sections.reduce((a, s) => a + String(s.text || '').length, 0))
      if (suggestion) stats.suggestion = suggestion
    }
    const metaPayload = { path: file, format: 'pdf', meta: model?.meta || {}, stats }
    return { ...metaPayload, content: JSON.stringify(metaPayload, null, 2) }
  }
  if (args.as === 'json') {
    const stats = statsOf(model, extra)
    const bare = noText(extra.sections)
    stats.pagesWithText = extra.sections.length - bare.length
    // 无文本层时 JSON 只剩页级空壳；不解释原因容易被读成"文件没内容"
    const notice = bare.length && bare.length === extra.sections.length
      ? `该 PDF 无文本层（${bare.length}/${extra.sections.length} 页为扫描/图片页）：下面的 JSON 只有空的页级结构，正文需 OCR —— 用 ocr:"always" + pages="1-5" 识别，结果缓存为同名 .ocr.md`
      : ''
    const payload = model && typeof model === 'object' ? (notice ? { notice, ...model } : model) : model
    // 任务 H（json 分支）：只改字符串叶子，JSON 结构不受影响
    const shown = args.as === 'json' && blanksPresentation() ? renderBlanksDeep(payload) : payload
    stats.fallback = 'none'
    // —— 第三轮连带修复：json 分支原来是 `capText(JSON.stringify(...))`，两处都会坏事 ——
    // ① PDF 的 json 分支**忽略 `pages`**，永远给整本模型；
    // ② 一旦超过 `READ_CAP`，`capText` 会在**字符串中间**切断 → 返回**无法 parse 的半截 JSON**。
    // 页序 bug 修好之前 ② 碰不到（每页塌成一行，整本模型才几十 KB）；修好之后单页正文量恢复正常，
    // 43 页中文 PDF 的模型立刻 >200000 字符，于是"as=json 拿到不可解析内容"立刻暴露。
    // 收口方式：honor `pages`；仍超限就**按页丢尾部**（保证是合法 JSON）并把丢弃写进 notice/status，
    // 绝不返回半截 JSON、也绝不静默丢。
    const { content: jsonContent, truncated, droppedPages, kept } = capJsonPayload(shown, args, selected)
    stats.truncateNote = truncated
      ? `JSON 模型超过内联上限 ${READ_CAP} 字符，已按页保留前 ${kept} 页`
        + `${droppedPages?.length ? `，省略第 ${pageRanges(droppedPages)} 页` : ''}；用 pages 分批取或调大 DSH_OFFICE_MAX_INLINE_CHARS`
      : undefined
    return {
      path: file, format: 'pdf', meta: model?.meta || {}, stats, notice: notice || undefined,
      truncated, content: jsonContent,
    }
  }

  const emptySelected = noText(selected)
  // —— 质量门：文档**有**文本层，但文本层不可信（CID 字体缺 ToUnicode 的乱码）——
  // 这类页在旧版被当成"正常读到正文"，于是把乱码灌进上下文。基线取本批全部文本，
  // CJK 覆盖率只作相对判据（英文文档天然为 0，不能当闸门）。
  const qualityBaseline = textQuality(selected.map(s => String(s.text || '')).join('\n'))
  const suspectPages = []
  const suspectQuality = new Map()
  for (const s of selected) {
    if (!String(s.text || '').replace(/\s/g, '').length) continue
    // 需求 4：与 textLayerProfile 同口径（structural 启发）——meta 说的与 read 遇到的一致
    const q = textQuality(s.text, qualityBaseline, { structural: true })
    if (q.garbled) { suspectPages.push(s.page); suspectQuality.set(s.page, q) }
  }
  let ocrPages = []
  /** 单次读取的 OCR 页数被 `OCR_MAX_PAGES` 砍掉时的记账（null = 没砍）。 */
  let ocrPagesCapped = null
  /**
   * "未指定 pages 只预览前 N 页"的记账（null = 用户指定了 pages 或没触发预览）。
   * 未指定 `pages` 时 `read` 只 OCR 前 `OCR_DEFAULT_PAGES` 页，此前**完全静默**，
   * 做批量收益对比时会被它误导出"批量更慢"的反向结论。这里抬到 stats + 脚注。
   * （第八轮自 WB 侧 P1-2 回移。）
   */
  let ocrPreviewLimited = null
  if (ocrMode !== 'never') {
    if (ocrMode === 'always') {
      if (args.pages) {
        ocrPages = selected.map(s => s.page)
      } else {
        const wanted = selected.map(s => s.page)
        ocrPages = wanted.slice(0, OCR_DEFAULT_PAGES)
        if (wanted.length > ocrPages.length) {
          ocrPreviewLimited = { preview: OCR_DEFAULT_PAGES, total: wanted.length, skipped: wanted.slice(OCR_DEFAULT_PAGES) }
        }
      }
    } else {
      // auto：① 无文本层的页 ② 文本层判定为乱码的页 —— 两类都去识别
      const auto = [...new Set([...emptySelected.map(s => s.page), ...suspectPages])].sort((a, b) => a - b)
      if (auto.length) {
        if (args.pages) {
          ocrPages = auto
        } else {
          ocrPages = auto.slice(0, OCR_DEFAULT_PAGES)
          if (auto.length > ocrPages.length) {
            ocrPreviewLimited = { preview: OCR_DEFAULT_PAGES, total: auto.length, skipped: auto.slice(OCR_DEFAULT_PAGES) }
          }
        }
      }
    }
    // 单次读取的 OCR 页数硬上限（OCR_MAX_PAGES）**必须明说**。
    // 实测：`pages=1-30 ocr="always"` 只会识别前 20 页，剩下 10 页静默落回文本层；
    // 而 stats 里只有 `ocrCovered: "1-20 / 35"`，看不出"我要的 1-30 被砍了 10 页"。
    // 这是"静默少做"的典型形态，所以既不静默、也不改变原有上限。
    const ocrWanted = ocrPages.slice()
    ocrPages = ocrPages.slice(0, OCR_MAX_PAGES)
    if (ocrPages.length < ocrWanted.length) {
      ocrPagesCapped = { limit: OCR_MAX_PAGES, requested: ocrWanted.length, applied: ocrPages.length, skipped: ocrWanted.slice(OCR_MAX_PAGES) }
    }
  }
  const ocrMap = new Map()
  let ocrError
  let ocrInfo = null
  if (ocrPages.length) {
    try {
      ocrInfo = await ocrPdfPages(file, ocrPages, exec, tools, ocrEngine, extra.pages, extra.identity)
      for (const r of ocrInfo.list) ocrMap.set(r.page, r)
    } catch (e) {
      ocrError = e.message
    }
  }
  if (!extra.textFound && !ocrMap.size && ocrMode === 'never') {
    throw new Error('pdf: 未提取到文本层（扫描件/图片型 PDF，需要 OCR）。去掉 ocr:"never" 即可自动 OCR，或用 pages 指定页码后重试。')
  }

  const multi = extra.pages > 1 || Boolean(args.pages)
  // 任务 H：只在 markdown 路径生效（json 走下面的 renderBlanksDeep；meta/text 不受影响）
  const showBlanks = args.as === 'markdown' && blanksPresentation()
  const parts = []
  for (const s of selected) {
    const head = multi ? `<!-- 第 ${s.page} 页 -->\n` : ''
    const ocr = ocrMap.get(s.page)
    if (ocr && String(ocr.text || '').trim()) {
      parts.push(`${head}${String(ocr.text).trim()}\n\n_（第 ${s.page} 页：${ocrSourceLabel(ocr, ocrEngine)}）_`)
    } else if (ocr && ocr.blank) {
      parts.push(`${head}_(第 ${s.page} 页：OCR 判定为无文字页（空白/纯图页），未消耗视觉识别)_`)
    } else if (!String(s.text || '').replace(/\s/g, '').length) {
      const hint = ocrMode === 'never' ? '；可用 ocr:"always" 识别' : ''
      parts.push(`${head}_(本页未提取到文本层，疑似扫描/图片页${hint}${ocr && ocr.failed ? `；识别未完成：${ocr.failed}` : ''})_`)
    } else {
      // 文本层有字但质量门判为乱码、且 OCR 没能接管 → 保留原文（宁留噪不删正文）
      // 但必须显式标注，否则会被读成"正文就这样"
      const q = suspectQuality.get(s.page)
      const mark = q
        ? `\n\n_（本页文本层疑似乱码、未能重识别：${q.reasons.join('；')}${ocr && ocr.failed ? `；OCR 未完成：${ocr.failed}` : ''}）_`
        : ''
      parts.push(`${head}${showBlanks ? renderBlanks(s.text) : s.text}${mark}`)
    }
  }
  let content = parts.join('\n\n')
  const notes = []
  if (ocrPages.length) {
    const tally = new Map()
    for (const r of ocrMap.values()) {
      const k = r.failed ? '未完成' : r.cached ? '缓存' : r.engine === LOCAL_ENGINE ? (r.blank ? '本地·判空白' : '本地') : r.engine === 'vision' ? '视觉桥' : '未完成'
      tally.set(k, (tally.get(k) || 0) + 1)
    }
    const summary = [...tally].map(([k, v]) => `${k} ${v} 页`).join(' / ') || '均未成功'
    // 视觉桥的"页"不等于"调用"：切了片就要把口径写出来（片数/调用数）
    const vCalls = ocrInfo?.visionCalls || 0
    const calls = vCalls > (tally.get('视觉桥') || 0) ? `（${vCalls} 次视觉调用）` : ''
    notes.push(`已 OCR 第 ${ocrPages.join('、')} 页（${summary}${calls}${ocrInfo?.engineName ? `｜${ocrInfo.engineName}` : ''}）；结果缓存为同名 .ocr.md`)
    if (ocrInfo?.escalated?.length) notes.push(`本地引擎对第 ${ocrInfo.escalated.join('、')} 页没把握，已交视觉模型复核`)
    if (ocrInfo?.forcedVision?.length) notes.push(`第 ${ocrInfo.forcedVision.join('、')} 页按 ocrEngine:"${ocrEngine}" 指定走视觉识别（非本地引擎自动升级）`)
    // 预算护栏：跳过的页必须点名，绝不静默少识别（只在设了 DSH_OFFICE_VISION_MAX_CALLS 时才可能出现）
    if (ocrInfo?.visionSkipped?.length) {
      const planned = ocrInfo.visionPlanned || 0
      const skipped = ocrInfo.visionSkipped.length
      const finished = [...ocrMap.values()].filter(r => r.engine === 'vision' && !r.failed).length
      notes.push(`第 ${pageRanges(ocrInfo.visionSkipped)} 页因视觉调用预算上限跳过（计划 ${planned} / 完成 ${finished} / 跳过 ${skipped}）`)
    }
    if (ocrInfo?.localError) notes.push(`本地 OCR 未完成：${ocrInfo.localError}`)
    // 换倍率重试要记账：这些页是"降级救回来"的，用户有权知道原始分辨率没过质量门
    if (ocrInfo?.retried?.size) {
      const pairs = [...ocrInfo.retried].sort((a, b) => a[0] - b[0])
      notes.push(`第 ${pageRanges(pairs.map(([p]) => p))} 页本地引擎在默认倍率下没通过质量门，已自动换倍率重试成功（`
        + `${pairs.map(([p, s]) => `${p}→scale=${s}`).join('，')}；DSH_OFFICE_OCR_RETRY_SCALES 可调）`)
    }
    if (ocrInfo?.cacheDirNote) notes.push(ocrInfo.cacheDirNote)
    // 旧解析器的 sidecar 被整份作废这件事必须**说出来**，否则用户只会看到"缓存没了"
    if (ocrInfo?.staleCache) notes.push(staleCacheNote(ocrInfo.staleCache))
    // 指定了引擎却吃到别家的缓存：说清楚没跑，否则会被误读成"换引擎无效"
    if (ocrEngine !== 'auto') {
      const hit = ocrInfo?.list?.filter(r => r.cached) || []
      const mism = hit.filter(r => r.src && r.src !== (ocrEngine === 'local' ? LOCAL_ENGINE : 'vision'))
      const unknown = hit.filter(r => !r.src)
      if (mism.length) {
        notes.push(`第 ${pageRanges(mism.map(r => r.page))} 页命中 ${[...new Set(mism.map(r => r.src === LOCAL_ENGINE ? 'local' : 'vision'))].join('/')} 缓存，指定的 ${ocrEngine} 未执行；如需强制重识别请删 ${ocrCachePath(file)}`)
      } else if (unknown.length) {
        notes.push(`第 ${pageRanges(unknown.map(r => r.page))} 页命中旧格式缓存（未记录来源），无法确认指定的 ${ocrEngine} 是否跑过；如需强制重识别请删 ${ocrCachePath(file)}`)
      }
    }
  }
  // 全书覆盖状态：只报"这批做了几页"是不够的，几十页文件会静默蒸发剩余页
  const scanned = noText(extra.sections).map(s => s.page)
  const covered = ocrInfo?.covered || []
  const uncovered = scanned.filter(p => !covered.includes(p))
  if (ocrPages.length && uncovered.length) {
    notes.push(`第 ${pageRanges(uncovered)} 页未识别（无文本层共 ${scanned.length} 页，已覆盖 ${pageRanges(covered) || '无'}${ocrInfo?.total ? ` / 全书 ${ocrInfo.total} 页` : ''}）`)
  }
  // 两种"静默少做"都必须显式记账 + 给可复制的续读命令（自 WB 侧第八轮 P1-2 回移）：
  // ① 未指定 pages 只做前 N 页预览；② 要求页数被 OCR_MAX_PAGES 砍掉。
  // 顺序与 WB 侧一致（capped 先、preview 后）——第八轮 P3 对账 #1 的唯一真实偏差：
  // 两条文案本身逐字一致，仅当同一次读取同时命中"预览截断 + 20 页上限"时脚注行序不同。
  // 文案本体统一由 ocrCapNotes() 产出（R9-3）：降级路径（finishRead）要重建同一批脚注，
  // 靠两份手写模板互相对齐迟早漂移 —— 这里与那边共用同一个函数。
  notes.push(...ocrCapNotes(ocrPagesCapped, ocrPreviewLimited, file))
  const remaining = noText(extra.sections)
  if (ocrMode === 'auto' && !ocrMap.size && remaining.length && !args.pages) {
    notes.push(`全文共 ${remaining.length} 页无文本层（第 ${remaining.slice(0, 10).map(s => s.page).join('、')}${remaining.length > 10 ? '…' : ''} 页），可用 pages 指定页码后重试`)
  }
  if (ocrError) notes.push(`OCR 未完成：${ocrError}`)
  // DSH 补充（任务六-1 / F2 F3）：批内有缺页（或整批报错）时，把 sidecar 现状一起端出来 ——
  // 上一轮会话被迫靠"反复触发失败调用 + 手动 grep sidecar"自救，这里把那条路变成官方路。
  const failedPages = [...ocrMap.values()].filter(r => r.failed)
  if (ocrPages.length && (failedPages.length || ocrError)) {
    const cov = sidecarCoverage(file, extra.identity)
    if (failedPages.length) {
      notes.push(`批内缺页：${failedPages.map(r => `第 ${r.page} 页（${String(r.failed).split('\n')[0].trim().slice(0, 80)}）`).join('、')}`
        + (cov ? `；${cov}` : '；无可用 sidecar（缺页未落缓存）'))
    } else if (cov) {
      notes.push(cov)
    }
  }
  // CID 乱码兜底：直接给**可复制**的参数串，别让模型自己拼
  const suspectNoOcr = suspectPages.filter(p => !String(ocrMap.get(p)?.text || '').trim())
  if (suspectPages.length) {
    const why = [...new Set(suspectPages.map(p => suspectQuality.get(p)?.reasons?.[0]).filter(Boolean))].join('；')
    notes.push(`第 ${pageRanges(suspectPages)} 页文本层疑似乱码${why ? `（${why}）` : ''}`)
    if (suspectNoOcr.length) {
      notes.push(`第 ${pageRanges(suspectNoOcr)} 页没能重识别，可直接复制：`
        + `office_read path="${file}" ocr="always" ocrEngine="local" pages="${pageRanges(suspectNoOcr)}"`
        + `（本地 RapidOCR，离线、不消耗视觉额度）`)
    } else {
      notes.push(`第 ${pageRanges(suspectPages)} 页已用 OCR 结果替代原文`)
    }
  }
  if (notes.length) content += `\n\n> ${notes.join('；')}`
  if (args.as === 'text') content = content.replace(/<!--[^>]*-->\s*/g, '').replace(/^_\((.*)\)_$/gm, '（$1）')

  // —— 需求 3：批间衔接（跨批读取的连续性保障）——
  // 分批（pages 指定且未覆盖全书）时，把"上一批末行回看 / 下一批首行预览"（各约 100 字符）
  // 作为**增量字段**挂到 stats（prevTail / nextHead / nextPage），并在 notice 里给一行摘要；
  // boundary=true 时再以引用行内联进正文首尾（格式固定为 `> 〔dsh-office 批间…〕`，拼接时剔除）。
  // 默认**不改 content**：分批拼接与整本读取的逐字 diff 依然成立，旧调用方零影响。
  const boundaryInfo = { prevTail: '', nextHead: '', prevPage: 0, nextPage: 0 }
  if (args.pages && selected.length) {
    const byPage = new Map(extra.sections.map(sec => [sec.page, sec]))
    const firstSel = selected[0].page
    const lastSel = selected[selected.length - 1].page
    const previewOf = (sec, which) => {
      const t = String(sec?.text || '').replace(/\s+/g, ' ').trim()
      if (!t) return ''
      return which === 'tail' ? (t.length > 100 ? `…${t.slice(-100)}` : t) : (t.length > 100 ? `${t.slice(0, 100)}…` : t)
    }
    const prevSec = firstSel > 1 ? byPage.get(firstSel - 1) : null
    const nextSec = lastSel < extra.pages ? byPage.get(lastSel + 1) : null
    boundaryInfo.prevPage = prevSec ? firstSel - 1 : 0
    boundaryInfo.nextPage = nextSec ? lastSel + 1 : 0
    boundaryInfo.prevTail = prevSec ? previewOf(prevSec, 'tail') : ''
    boundaryInfo.nextHead = nextSec ? previewOf(nextSec, 'head') : ''
    if (args.boundary && args.as === 'markdown') {
      if (boundaryInfo.prevTail) content = `> 〔dsh-office 批间回看｜第 ${boundaryInfo.prevPage} 页末尾〕${boundaryInfo.prevTail}\n\n${content}`
      if (boundaryInfo.nextHead) content += `\n\n> 〔dsh-office 批间预览｜第 ${boundaryInfo.nextPage} 页开头〕${boundaryInfo.nextHead}`
    }
  }

  const offset = Math.max(0, Number(args.offset) || 0)
  const limit = Math.min(Math.max(100, Number(args.limit) || READ_CAP), READ_CAP)
  const body = offset ? content.slice(offset) : content
  const capped = capWithOffset(body, limit, offset)
  const stats = {
    ...statsOf(model, extra),
    ocrPages: [...ocrMap.keys()].sort((a, b) => a - b),
    ocrEngine: ocrEngineUsed(ocrInfo, ocrPages),
    ocrRequested: ocrEngine,
    ocrCovered: ocrInfo ? `${pageRanges(ocrInfo.covered) || '无'}${ocrInfo.total ? ` / ${ocrInfo.total}` : ''}` : undefined,
    ocrUncovered: uncovered.length ? pageRanges(uncovered) : undefined,
    // 第八轮（自 WB 侧 P1-2 回移）：未指定 pages 只做前 N 页预览、以及被 OCR_MAX_PAGES 砍掉的页，
    // 形状与 WB 侧逐字一致（ocrPagesCapped.skipped 是页码数组，ocrPreview.skipped 是区间字符串）。
    ocrPagesCapped: ocrPagesCapped || undefined,
    ocrPreview: ocrPreviewLimited
      ? { preview: ocrPreviewLimited.preview, total: ocrPreviewLimited.total, skipped: pageRanges(ocrPreviewLimited.skipped) }
      : undefined,
    ocrVisionCalls: ocrInfo?.visionCalls || undefined,
    // 只在预算护栏真的拦下页时才外显（未设 DSH_OFFICE_VISION_MAX_CALLS 时永远不出现）
    ocrVisionSkipped: ocrInfo?.visionSkipped?.length ? pageRanges(ocrInfo.visionSkipped) : undefined,
    ocrEscalated: ocrInfo?.escalated || [],
    ocrFailed: [...ocrMap.values()].filter(r => r.failed).map(r => r.page),
    // DSH 补充（任务三-3 / F3 的账）：批内每页失败都要"页码 + 一句话原因"一起给，
    // 调用方**不 grep sidecar 也能看到缺哪页、为什么缺**。空数组 = 本批无缺页（不是缺字段）。
    ocrFailedPages: [...ocrMap.values()].filter(r => r.failed)
      .map(r => ({ page: r.page, reason: String(r.failed).split('\n')[0].trim().slice(0, 160) }))
      .sort((a, b) => a.page - b.page),
    // 降级记账：none=文本层干净可用；ocr=走了识别；sidecar=已转存文件（由 finishRead 决定）
    fallback: ocrPages.length ? 'ocr' : 'none',
    ocrFromQualityGate: suspectPages.length ? pageRanges(suspectPages) : undefined,
  }
  // 换倍率重试的记账：只有真发生过才外显，形状与 sidecar manifest 的 retry: 一致
  if (ocrInfo?.retried?.size) {
    stats.ocrRetried = pageRanges([...ocrInfo.retried.keys()])
    stats.ocrRetryScale = Object.fromEntries([...ocrInfo.retried].sort((a, b) => a[0] - b[0]).map(([p, s]) => [String(p), Number(s)]))
  }
  if (capped.note) stats.truncateNote = capped.note
  // 需求 3 的增量字段：本批覆盖页区间 + 下一未读页 + 批间回看/预览（约 100 字符各一）。
  // 整本读取或全书覆盖时不出现这些字段（与旧返回逐字一致）。
  if (args.pages && selected.length) {
    stats.pageFrom = selected[0].page
    stats.pageTo = selected[selected.length - 1].page
    if (boundaryInfo.nextPage) stats.nextPage = boundaryInfo.nextPage
    if (boundaryInfo.prevTail) stats.prevTail = { page: boundaryInfo.prevPage, text: boundaryInfo.prevTail }
    if (boundaryInfo.nextHead) stats.nextHead = { page: boundaryInfo.nextPage, text: boundaryInfo.nextHead }
  }
  const out = {
    path: file,
    format: 'pdf',
    meta: model?.meta || {},
    stats,
    truncated: capped.truncated,
    nextOffset: capped.nextOffset,
    content: capped.content,
  }
  // 截断说明只放 notice（content 里一个字符都不掺，见 capWithOffset 的协议说明）
  if (capped.note) out.notice = capped.note
  // 批间衔接的一行摘要：boundary=true 时正文已内联，notice 不重复
  if (!args.boundary && (boundaryInfo.prevTail || boundaryInfo.nextHead)) {
    const bits = []
    if (boundaryInfo.prevTail) bits.push(`上承第 ${boundaryInfo.prevPage} 页末行「${boundaryInfo.prevTail}」`)
    if (boundaryInfo.nextHead) bits.push(`下接第 ${boundaryInfo.nextPage} 页开头「${boundaryInfo.nextHead}」`)
    const msg = `批间衔接（详见 stats.prevTail/nextHead/nextPage）：${bits.join('；')}`
    out.notice = typeof out.notice === 'string' && out.notice ? `${out.notice}；${msg}` : msg
  }
  // 同 non-PDF 分支（下方 `nonPdfOut.__fullBody = content`）：截断过时必须把**截断前的完整正文**
  // 交给 finishRead，sidecar 才真落得下"整篇"（见该函数头注释）。第十轮补回 —— 第三轮误删了本行，
  // 而 finishRead 一直在消费 `out.__fullBody`：PDF 正文被内联护栏截断且质量门判乱码时，sidecar
  // 会落**截断版**，notice 承诺的"整篇正文"就成了假话。
  if (capped.truncated) out.__fullBody = content
  return out
}

// ---------------------------------------------------------------------------
// 需求 2：批量文件 stats 探测（office_read 的 paths 批量形态）
// ---------------------------------------------------------------------------
//
// 病灶：盘点 25 个 PDF 需要 25 次 as="meta"（每次一文件）。这里给一次调用返回
// 逐文件轻量 stats 的批量形态：`office_read paths=[...]`（条目可以是文件或**目录**——
// 目录会展开为其中受支持的办公文件）。行字段与单文件 meta.stats 同名：
// format / pages / characters / textLayerUsable / scannedPages / garbledPages，
// 另加 suggestedBatches（建议分批数）。只聚合元数据，**绝不返回正文**。
//
// 单文件失败不连累整批（inventory 场景缺一行错误信息比整批报错有用），但
// 一个文件都扫不到时按四要素报错。

const SCAN_EXTS = new Set([
  'pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'doc', 'xls', 'ppt',
  'wps', 'et', 'dps', 'pps', 'csv', 'tsv', 'md', 'txt', 'rtf', 'html', 'htm',
  'json', 'jsonl',
])
const SCAN_MAX_FILES = 500

/** 单文件轻量画像：绝不返回正文；错误进行内 error 字段而不是抛出。 */
async function scanOneRow(file, exec) {
  const row = { path: file, name: basename(file) }
  try {
    const buf = await readBuffer(file)
    throwIfAborted(exec)
    const ext = extOf(file)
    if (ext === 'pdf' || sniff(buf, ext) === 'pdf') {
      const pdf = await readPdfMemo(file, buf)
      const chars = pdf.sections.reduce((a, s) => a + String(s.text || '').length, 0)
      const scanned = pdf.sections.filter(s => !String(s.text || '').replace(/\s/g, '').length)
      const prof = textLayerProfile(pdf.sections)          // 纯 CPU：与 meta 同一套质量门
      row.format = 'pdf'
      row.pages = pdf.pages
      row.characters = chars
      row.textLayerUsable = Boolean(pdf.textFound) && prof.garbled.length === 0
      row.scannedPages = scanned.length
      row.garbledPages = prof.garbled.length
      if (prof.garbled.length) row.garbledPageSpec = garbledSpec(prof.garbled)
      // 建议分批数与 meta.suggestion 同口径：干净书 15 页/批；乱码书走 OCR 20 页/批
      row.suggestedBatches = prof.garbled.length
        ? Math.max(1, Math.ceil((pdf.pages || 1) / 20))
        : (pdf.pages > 12 ? Math.ceil(pdf.pages / 15) : 1)
      return row
    }
    const { kind, model } = await loadModel(file, exec)
    row.format = kind
    if (typeof model === 'string') row.characters = model.length
    else row.characters = modelCharCount(model)
    const probe = typeof model === 'string' ? model : modelToText(model)
    const q = textQuality(probe)
    row.textLayerUsable = !q.garbled
    row.scannedPages = 0
    row.garbledPages = 0
    if (model?.kind === 'document') row.blocks = normalizeDocument(model).blocks.length
    if (model?.kind === 'workbook') {
      const w = normalizeWorkbook(model)
      row.sheets = w.sheets.map(s => ({ name: s.name, rows: s.rows.length }))
    }
    if (model?.kind === 'slides') row.slides = normalizeSlides(model).slides.length
    const cap = maxInlineChars()
    row.suggestedBatches = Number.isFinite(cap) && row.characters > cap
      ? Math.ceil(row.characters / cap)
      : 1
  } catch (e) {
    row.error = String(safeMessage(e)).split('\n')[0].slice(0, 200)
  }
  return row
}

/**
 * 批量扫描入口：paths（string 数组，条目可为文件或目录）→ 逐文件 stats 清单。
 * 返回值：`{ format: 'scan', total, ok, failed, files: [...], content: <管道表> }`。
 */
async function scanFilesStats(args, exec) {
  if (args.as && args.as !== 'meta') {
    throw new Error(`批量扫描（paths）只返回逐文件 stats，不返回正文，不支持 as="${args.as}"；请去掉 as（或用 as="meta"）`)
  }
  const raw = (Array.isArray(args.paths) ? args.paths : []).map(p => String(p ?? '').trim()).filter(Boolean)
  if (!raw.length) throw new Error('【批量扫描】paths 为空：至少给一个文件或目录路径；下一步=office_read paths=["<目录或文件>", …]')
  const files = []
  const skipped = []
  for (const p of raw) {
    throwIfAborted(exec)
    const fp = hostPath(p, exec)
    let st = null
    try { st = statSync(fp) } catch { /* 不存在，走 skipped */ }
    if (!st) { skipped.push({ path: fp, reason: '路径不存在' }); continue }
    if (st.isDirectory()) {
      let names = []
      try { names = readdirSync(fp) } catch (e) { skipped.push({ path: fp, reason: `目录不可读（${e.code || e.message}）` }); continue }
      // 目录展开只收用户文档：插件自己的识别/兜底缓存（.ocr.md / .read.md / .ocr.json）不算盘点对象
      const found = names.filter(n => SCAN_EXTS.has(extOf(n)) && !isProtectedCacheName(n)).map(n => join(fp, n)).sort()
      if (found.length) files.push(...found)
      else skipped.push({ path: fp, reason: '目录内没有受支持的办公文件' })
    } else files.push(fp)
  }
  const seen = new Set()
  const list = []
  for (const f of files) {
    const k = f.toLowerCase()
    if (!seen.has(k)) { seen.add(k); list.push(f) }
  }
  let truncated = false
  if (list.length > SCAN_MAX_FILES) { list.length = SCAN_MAX_FILES; truncated = true }
  if (!list.length) {
    throw new Error('【批量扫描】没有可扫描的文件（页码=全部；格式=清单；根因='
      + `${skipped.map(s => `${s.path}（${s.reason}）`).join('、') || '未给出有效路径'}；`
      + `下一步=确认路径与扩展名后重试，受支持扩展名：${[...SCAN_EXTS].map(e => '.' + e).join(' ')}`)
  }
  const rows = []
  for (const f of list) rows.push(await scanOneRow(f, exec))
  const okCount = rows.filter(r => !r.error).length
  const cwd = sessionCwd(exec)
  const show = f => {
    const rel = relativePath(cwd, f)
    return rel && !rel.startsWith('..') && rel.length < f.length ? rel : f
  }
  const esc = s => String(s ?? '').replace(/\|/g, '\\|').replace(/\r?\n/g, ' ')
  const lines = ['| 文件 | 格式 | 页 | 字符 | 文字层 | 扫描页 | 乱码页 | 建议批次 | 备注 |', '| --- | --- | ---: | ---: | --- | ---: | ---: | ---: | --- |']
  for (const r of rows) {
    if (r.error) {
      lines.push(`| ${esc(show(r.path))} | — | — | — | — | — | — | — | ❌ ${esc(r.error)} |`)
    } else {
      lines.push(`| ${esc(show(r.path))} | ${esc(r.format)} | ${r.pages ?? '—'} | ${r.characters ?? '—'}`
        + ` | ${r.textLayerUsable ? '✓' : '✗'} | ${r.scannedPages} | ${r.garbledPages}${r.garbledPageSpec ? `（${esc(r.garbledPageSpec)}）` : ''}`
        + ` | ${r.suggestedBatches} | ${esc(r.sheets ? r.sheets.map(s => `${s.name}:${s.rows}行`).join('、') : '')} |`)
    }
  }
  const bits = [`共 ${list.length} 个文件：成功 ${okCount} / 失败 ${rows.length - okCount}`]
  if (skipped.length) bits.push(`跳过 ${skipped.length} 项（${skipped.map(s => `${basename(s.path)}：${s.reason}`).join('；')}）`)
  if (truncated) bits.push(`超过单次上限 ${SCAN_MAX_FILES} 个，已截断（可分目录多次扫描）`)
  const garbled = rows.filter(r => !r.error && r.garbledPages > 0)
  if (garbled.length) {
    bits.push(`${garbled.length} 个文件存在乱码页：${garbled.map(r => `${r.name}（第 ${r.garbledPageSpec} 页）`).join('、')}——盘点后按 office_read path=… ocr="always" 逐个处理`)
  }
  const unusable = rows.filter(r => !r.error && r.textLayerUsable === false)
  if (unusable.length > garbled.length) {
    bits.push(`${unusable.length} 个文件文字层不可用（含非 PDF 乱码），建议先 as="meta" 逐个看 qualityGate`)
  }
  return {
    format: 'scan',
    total: list.length,
    ok: okCount,
    failed: rows.length - okCount,
    files: rows,
    ...(skipped.length ? { skipped } : {}),
    ...(truncated ? { truncated: true } : {}),
    notice: bits.join('；'),
    content: `${lines.join('\n')}\n\n> ${bits.join('；')}`,
  }
}

// ---------------------------------------------------------------------------
// plugin
// ---------------------------------------------------------------------------

function apply(ctx, config = {}) {
  void config
  const tools = ctx.get('tools')
  if (!tools) throw new Error('dsh-office: 需要 tools 服务（请确认插件装配在 dsh 基础组合的 host 层）')

  tools.register(defineToolLite({
    name: 'office_read',
    description: '读取办公文档为 Markdown / 纯文本 / 结构化 JSON。支持 Word(.docx/.doc/.wps)、Excel(.xlsx/.xls/.csv/.tsv)、PPT(.pptx/.ppt/.dps)、PDF(.pdf，含扫描件自动 OCR)、OpenDocument(.odt/.ods/.odp)、网页(.html/.htm，style/script 剥离、嵌套列表/表格/实体保留)、Markdown/文本(.md/.txt/.json/.rtf)。path 支持绝对路径或相对会话工作目录；paths=[...] 为批量形态：一次返回逐文件轻量 stats 清单（format/pages/characters/textLayerUsable/scannedPages/garbledPages/建议分批数，条目可为文件或目录，不返回正文）；as=meta 只看元数据；超长内容用 offset/limit 续读；PDF 页级续读用 pageFrom/pageTo（从第 N 页读到第 M 页，等价 pages="N-M" 并附 stats.pageFrom/pageTo/nextPage）；分批读取默认在 stats.prevTail/nextHead 附上一批末行回看与下一批首行预览（boundary=true 时以 "> 〔dsh-office 批间…〕" 引用行内联进正文首尾）；扫描版 PDF 默认自动 OCR（ocr:"auto"，可用 pages 指定页码）——本地 RapidOCR 引擎优先，可疑页才回落视觉模型（ocrEngine:"local" 可完全不消耗视觉调用）。可选环境变量（默认行为不变）：DSH_OFFICE_VISION_CONCURRENCY=2|3 让视觉复核页并发；DSH_OFFICE_VISION_MAX_CALLS=N 限单次视觉调用总数；DSH_OFFICE_RENDER_SCALE=0.5~4 按 96dpi×scale 栅格化；DSH_OFFICE_OCR_RETRY_SCALES=2,1.5,1 让没过质量门的页自动换渲染倍率重试；DSH_OFFICE_CACHE_DIR=<目录> 让 .ocr.md / .read.md 落到该目录；DSH_OFFICE_MAX_INLINE_CHARS=120000 内联体积护栏阈值。',
    parameters: {
      path: { type: 'string', description: '办公文件路径（绝对或相对会话目录）；与 paths 二选一' },
      paths: { type: 'array', items: { type: 'string' }, description: '（批量形态）文件/目录路径数组：一次返回逐文件轻量 stats 清单（行字段与 meta.stats 一致），不返回正文；条目为目录时展开其中受支持的办公文件' },
      as: { type: 'string', enum: ['markdown', 'text', 'meta', 'json'], description: '输出形态，默认 markdown。**json：content 是 JSON 文本字符串**（`JSON.stringify(model, null, 2)` 的产物 —— 调用方需自行 `JSON.parse`，它**不是**对象；根形如 {kind:"document"|"workbook"|"slides", meta, blocks|sheets|slides}）。json/meta 形态：不写 sidecar、不带 nextOffset、不受 DSH_OFFICE_MAX_INLINE_CHARS 约束（由 READ_CAP=200000 字符收口），超限时按页丢弃尾部并置 truncated=true + stats.truncateNote，但**返回内容永远是合法 JSON，绝不返回半截**。meta：把元数据包成 JSON 文本放进 content。text：剥离标记的纯正文' },
      sheet: { type: 'string', description: '（表格）只读取指定工作表名' },
      pages: { type: 'string', description: '（PDF）页码范围，如 "1-5" 或 "2,4,7"，默认全部' },
      pageFrom: { type: 'integer', description: '（PDF）页级续读：起始页（含）。与 pageTo 搭配，不能与 pages 同给' },
      pageTo: { type: 'integer', description: '（PDF）页级续读：结束页（含），缺省=最后一页。pageFrom 缺省=第 1 页' },
      boundary: { type: 'boolean', description: '（PDF 分批）把上一批末行回看/下一批首行预览以引用行内联进正文首尾（默认 false=只进 stats.prevTail/nextHead 与 notice）' },
      ocr: { type: 'string', enum: ['auto', 'always', 'never'], description: '（PDF）扫描页处理：auto=无文本层或质量门判乱的页自动 OCR（默认，未指定 pages 时最多 3 页）；always=强制 OCR 指定页；never=只报文本层情况' },
      ocrEngine: { type: 'string', enum: ['auto', 'local', 'vision'], description: '（PDF）识别引擎：auto=本地 RapidOCR 优先、可疑页回落视觉模型（默认）；local=只用本地引擎，不消耗视觉调用；vision=只用视觉模型' },
      offset: { type: 'integer', description: '内容起始字符偏移，用于超长文档续读' },
      limit: { type: 'integer', description: '最多返回字符数，默认 200000' },
      recalc: { type: 'boolean', description: '（表格 xlsx/ods/csv…）用内置轻量求值器**重算公式**（第十二轮需求 3；第二轮 R13 扩函数子集）：命中的公式格返回计算值，并在 stats.recalc 里给出 {formulaCells, evaluated, unsupported, errors, details:[{sheet,cell,formula,value,error}]}。支持函数子集：SUM AVERAGE MIN MAX COUNT COUNTA COUNTIF COUNTIFS SUMIF AVERAGEIF IF VLOOKUP ROUND ROUNDUP ROUNDDOWN ABS INT MOD LEFT RIGHT MID LEN TRIM UPPER LOWER CONCAT CONCATENATE TEXT VALUE 与四则/比较/&/百分比/区域引用；区域里的文本/布尔按真 Excel 口径忽略（直接参数仍换算）。**不支持的函数显式标 unsupported 并保留原值，绝不猜**；整列引用 A:A / 外部工作簿引用 / _xlfn 前缀同样显式记账（带单元格地址）。默认 false —— 不设时返回值与旧版逐字一致；大表可用 DSH_OFFICE_RECALC_MAX_CELLS 设单元格上限（超限记 skipped，不静默截断）。' },
    },
    render: (_a, v) => {
      if (typeof v?.content === 'string') return [{ type: 'text', text: v.content }]
      const rest = { ...(v ?? {}) }
      delete rest.content
      return [{ type: 'text', text: JSON.stringify(Object.keys(rest).length ? rest : v ?? null, null, 2) }]
    },
    async execute(args, exec) {
      // 需求 2：批量形态 —— paths（含空数组）显式路由；单文件读取要求 path
      if (Array.isArray(args.paths)) {
        if (args.path) throw new Error('path 与 paths 只能二选一：单文件读取用 path，批量扫描用 paths')
        if (!args.paths.length) throw new Error('【批量扫描】paths 为空：至少给一个文件或目录路径；下一步=office_read paths=["<目录或文件>", …]')
        return scanFilesStats(args, exec)
      }
      if (args.path === undefined) throw new Error('需要提供 path（单文件读取）或 paths（批量扫描，字符串数组），两者二选一')
      const file = hostPath(args.path, exec)
      let kind = extOf(file)
      try {
        const loaded = await loadModel(file, exec)
        kind = loaded.kind
        const { model: loadedModel, extra } = loaded
        let model = loadedModel
        // ---- 需求 3（第十二轮）：表格公式重算（显式 opt-in；不设 recalc 时行为与旧版逐字相同）----
        let recalcReport = null
        if (args.recalc) {
          if (model && model.kind === 'workbook') {
            const rr = recalcWorkbook(model)
            model = rr.model
            recalcReport = rr.report
          } else {
            recalcReport = { skipped: `recalc 只对表格生效，本文件是 ${kind}` }
          }
        }
        const ocrMode = args.ocr || 'auto'
        const isPdf = kind === 'pdf' && Array.isArray(extra?.sections)
        if (isPdf) return finishRead(await readPdf(args, exec, file, model, extra, ocrMode, tools, args.ocrEngine || 'auto'), args, file)
        if (args.as === 'meta') {
          const stats = statsOf(model, extra)
          stats.fallback = 'none'
          // 非 PDF 也有"文字层可信度"：docx/xlsx 里塞乱码同样要让第一眼看出来（纯 CPU，不渲染）
          const probe = typeof model === 'string' ? model : modelToText(model)
          const q = textQuality(probe)
          stats.textLayerUsable = !q.garbled
          if (q.garbled) stats.qualityGate = { reasons: q.reasons }
          const suggestion = readSuggestion(kind, model, extra, args, modelCharCount(model))
          if (suggestion) stats.suggestion = suggestion
          if (recalcReport) stats.recalc = recalcReport
          const metaPayload = { path: file, format: kind, meta: model?.meta || {}, stats }
          return { ...metaPayload, content: JSON.stringify(metaPayload, null, 2) }
        }
        let content
        if (args.as === 'json') content = JSON.stringify(model ?? null, null, 2)
        else if (args.as === 'text') {
          // 需求 1：html 的 as="text" 给真正的纯文本（块边界换行、表格制表符分隔），
          // 不残留 markdown 的 # / | / ** 标记；其他格式维持旧的 modelToText 行为逐字不变
          content = kind === 'html' ? plainTextOf(model) : modelToText(model)
        } else content = modelToMarkdown(model)
        if (args.sheet) {
          const w = normalizeWorkbook(model)
          const sh = w.sheets.find(s => s.name === args.sheet)
          if (!sh) throw new Error(`找不到工作表 "${args.sheet}"（现有: ${w.sheets.map(s => s.name).join(', ')}）`)
          content = `## 工作表: ${sh.name}\n` + tableMd(sh.rows.map(r => r.map(cellText)), false)
        }
        const offset = Math.max(0, Number(args.offset) || 0)
        const limit = Math.min(Math.max(100, Number(args.limit) || READ_CAP), READ_CAP)
        const body = offset ? content.slice(offset) : content
        const capped = capWithOffset(body, limit, offset)
        const stats = { ...statsOf(model, extra), fallback: 'none' }
        if (capped.note) stats.truncateNote = capped.note
        if (recalcReport) stats.recalc = recalcReport
        const nonPdfOut = {
          path: file,
          format: kind,
          meta: model?.meta || {},
          stats,
          truncated: capped.truncated,
          nextOffset: capped.nextOffset,
          content: capped.content,
        }
        if (capped.note) nonPdfOut.notice = capped.note
        // 同 readPdf：截断过时把完整正文交给 finishRead，sidecar 才真落得下"整篇"
        if (capped.truncated) nonPdfOut.__fullBody = content
        return finishRead(nonPdfOut, args, file)
      } catch (e) {
        if (e && e.__dshOfficeReadHint) throw e
        // 错误分支四要素（页码 / 格式 / 根因 / 可复制的下一步），绝不把裸错误抛给模型
        const err = new Error(`${safeMessage(e)}\n${readErrorHint(e, args, kind, file)}`)
        err.__dshOfficeReadHint = true
        throw err
      }
    },
  }))

  tools.register(defineToolLite({
    name: 'office_create',
    description: '创建办公文件（目标格式由扩展名决定）：.docx/.odt/.pdf（文档）、.xlsx/.ods/.csv/.tsv（表格）、.pptx/.odp（演示）、.md/.txt/.json/.html（.html 输出语义化 HTML5：h1-h6/table/列表/引用/加粗，UTF-8 + &lt;meta charset&gt;，中文不转实体）。内容来源任选其一：markdown（文档/演示）、document/slides/workbook/table 结构化对象（表格与精确排版推荐）、from（从既有文件转换式创建）。.csv 默认 UTF-8 with BOM（Excel 双击不乱码），含逗号/引号/换行的字段自动按 RFC 4180 转义；encoding 参数可为文本类目标（csv/tsv/md/txt）显式指定 utf-8-sig（带 BOM）或 utf-8（无 BOM）。',
    parameters: {
      path: { type: 'string', required: true, description: '目标文件路径（含扩展名）' },
      markdown: { type: 'string', description: 'Markdown 源内容（# 标题/- 列表/| 表格 |/``` 代码块）' },
      document: { type: 'object', description: '文档模型 {meta, blocks:[{type:"heading"|"paragraph"|"list"|"table"|"code"|"quote"|"hr"|"pagebreak", ...}]}' },
      slides: { type: 'object', description: '演示模型 {meta, slides:[{title, subtitle, bullets:[{text,level}], table:{rows}, notes, layout}]}' },
      workbook: { type: 'object', description: '表格模型 {meta, sheets:[{name, columns:[{title,width}], rows:[[值或{v,t,f}]]}]}' },
      table: { type: 'object', description: '简单二维表 {columns:[..], rows:[[..]]}' },
      from: { type: 'string', description: '从该既有文件读取内容再按目标扩展名写出（即转换）' },
      encoding: { type: 'string', enum: ['utf-8-sig', 'utf-8'], description: '（csv/tsv/md/txt 目标）写出编码：utf-8-sig=带 BOM，utf-8=无 BOM；默认 csv 带 BOM、其余不带。gb18030 等旧代码页写出需非零依赖实现，暂不提供（读取侧仍自动识别 GBK）' },
      grid: { type: 'string', types: ['string', 'number'], description: '（.docx 目标）**申论稿纸网格**：写 `<w:docGrid w:type="linesAndChars">`，把每行字数/每页行数钉死。接受**数字 20**，或字符串 "20"、"20x25"（分隔符也接受 X / × / *，如 "20X25"）；每行字数/每页行数都会被夹到 2..60。非 .docx 目标会显式报错（不静默忽略）。注意：这只是 Word 的"文档网格"，可见的方格线框属于页面背景，不在本参数范围' },
    },
    render: (_a, v) => [{ type: 'text', text: `已创建 ${v.path}（${v.format}，${v.bytes} 字节；${v.summary}）` }],
    async execute(args, exec) {
      const file = hostPath(args.path, exec)
      throwIfAborted(exec)
      const ext = extOf(file)
      if (!(ext in WRITERS)) throw new Error(`暂不支持创建 .${ext || '(无扩展名)'}；可创建: ${Object.keys(WRITERS).map(e => '.' + e).join(' ')}`)
      if (args.encoding !== undefined && !['csv', 'tsv', 'md', 'txt'].includes(ext)) {
        throw new Error(`【写出编码】encoding 只对文本类目标（csv/tsv/md/txt）生效，目标 .${ext} 不支持；下一步=去掉 encoding，或把目标改为 .csv/.tsv/.md/.txt`)
      }
      const provided = ['document', 'slides', 'workbook', 'table', 'markdown', 'from'].filter(k => args[k] !== undefined && args[k] !== '')
      if (provided.length > 1) throw new Error(`内容来源只能选一个，收到: ${provided.join(', ')}`)
      let model = null
      if (args.document) model = normalizeDocument(args.document)
      else if (args.slides) model = normalizeSlides(args.slides)
      else if (args.workbook) model = normalizeWorkbook(args.workbook)
      else if (args.table) model = normalizeWorkbook({ sheets: [{ name: 'Sheet1', rows: (args.table.rows || []).map(r => (Array.isArray(r) ? r : [r])) }] })
      else if (typeof args.markdown === 'string' && args.markdown.trim()) {
        model = ext === 'pptx' || ext === 'odp' ? markdownToSlides(args.markdown)
          : ['xlsx', 'ods', 'csv', 'tsv', 'json'].includes(ext) ? markdownToWorkbookModel(args.markdown)
            : markdownToDocument(args.markdown)
      } else if (args.from) {
        const loaded = await loadModel(hostPath(args.from, exec), exec)
        model = loaded.model
      } else if (ext === 'md' || ext === 'txt') {
        model = { kind: 'document', meta: {}, blocks: [] }
      } else {
        throw new Error('需要提供 markdown / document / slides / workbook / table / from 之一')
      }
      const finalModel = adaptModel(model, ext)
      // 需求 4c：申论稿纸网格（只对 docx 生效；无法识别的取值显式报错，不静默忽略）
      if (args.grid !== undefined && args.grid !== '') {
        if (ext !== 'docx') {
          throw new Error(`【稿纸网格】grid 只对 .docx 目标生效（当前目标 .${ext}）；下一步=把目标改成 .docx，或去掉 grid`)
        }
        const g = normalizeGrid(args.grid)
        if (!g) throw new Error(`【稿纸网格】grid 取值无法识别："${args.grid}"；下一步=用 "20x25"（每行字数×每页行数）或 "20" 或数字 20`)
        finalModel.meta = { ...(finalModel.meta || {}), grid: g }
      }
      // 插图兜底：目标写出端不支持 image 块时显式降级（绝不静默丢图）
      const degraded = degradeImageBlocks(finalModel, ext)
      const winfo = {}
      const bytes = await WRITERS[ext](degraded.model, args.encoding, winfo)
      // 阶段二：PDF 目标过产出质量门（临时文件 → 校验 → 原子改名；失败即删、绝不落盘）
      const pdfQuality = ext === 'pdf' ? await pdfOutputGate(file, bytes, finalModel, winfo) : null
      if (!pdfQuality) await saveBuffer(file, bytes)
      const out = { path: file, format: ext, bytes: byteLength(bytes), summary: summaryOf(finalModel) }
      if (pdfQuality) {
        out.stats = { pdfQuality }
        if (pdfQuality.notice) out.notice = pdfQuality.notice
      }
      if (degraded.degraded.length) {
        out.stats = { ...(out.stats || {}), imageFallback: degraded.degraded }
        out.notice = [out.notice, `${degraded.degraded.length} 张图片在 .${ext} 写出端不内嵌：`
          + degraded.degraded.map(d => `${d.name || '(无路径)'}（${d.alt}）`).join('；')].filter(Boolean).join('；')
      }
      // docx/别的写出端的图片记账（内嵌成功、尺寸换算、内容去重、读不到而退化）
      applyImageWriteInfo(out, winfo, ext)
      return out
    },
  }))

  // 任务四：office_edit 的 operations 元素约束。`op` 取值表 = 实现真正支持的全集（16 个），
  // schema（enum）与运行时校验（itemViolations）共用同一份，杜绝"文档说的与实现支持的"漂移。
  const EDIT_OPS_UNION = [
    'replace_text', 'append_markdown', 'set_meta', 'insert_image', 'append_image',
    'set_cell', 'append_rows', 'replace_value', 'add_sheet', 'rename_sheet', 'delete_sheet',
    'add_slide', 'update_slide', 'delete_slide', 'append_text', 'prepend_text',
  ]
  // 「与容器无关都必需」的字段：只补今天会被静默吞掉的几类（缺 newName 的 rename_sheet
  // 会写出 `Sheet1 → undefined`，缺 name 的 delete_sheet 会打印 `删除 "undefined"`…）。
  // 容器相关的必需性继续由执行期精确报错（各格式的兜底文案本来就有"可用: …"清单）。
  const EDIT_OP_RULES = {
    ops: EDIT_OPS_UNION,
    required: {
      replace_text: [{ any: ['find', 'regex'] }],
      append_markdown: [{ any: ['markdown', 'text'] }],
      insert_image: [{ any: ['path', 'base64'] }],
      append_image: [{ any: ['path', 'base64'] }],
      set_cell: [{ field: 'cell' }],
      append_rows: [{ field: 'rows' }],
      rename_sheet: [{ field: 'name' }, { field: 'newName' }],
      delete_sheet: [{ field: 'name' }],
      update_slide: [{ field: 'index' }],
      delete_slide: [{ field: 'index' }],
      append_text: [{ field: 'text' }],
      prepend_text: [{ field: 'text' }],
    },
    fields: {
      update_slide: { index: 'integer' },
      delete_slide: { index: 'integer' },
      insert_image: { width: 'number' },
      append_image: { width: 'number' },
      append_rows: { rows: 'array' },
      // R18 任务 D2：这些字段的实现一律是 `String(op.x ?? '')` —— 传数字/对象**过去会被静默
      // 字符串化**（`find: 2024` 变成搜索 "2024"，`{a:1}` 变成搜索 "[object Object]"），
      // 调用方拿不到任何提示。schema（EDIT_OP_FIELDS）本来就声明 string，这里让运行时与
      // schema 一致：非字符串给带下标的清晰文案。字符串形式逐字兼容。
      replace_text: { find: 'string', replace: 'string' },
      replace_value: { find: 'string', replace: 'string' },
      append_text: { text: 'string' },
      prepend_text: { text: 'string' },
      append_markdown: { markdown: 'string', text: 'string' },
      set_cell: { cell: 'string' },
      add_sheet: { name: 'string' },
      rename_sheet: { name: 'string', newName: 'string' },
      delete_sheet: { name: 'string' },
    },
  }
  const EDIT_OP_FIELDS = {
    op: { type: 'string', required: true, enum: EDIT_OPS_UNION, description: '操作名（必须属于目标文件类型支持的那一组）' },
    find: { type: 'string', description: 'replace_text / replace_value 的查找文本（regex=true 时是正则）' },
    replace: { type: 'string', description: '替换成的文本（缺省=删除）' },
    regex: { type: 'boolean', description: 'find 按正则解释' },
    whole: { type: 'boolean', description: '（xlsx replace_value）整格匹配，默认 true' },
    markdown: { type: 'string', description: '（docx/odt）要追加的 Markdown' },
    text: { type: 'string', description: '（append_text/prepend_text）要写入的文本' },
    sheet: { type: 'string' }, cell: { type: 'string', description: '（xlsx set_cell）A1 形式单元格' },
    value: { type: 'string' }, rows: { type: 'array' },
    name: { type: 'string', description: '（xlsx）工作表名' },
    newName: { type: 'string', description: '（xlsx rename_sheet）新工作表名' },
    index: { type: 'integer', description: '（pptx）1 起的幻灯片序号' },
    path: { type: 'string', description: '（插图）图片文件路径（PNG/JPEG/GIF/BMP）' },
    base64: { type: 'string', description: '（插图）图片数据的 base64' },
    alt: { type: 'string', description: '（插图）替代文字' },
    width: { type: 'number', description: '（插图，磅）省略/0/负数/NaN 时按原图像素换算：宽 = 原图宽px × 72/96，高按纵横比；超过 A4 可用宽 451.3pt 时等比缩到 451.3pt 并记账 capped=true' },
    after: { type: 'string', description: '（insert_image）文本锚点' },
    at: { type: 'string', enum: ['start', 'end'], description: '（append_markdown）插入位置' },
    title: { type: 'string' }, subtitle: { type: 'string' }, bullets: { type: 'array' },
    table: { type: 'object' }, notes: { type: 'string' }, layout: { type: 'string' }, slide: { type: 'object' },
  }

  tools.register(defineToolLite({
    name: 'office_edit',
    description: '原地修改已有办公文件（zip/XML 级，尽量保留原有内容与样式）。操作按 operations 数组顺序执行：docx → replace_text/append_markdown/set_meta/insert_image/append_image（插图：{path 或 base64, alt?, width?, after?（文本锚点，仅 insert_image）}，支持 PNG/JPEG/GIF/BMP；**width 单位磅，且省略时有真实换算、没有固定默认值**——省略/0/负数/NaN 时按原图像素换算：宽 = 原图宽px × 72/96，高按纵横比，超过 A4 可用宽 451.3pt 时等比缩到 451.3pt 并在 notice/stats.imageSizing 记账 capped=true）；xlsx → set_cell/append_rows/replace_value/add_sheet/rename_sheet/delete_sheet；pptx → replace_text/add_slide/update_slide/delete_slide；odt/ods/odp → replace_text/append_markdown(.odt)；csv/tsv → replace_value/replace_text/append_rows；md/txt/html → replace_text/append_text/prepend_text（.html/.htm 是**文件级文本替换**，不改结构、不重新排版）。',
    parameters: {
      path: { type: 'string', required: true, description: '要修改的文件（必须已存在）' },
      operations: {
        type: 'array',
        required: true,
        description: '操作列表，按顺序执行；每项形如 {op:"replace_text", find:"旧", replace:"新"}。op 可用集合见 items.properties.op；操作与目标格式不匹配时执行期会报"不支持操作 + 可用清单"',
        items: { type: 'object', additionalProperties: true, properties: EDIT_OP_FIELDS, itemRules: EDIT_OP_RULES },
      },
    },
    render: (_a, v) => [{ type: 'text', text: `已修改 ${v.path}\n- ${v.summary.join('\n- ')}` }],
    async execute(args, exec) {
      const file = hostPath(args.path, exec)
      throwIfAborted(exec)
      const buf = await readBuffer(file)
      const ext = extOf(file)
      const container = sniff(buf, ext)
      const zipObj = container === 'zip' ? openZip(buf) : null
      const kind = resolveKind(buf, ext, zipObj ? zipObj.names : null)
      const ops = Array.isArray(args.operations) ? args.operations : []
      if (!ops.length) throw new Error('operations 不能为空')
      let result
      if (kind === 'docx') result = await editDocx(file, ops, exec)
      else if (kind === 'xlsx') result = await editXlsx(file, ops, exec)
      else if (kind === 'pptx') result = await editPptx(file, ops, exec)
      else if (kind === 'odt' || kind === 'ods' || kind === 'odp') result = await editOdf(file, ops, exec)
      else if (kind === 'csv' || kind === 'tsv' || kind === 'text' || kind === 'md' || ext === 'txt' || ext === 'md' || ext === 'json' || ext === 'html' || ext === 'htm') result = await editTextFile(file, ops, exec)
      else throw new Error(`暂不支持编辑 .${ext}（可编辑: docx/xlsx/pptx/odt/ods/odp/csv/tsv/md/txt/html；PDF/doc/xls/ppt 请先 office_convert 转换后再编辑）`)
      await saveBuffer(file, result.bytes)
      const edited = { path: file, format: kind, bytes: byteLength(result.bytes), summary: result.summary }
      if (result.notice) edited.notice = result.notice
      if (result.stats) edited.stats = result.stats
      return edited
    },
  }))

  tools.register(defineToolLite({
    name: 'office_convert',
    description: '办公格式互转：读取 source（任意受支持格式，含 .html/.htm 网页），按统一内容模型写出 target（扩展名决定目标格式）。例：Excel→CSV、Word→PDF、HTML→Markdown、Markdown→docx/pdf/pptx/html、PPT→Markdown（.html 目标是语义化 HTML5，UTF-8 + &lt;meta charset&gt;）。写盘前会校验源文件文本层质量：整本乱码且无 OCR 缓存时**直接拒绝**（绝不产出二进制垃圾），有完整 .ocr.md 缓存则用 OCR 文本替代这些页并返回 stats.fallback="ocr" / stats.garbledPages / notice。.csv 目标默认 UTF-8 with BOM（Excel 双击不乱码）；encoding 参数可为文本类目标（csv/tsv/md/txt）显式指定 utf-8-sig / utf-8。',
    parameters: {
      source: { type: 'string', required: true, description: '源文件路径' },
      target: { type: 'string', required: true, description: '目标文件路径（扩展名决定格式）' },
      sheet: { type: 'string', description: '（表格源）只转换指定工作表' },
      encoding: { type: 'string', enum: ['utf-8-sig', 'utf-8'], description: '（csv/tsv/md/txt 目标）写出编码：utf-8-sig=带 BOM，utf-8=无 BOM；默认 csv 带 BOM、其余不带' },
    },
    render: (_a, v) => [{ type: 'text', text: `${v.source} (${v.sourceFormat}) → ${v.target} (${v.targetFormat})，${v.bytes} 字节` }],
    async execute(args, exec) {
      const source = hostPath(args.source, exec)
      const target = hostPath(args.target, exec)
      throwIfAborted(exec)
      const ext = extOf(target)
      if (!(ext in WRITERS)) throw new Error(`暂不支持转换为 .${ext || '(无扩展名)'}；可转换目标: ${Object.keys(WRITERS).map(e => '.' + e).join(' ')}`)
      if (args.encoding !== undefined && !['csv', 'tsv', 'md', 'txt'].includes(ext)) {
        throw new Error(`【写出编码】encoding 只对文本类目标（csv/tsv/md/txt）生效，目标 .${ext} 不支持；下一步=去掉 encoding，或把目标改为 .csv/.tsv/.md/.txt`)
      }
      const { kind: sourceFormat, model, extra } = await loadModel(source, exec)
      let effective = model
      if (args.sheet && model && model.kind === 'workbook') {
        const w = normalizeWorkbook(model)
        const sh = w.sheets.find(s => s.name === args.sheet)
        if (!sh) throw new Error(`找不到工作表 "${args.sheet}"（现有: ${w.sheets.map(s => s.name).join(', ')}）`)
        effective = { kind: 'workbook', meta: w.meta, sheets: [sh] }
      }

      // ---- 出站质量门：必须在 saveBuffer 之前（G1：绝不产出目标文件）----
      const outStats = {}
      let notice
      const guard = convertSourceGuard(source, sourceFormat, extra)
      if (guard.garbled.length) {
        if (!guard.fromCache) throw new Error(convertRefusalError(source, guard, sourceFormat))
        // 有完整缓存：用 OCR 文本替换这些页后继续转换
        effective = pdfToDocument({
          meta: model?.meta || {},
          sections: extra.sections.map(s => (guard.garbled.includes(s.page)
            ? { ...s, text: String(guard.cache.get(s.page)).trim() }
            : s)),
        })
        outStats.fallback = 'ocr'
        outStats.garbledPages = pageRanges(guard.garbled)
        outStats.ocrCache = ocrCachePath(source)
        notice = `第 ${pageRanges(guard.garbled)} 页文字层不可信，已用本地 OCR 文本替代（OCR 缓存：${ocrCachePath(source)}）`
      } else {
        // 无乱码 / 非 PDF：仍对写出的正文本跑一次模型级质量门（防 docx 里塞乱码）
        const probe = typeof effective === 'string' ? effective : modelToText(effective)
        const q = textQuality(probe)
        if (q.garbled) throw new Error(convertRefusalError(source, { garbled: [], reasons: q.reasons }, sourceFormat))
      }

      const outModel = adaptModel(effective, ext)
      // 插图兜底：目标写出端不支持 image 块时显式降级（绝不静默丢图）
      const degraded = degradeImageBlocks(outModel, ext)
      const winfo = {}
      const bytes = await WRITERS[ext](degraded.model, args.encoding, winfo)
      // 阶段二：PDF 目标过产出质量门（与 office_create 同一条闸）
      const pdfQuality = ext === 'pdf' ? await pdfOutputGate(target, bytes, degraded.model, winfo) : null
      if (!pdfQuality) await saveBuffer(target, bytes)
      const out = { source, target, sourceFormat, targetFormat: ext, bytes: byteLength(bytes) }
      if (Object.keys(outStats).length || pdfQuality || degraded.degraded.length) {
        out.stats = { ...outStats }
        if (pdfQuality) out.stats.pdfQuality = pdfQuality
        if (degraded.degraded.length) out.stats.imageFallback = degraded.degraded
      }
      const notices = [notice, pdfQuality?.notice,
        degraded.degraded.length ? `${degraded.degraded.length} 张图片在 .${ext} 写出端不内嵌（已降级为文本）` : '',
      ].filter(Boolean)
      if (notices.length) out.notice = notices.join('；')
      // docx/别的写出端的图片记账（内嵌成功、尺寸换算、内容去重、读不到而退化）
      applyImageWriteInfo(out, winfo, ext)
      return out
    },
  }))
}

function parsePageSpec(spec, total) {
  const set = new Set()
  for (const chunk of String(spec).split(/[,，\s]+/)) {
    if (!chunk) continue
    const m = /^(\d+)\s*[-~]\s*(\d+)$/.exec(chunk)
    if (m) {
      const from = Math.max(1, Number(m[1]))
      const to = Math.min(total, Number(m[2]))
      for (let i = from; i <= to; i++) set.add(i)
    } else if (/^\d+$/.test(chunk)) {
      const n = Number(chunk)
      if (n >= 1 && n <= total) set.add(n)
    }
  }
  return set
}

function statsOf(model, extra) {
  const stats = { format: model?.kind || typeof model }
  try {
    if (model && model.kind === 'document') {
      const d = normalizeDocument(model)
      stats.blocks = d.blocks.length
      stats.characters = d.blocks.reduce((a, b) => a
        + String(b.text || '').length
        + ((b.runs || []).reduce((x, r) => x + String(r.text || '').length, 0))
        + ((b.rows || []).reduce((x, row) => x + row.join('').length, 0)), 0)
    } else if (model && model.kind === 'workbook') {
      const w = normalizeWorkbook(model)
      stats.sheets = w.sheets.map(s => ({ name: s.name, rows: s.rows.length, columns: Math.max(0, ...s.rows.map(r => r.length)) }))
    } else if (model && model.kind === 'slides') {
      const s = normalizeSlides(model)
      stats.slides = s.slides.length
      stats.titles = s.slides.map(x => x.title).filter(Boolean).slice(0, 20)
    } else if (typeof model === 'string') {
      stats.characters = model.length
      stats.lines = model.split('\n').length
    }
  } catch { /* stats are best-effort */ }
  if (extra?.pages) stats.pages = extra.pages
  return stats
}

function summaryOf(model) {
  if (model && model.kind === 'workbook') return normalizeWorkbook(model).sheets.map(s => `${s.name}: ${s.rows.length} 行`).join('; ')
  if (model && model.kind === 'slides') return `${normalizeSlides(model).slides.length} 页幻灯片`
  if (model && model.kind === 'document') return `${(model.blocks || []).length} 个内容块`
  return typeof model === 'string' ? `${model.length} 字符` : '空内容'
}

function markdownToWorkbookModel(md) {
  const doc = markdownToDocument(md)
  const tables = doc.blocks.filter(b => b.type === 'table')
  if (tables.length) {
    return normalizeWorkbook({
      sheets: tables.map((t, i) => ({ name: `表${i + 1}`, rows: t.rows.map(r => r.map(c => ({ v: String(c ?? ''), t: typeofVal(c) }))) })),
    })
  }
  const rows = []
  for (const b of doc.blocks) {
    if (b.type === 'heading') rows.push([b.text])
    else if (b.type === 'paragraph') rows.push([plainOf(b.runs)])
    else if (b.type === 'list') for (const it of b.items) rows.push([it.text])
  }
  return normalizeWorkbook({ sheets: [{ name: 'Sheet1', rows }] })
}

export {
  apply, inject, name, cleanVisionText, ocrPageBodies, ocrSourceLabel, runPool,
  // sanitizeTextForReturn / losslessJsonProblem / finalizeToolValue / textQuality /
  // ocrRetryScales / currentRenderScale 已就地 export
  // 读取降级链、截断协议、缓存位置、重试目录（测试用）
  readSidecarPath, maxInlineChars, capWithOffset, ocrCachePath, cacheDirState, renderDirFor,
  // DSH 补充（任务六-1）：sidecar 现状探测（失败通道可观测性）
  sidecarCoverage,
}
