// repro.mjs — office_read 返回值"无损 JSON"最小复现 + 常驻哨兵
//
// 症状（另一会话已复现，本脚本把它固化成可重跑的一条命令）：
//   任何含正文的 office_read 对某些 PDF 返回
//     tool "office_read" returned invalid output: value is not lossless JSON
//   正文全部丢失，模型只看到这一句；`as="meta"` 因为不含正文反而正常。
//
// 本脚本做三件事：
//   1) 复现：走真实提取链路读样本 PDF 第 1 页，对返回的**字符串与结构**做四查
//      ① 游离代理（U+D800–U+DBFF 无配对 / U+DC00–U+DFFF 孤儿）
//      ② 所有码点可 UTF-8 往返（不能变成 `?` / U+FFFD）
//      ③ 异常控制字符（U+0000–U+001F 除 \t\n\r、U+007F、U+2028/2029）
//      ④ NFC 规范化
//      外加 harness 真正拒收的那几类结构问题（undefined 属性 / 非有限数 /
//      稀疏数组 / 非纯对象 / 循环引用）——实测这才是"整条拒收"的元凶。
//   2) 判定坏字符来自**提取层**还是**组装层**（打印各自计数）。
//   3) 修复后转绿，并被 test.mjs 复用防回归。
//
// 用法：
//   node repro.mjs                         # 自动找样本（见 pickSample）
//   node repro.mjs --pages 1-43
//   DSH_OFFICE_REPRO_PDF=D:/x.pdf node repro.mjs
//   DSH_OFFICE_REPRO_KEEP=1 node repro.mjs  # 保留下载/拷贝出来的临时样本
//
// 退出码：0 = 绿（返回值 harness 一定收）；1 = 红（会被拒收）；2 = 脚本自身出错。

import { apply } from './index.js'
import { mkdir, rm, copyFile, readFile } from 'node:fs/promises'
import { existsSync, readdirSync, statSync } from 'node:fs'
import { join, dirname, resolve, basename, extname } from 'node:path'
import { tmpdir } from 'node:os'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { readPdfText } from './pdf.js'

const HERE = dirname(fileURLToPath(import.meta.url))
const argv = process.argv.slice(2)
const argOf = (k, d) => {
  const i = argv.indexOf(`--${k}`)
  if (i >= 0 && i + 1 < argv.length) return argv[i + 1]
  const eq = argv.find(a => a.startsWith(`--${k}=`))
  return eq ? eq.slice(k.length + 3) : d
}

// ---------------------------------------------------------------------------
// 一、四查（纯函数，可被 test.mjs 直接 import）
// ---------------------------------------------------------------------------

/** 唯一允许原样保留的控制符：制表、换行、回车。 */
const KEEP_CONTROL = new Set([0x09, 0x0a, 0x0d])

/**
 * 扫一个字符串里的坏码点。位置用**码元下标**（与 String#slice 一致），
 * 方便直接回原文定位。
 *
 * @param {string} s
 * @returns {{loneSurrogates: Array, noncharacters: Array, controls: Array, separators: Array, bad: number, warnings: number}}
 */
export function scanBadCodePoints(s) {
  const str = typeof s === 'string' ? s : String(s ?? '')
  const out = { loneSurrogates: [], noncharacters: [], controls: [], separators: [], bad: 0, warnings: 0 }
  for (let i = 0; i < str.length; i++) {
    const cp = str.codePointAt(i)
    if (cp > 0xffff) { i++; continue }              // 合法配对 → 补充平面字符，正常
    if (cp >= 0xd800 && cp <= 0xdfff) {
      out.loneSurrogates.push({ at: i, cp }); out.bad++
    } else if (cp === 0xfffe || cp === 0xffff) {
      out.noncharacters.push({ at: i, cp }); out.bad++
    } else if ((cp <= 0x1f && !KEEP_CONTROL.has(cp)) || (cp >= 0x7f && cp <= 0x9f)) {
      out.controls.push({ at: i, cp }); out.bad++
    } else if (cp === 0x2028 || cp === 0x2029) {
      // 合同上算"异常"，实现上刻意保留（Word/PPT 真的在用，属"宁留噪不删正文"）
      out.separators.push({ at: i, cp }); out.warnings++
    }
  }
  return out
}

/** ② 所有码点可 UTF-8 往返：往返不等 = 有码点会被编码器静默替换。 */
export function utf8RoundTrips(s) {
  const str = typeof s === 'string' ? s : String(s ?? '')
  try { return Buffer.from(str, 'utf8').toString('utf8') === str } catch { return false }
}

/** ④ NFC 检查。 */
export function isNfc(s) {
  return String(s ?? '') === String(s ?? '').normalize('NFC')
}

const fmtCp = cp => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`

/**
 * ③+①+②+④ 汇总成一段人话，`bad===0` 时返回 null。
 */
export function describeStringProblems(s, where = '') {
  const bad = scanBadCodePoints(s)
  const zero = bad.loneSurrogates.length + bad.noncharacters.length + bad.controls.length
  if (zero === 0 && utf8RoundTrips(s) && isNfc(s)) return null
  const bits = []
  if (bad.loneSurrogates.length) bits.push(`${fmtCp(bad.loneSurrogates[0].cp)}@${bad.loneSurrogates[0].at}`)
  if (bad.noncharacters.length) bits.push(`${fmtCp(bad.noncharacters[0].cp)}@${bad.noncharacters[0].at}`)
  if (bad.controls.length) bits.push(`${fmtCp(bad.controls[0].cp)}@${bad.controls[0].at}`)
  if (!utf8RoundTrips(s)) bits.push('UTF-8 往返不等')
  if (!isNfc(s)) bits.push('非 NFC')
  const detail = [
    bad.loneSurrogates.length ? `游离代理 ${bad.loneSurrogates.length}` : '',
    bad.noncharacters.length ? `非字符 ${bad.noncharacters.length}` : '',
    bad.controls.length ? `控制符 ${bad.controls.length}` : '',
  ].filter(Boolean).join(' / ')
  return `${where}${where ? '：' : ''}${detail || '编码问题'}（首处 ${bits.join(' ')}）`
}

// ---------------------------------------------------------------------------
// 二、harness 侧的"无损 JSON"判定（照 snapshotJsonValue 的语义抄一份）
// ---------------------------------------------------------------------------

/**
 * 语义对齐 harness 的 `snapshotJsonValue`。返回第一处违规的**路径描述**，
 * 合法时返回 null。这不是"猜"：`undefined` 属性、非有限数、稀疏数组、
 * 非纯对象、循环引用都是实测会被整条拒收的东西。
 */
export function losslessJsonProblem(value, path = '$', seen = new Set()) {
  if (value === null) return null
  const t = typeof value
  if (t === 'string' || t === 'boolean') return null
  if (t === 'number') return Number.isFinite(value) && !Object.is(value, -0) ? null : `${path} 不是无损 JSON 数字（${String(value)}）`
  if (t === 'undefined') return `${path} 是 undefined（harness 会整条拒收）`
  if (t === 'bigint') return `${path} 是 bigint`
  if (t === 'function' || t === 'symbol') return `${path} 是 ${t}`
  if (t !== 'object') return `${path} 类型 ${t} 不可序列化`
  if (seen.has(value)) return `${path} 存在循环引用`
  seen.add(value)
  try {
    if (Array.isArray(value)) {
      if (Reflect.ownKeys(value).length !== value.length + 1) return `${path} 是稀疏数组或带额外自有键`
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

/** 收集一个值里所有字符串（含键）以及坏码点统计。 */
export function collectStrings(value, path = '$', acc = []) {
  if (typeof value === 'string') { acc.push({ path, value }); return acc }
  if (value === null || typeof value !== 'object') return acc
  if (Array.isArray(value)) {
    value.forEach((v, i) => collectStrings(v, `${path}[${i}]`, acc))
    return acc
  }
  for (const k of Object.keys(value)) collectStrings(value[k], `${path}.${k}`, acc)
  return acc
}

// ---------------------------------------------------------------------------
// 三、样本定位
// ---------------------------------------------------------------------------

/** 附件目录：`~/.dsh/attachments` 下按 mtime 找最新的 .pdf（可被 env 覆盖）。 */
function pickSample() {
  const env = process.env.DSH_OFFICE_REPRO_PDF
  if (env) return env
  const root = join(process.env.USERPROFILE || process.env.HOME || '.', '.dsh', 'attachments')
  const found = []
  const walk = (dir, depth) => {
    if (depth > 6) return
    let entries = []
    try { entries = readdirSync(dir, { withFileTypes: true }) } catch { return }
    for (const e of entries) {
      const p = join(dir, e.name)
      if (e.isDirectory()) walk(p, depth + 1)
      else if (extname(e.name).toLowerCase() === '.pdf') {
        try { found.push({ p, m: statSync(p).mtimeMs, s: statSync(p).size }) } catch { /* ignore */ }
      }
    }
  }
  walk(root, 0)
  // 备考资料多为几 MB 的中文 PDF：优先大文件（扫描/多页），其次最新
  found.sort((a, b) => (b.s - a.s) || (b.m - a.m))
  return found[0]?.p || null
}

// ---------------------------------------------------------------------------
// 四、跑
// ---------------------------------------------------------------------------

const TMP = join(tmpdir(), 'dsh-office-repro')

async function main() {
  const sample = pickSample()
  if (!sample) {
    console.error('找不到样本 PDF：用 DSH_OFFICE_REPRO_PDF=<路径> 指定一个中文 PDF')
    process.exitCode = 2
    return
  }

  // 附件目录只读且旁边可能有真实 .ocr.md —— 一律拷到临时目录里操作
  await mkdir(TMP, { recursive: true })
  const local = join(TMP, basename(sample))
  await copyFile(sample, local)

  const captured = new Map()
  apply({ get: n => (n === 'tools' ? { register: d => { captured.set(d.name, d) } } : undefined) })
  const exec = { agent: { session: { header: { cwd: process.cwd() } } } }

  const pages = argOf('pages', '1')
  console.log(`样本: ${sample}`)
  console.log(`工作副本: ${local}`)
  console.log(`调用: office_read pages="${pages}"\n`)

  // ---- 4.1 组装层（工具返回值） ----
  const raw = await captured.get('office_read').execute({ path: local, pages }, exec)

  const structural = losslessJsonProblem(raw)
  const strings = collectStrings(raw)
  const stringProblems = strings.map(s => ({ ...s, problem: describeStringProblems(s.value, s.path) })).filter(s => s.problem)
  const totalBad = strings.reduce((a, s) => a + scanBadCodePoints(s.value).bad, 0)
  const totalWarn = strings.reduce((a, s) => a + scanBadCodePoints(s.value).warnings, 0)
  const notNfc = strings.filter(s => !isNfc(s.value)).length
  const badUtf8 = strings.filter(s => !utf8RoundTrips(s.value)).length

  // ---- 4.2 提取层（原始文本，绕开组装） ----
  let extractBad = 0
  let extractNfc = 0
  let extractNote = ''
  let preview = ''
  try {
    const doc = readPdfText(await readFile(local))          // 注意：吃 buffer，不吃路径
    const all = doc.sections.map(s => String(s.text || '')).join('\n')
    for (const t of doc.sections) extractBad += scanBadCodePoints(String(t.text || '')).bad
    extractNfc = doc.sections.filter(t => !isNfc(String(t.text || ''))).length
    extractNote = `${doc.sections.length} 页 / ${all.length} 字符`
    preview = String(doc.sections[0]?.text || '').slice(0, 90).replace(/\n/g, '⏎')
  } catch (e) {
    extractNote = `直读失败：${e.message}`
  }

  // ---- 4.3 报告 ----
  console.log('【组装层】office_read 返回值')
  console.log(`  结构无损性        : ${structural ? `❌ ${structural}` : '✅ 通过'}`)
  console.log(`  字符串坏码点总数  : ${totalBad}`)
  console.log(`  UTF-8 往返失败串  : ${badUtf8}`)
  console.log(`  非 NFC 串         : ${notNfc}`)
  if (totalWarn) console.log(`  U+2028/2029（保留）: ${totalWarn}`)
  for (const s of stringProblems.slice(0, 12)) console.log(`  ❌ ${s.problem}`)
  if (stringProblems.length > 12) console.log(`  …另有 ${stringProblems.length - 12} 处`)

  console.log('\n【提取层】readPdfText 原始文本')
  console.log(`  ${extractNote}`)
  console.log(`  坏码点总数        : ${extractBad}`)
  console.log(`  非 NFC 页数       : ${extractNfc}`)
  if (preview) console.log(`  第 1 页开头       : ${preview}`)

  console.log('\n【判定】坏字符来源')
  console.log(`  ${extractBad > 0
    ? `提取层：readPdfText 原始文本已含 ${extractBad} 个坏码点（CID 字体解码产物）`
    : '提取层：raw 文本层干净（0 个坏码点）'}`)
  console.log(`  ${structural
    ? `组装层：返回值结构已违规 —— ${structural}`
    : '组装层：结构合法'}`)

  const green = !structural && totalBad === 0 && badUtf8 === 0 && notNfc === 0
  console.log(`\n${green ? '🟢 GREEN' : '🔴 RED'}  ${green ? '返回值可无损序列化，harness 一定收' : '返回值会被 harness 整条拒收（value is not lossless JSON）'}`)
  console.log(`复现命令：node repro.mjs --pages "${pages}"`)

  if (!process.env.DSH_OFFICE_REPRO_KEEP) await rm(TMP, { recursive: true, force: true })
  process.exitCode = green ? 0 : 1
}

// 只有直接执行才跑 main()；被 test.mjs import 时只借用上面的纯函数。
const DIRECT = (() => {
  try { return process.argv[1] ? import.meta.url === pathToFileURL(resolve(process.argv[1])).href : false } catch { return false }
})()

if (DIRECT) {
  main().catch(e => {
    console.error('HARNESS ERROR:', e)
    process.exitCode = 2
  })
}
