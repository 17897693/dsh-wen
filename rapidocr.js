// Local OCR engine adapter: RapidOCR-json (PP-OCRv4 简中, ONNX Runtime, offline).
//
// Why this exists: the plugin's only other OCR route is the vision bridge, which
// costs one model call per page (and per retry band). For scanned handouts that
// is slow and quota-hungry, while a local engine reads a page in ~0.5s for free.
// So the read path tries this first and only escalates *dubious* pages to vision.
//
// Protocol notes (RapidOCR-json v1.1.0 / v0.2.0 release, verified on Windows):
//   * stdin is line-delimited JSON: {"image_path":"C:\\a\\b.png"} — key is
//     `image_path`; some builds print a stale "imagePath" in error text.
//   * stdout: two banner lines ("RapidOCR-json x.y.z", "OCR init completed."),
//     then exactly one result line per input image, in order:
//       {"code":100,"data":[{"box":[[x,y]*4],"score":0.98,"text":"…"}, …]}
//       {"code":101,"data":"No text found in image."}   ← data is a STRING here
//   * the process does NOT exit on EOF: it then floods {"code":299,…} forever
//     (~11k lines/s), so the caller must count result lines and hard-kill it.
//   * the engine is spawned DIRECTLY (method A, ported from the office-docs
//     skill): fs.openSync fds feed stdin/stdout, so `child.pid` is the engine
//     process and `taskkill /PID <pid> /T /F` can actually reach it. The old
//     `cmd /c … < in > out` wrapper made child.pid cmd.exe, which exits before
//     the engine starts → taskkill hit a dead pid → orphaned engine flooded
//     299s (2026-09-13: 8.47 GB of junk). Kill first, confirm the pid is gone
//     (signal-0 probe), only then delete the scratch files. stdin stays a real
//     file (LF, no BOM) and stdout/stderr go to a real file, so no named pipes
//     are needed (sandboxed hosts block pipes; fds to files always work).
//   * --numThread / OMP_* do not actually cap core usage on this build
//     (observed ~3.9 cores regardless) — hence the small per-batch page cap.
import { existsSync, mkdirSync, openSync, readSync, closeSync, writeFileSync, rmSync, rmdirSync, readdirSync, statSync } from 'node:fs'
import { spawn, spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { homedir, tmpdir } from 'node:os'
import { dirname, join } from 'node:path'

const HERE = dirname(fileURLToPath(import.meta.url))

/** Engine identity used in annotations and the OCR cache. */
export const LOCAL_ENGINE = 'rapidocr'

/** Max images per engine call; keeps a runaway batch from pegging the box. */
export const LOCAL_MAX_IMAGES = 20

const DEFAULT_MODELS = {
  det: 'ch_PP-OCRv4_det_infer.onnx',
  rec: 'rec_ch_PP-OCRv4_infer.onnx',
  cls: 'ch_ppocr_mobile_v2.0_cls_infer.onnx',
  keys: 'dict_chinese.txt',
}

/** Quality-gate thresholds: below these a page is escalated to the vision bridge. */
export const GATE = {
  minScore: Number(process.env.DSH_OFFICE_OCR_MIN_SCORE) || 0.88,
  /** many boxes *and* soft-ish confidence ⇒ dense/complex layout */
  hardBoxes: 45,
  hardBoxesScore: 0.95,
  /**
   * mostly ≤6-char fragments ⇒ the detector chopped text up ⇒ unreliable order.
   *
   * ⚠ 表格/数字页的**正常形态就是"碎"**：资料分析页由大量 ≤6 字片段组成，短句占比
   * 天然高。所以这一条只在**低置信**时判硬（`shortRatioScore` 联动），识别很准
   * （avg 高）的页绝不因"碎"被误杀——否则置信度 0.989、文本完全可读的页会被丢成
   * 空字符串。三个阈值都可用环境变量覆盖。
   *
   * 2026-09-24 第八轮：自 WB 侧同源实现回移（原实现是硬编码 0.6/30 且**无条件判硬**，
   * 既不可配、又会误杀高置信表格页）。
   */
  shortRatio: Number(process.env.DSH_OFFICE_OCR_SHORT_RATIO) || 0.6,
  shortMinBoxes: Number(process.env.DSH_OFFICE_OCR_SHORT_MIN_BOXES) || 30,
  /** 短句判据联动的置信度下限：avg 达到该值即"认得很准"，放行（不判短句硬伤）。 */
  shortRatioScore: Number(process.env.DSH_OFFICE_OCR_SHORT_SCORE) || 0.95,
}

let engineMemo

/** Candidate directories, first match wins. */
function engineCandidates() {
  const env = process.env.DSH_OFFICE_RAPIDOCR_DIR
  const list = []
  if (env) list.push(env)
  list.push(join(HERE, 'vendor'), join(HERE, 'vendor', 'RapidOCR-json_v0.2.0'))
  list.push(join(homedir(), '.dsh', 'ocr'), join(homedir(), '.dsh', 'ocr', 'RapidOCR-json_v0.2.0'))
  for (const p of String(process.env.PATH || '').split(';')) {
    if (p && p.trim()) list.push(join(p.trim(), 'RapidOCR-json_v0.2.0'), p.trim())
  }
  return list.filter(Boolean)
}

function probeDir(dir) {
  if (!dir || !existsSync(dir)) return null
  if (existsSync(join(dir, 'RapidOCR-json.exe'))) return dir
  // vendor/ may hold the release folder one level down
  try {
    for (const n of ['RapidOCR-json_v0.2.0', 'RapidOCR-json']) {
      if (existsSync(join(dir, n, 'RapidOCR-json.exe'))) return join(dir, n)
    }
  } catch { /* ignore */ }
  return null
}

/**
 * Locate a usable local engine, or null. A "usable" engine needs the exe plus
 * the default model set; anything missing means the caller should just use the
 * vision bridge instead of failing the read.
 */
export function findEngine({ refresh = false } = {}) {
  if (process.env.DSH_OFFICE_OCR_DISABLED) return null
  if (engineMemo !== undefined && !refresh) return engineMemo
  let found = null
  for (const dir of engineCandidates()) {
    const hit = probeDir(dir)
    if (!hit) continue
    const models = join(hit, 'models')
    const missing = Object.values(DEFAULT_MODELS).filter(f => !existsSync(join(models, f)))
    if (missing.length) continue
    found = { dir: hit, exe: join(hit, 'RapidOCR-json.exe'), models, version: '' }
    break
  }
  engineMemo = found
  return found
}

/** 渲染侧边长上限的基准：A4 长边在 96dpi 下的像素数（1123 = 841.89pt / 72 * 96）。 */
const SCALE_BASE_LONG_SIDE = 1123

/**
 * CLI flags for one batch run.
 *
 * `DSH_OFFICE_RENDER_SCALE` is shared with pdf-render.ps1: rendering放大多少，引擎的
 * 边长上限就必须同步放开多少，否则放大的像素会被引擎自己缩回 1024 长边，白忙一场。
 * 未设 / 非数字 / 落在 0.5–4 之外时**完全不传** --maxSideLen（引擎默认值，行为不变）。
 *
 * @param {string|undefined} rawScale 显式倍率（质量门重试走这条）；不传 = 读环境变量
 */
export function engineArgs(rawScale = process.env.DSH_OFFICE_RENDER_SCALE) {
  const models = process.env.DSH_OFFICE_OCR_MODELS || 'models'
  const pick = (k, envName) => process.env[envName] || DEFAULT_MODELS[k]
  const args = [
    `--models=${models}`,
    `--det=${pick('det', 'DSH_OFFICE_OCR_DET')}`,
    `--cls=${pick('cls', 'DSH_OFFICE_OCR_CLS')}`,
    `--rec=${pick('rec', 'DSH_OFFICE_OCR_REC')}`,
    `--keys=${pick('keys', 'DSH_OFFICE_OCR_KEYS')}`,
  ]
  if (process.env.DSH_OFFICE_OCR_NO_ANGLE) args.push('--doAngle=0', '--mostAngle=0')
  const scale = Number(rawScale)
  // 与 pdf-render.ps1 同一接受窗口（0.5–4）：两边必须一起动，窗口也必须一致，
  // 否则会出现"渲染没放大、引擎上限却放开了"这种只有一半生效的状态。
  if (rawScale !== undefined && String(rawScale).trim() !== '' && Number.isFinite(scale)
    && scale >= 0.5 && scale <= 4) {
    args.push(`--maxSideLen=${Math.max(1024, Math.round(SCALE_BASE_LONG_SIDE * scale / 256) * 256)}`)
  }
  return args
}

/** Parse one result line into a normalised page result. */
export function parseResultLine(line) {
  if (typeof line !== 'string' || !line.startsWith('{"code":')) return null
  let obj
  try { obj = JSON.parse(line) } catch { return null }
  const code = Number(obj.code)
  const boxes = Array.isArray(obj.data) ? obj.data.filter(b => b && typeof b.text === 'string') : []
  const note = typeof obj.data === 'string' ? obj.data : ''
  const text = boxesToText(boxes)
  const avg = boxes.length
    ? Math.round((boxes.reduce((s, b) => s + (Number(b.score) || 0), 0) / boxes.length) * 1000) / 1000
    : 0
  const shortRatio = boxes.length
    ? Math.round((boxes.filter(b => b.text.length <= 6).length / boxes.length) * 100) / 100
    : 1
  return { code, boxes, text, avg, shortRatio, note }
}

/** Reading order: group boxes into 24px rows by y-centre, sort each row by x. */
export function boxesToText(boxes) {
  if (!boxes.length) return ''
  const rows = new Map()
  for (const b of boxes) {
    const y = Array.isArray(b.box) && b.box[0] ? Math.floor(b.box[0][1] / 24) : 0
    if (!rows.has(y)) rows.set(y, [])
    rows.get(y).push(b)
  }
  const out = []
  for (const y of [...rows.keys()].sort((a, b) => a - b)) {
    for (const b of rows.get(y).sort((p, q) => (p.box?.[0]?.[0] ?? 0) - (q.box?.[0]?.[0] ?? 0))) {
      if (b.text) out.push(b.text)
    }
  }
  return out.join('\n')
}

/**
 * Decide whether a locally recognised page is trustworthy.
 * `blank` (code 101) is NOT escalated — the detector's own "no text found" is
 * more credible than burning a vision call on a decorative page.
 *
 * `retryable` 标记"换倍率重试有没有意义"：`短句占比高` 是版面**结构**特征，
 * 不随渲染倍率改变（放大只会让同一批碎句更清晰，占比几乎不动），换倍率纯属浪费；
 * 其余（无结果 / 置信度低 / 版面复杂）才值得换倍率救。
 */
export function gateResult(r, gate = GATE) {
  const hard = (reason, retryable = true) => ({ hard: true, reason, retryable })
  if (!r) return hard('本地引擎无结果')
  if (r.code === 101) return { hard: false, blank: true, reason: '', retryable: true }
  if (r.code !== 100) return hard(`本地引擎 code=${r.code}${r.note ? ' ' + r.note : ''}`)
  if (!r.boxes.length || !r.text.trim()) return hard('本地引擎无文字')
  if (r.avg < gate.minScore) return hard(`置信度低(${r.avg})`)
  if (r.boxes.length >= gate.hardBoxes && r.avg < gate.hardBoxesScore) {
    return hard(`版面复杂(${r.boxes.length} 框, ${r.avg})`)
  }
  // 短句占比高只在**低置信**时判硬：表格/数字页天然碎，识别准（avg 高）时放行。
  if (r.shortRatio > gate.shortRatio && r.boxes.length >= gate.shortMinBoxes && r.avg < gate.shortRatioScore) {
    return hard(`短句占比高(${r.shortRatio}, 置信 ${r.avg})`, false)
  }
  return { hard: false, blank: false, reason: '', retryable: true }
}

const POLL_MS = 120
const KILL_WAIT_MS = 12000
const SWEEP_PREFIX = 'rapid-'
/** 缓存产物（花钱生成的识别结果）绝不能被清扫碰到。 */
const PROTECTED_SUFFIX = ['.ocr.md', '.ocr.json']

const sleep = ms => new Promise(r => setTimeout(r, ms))

function isProtectedName(name) {
  const lower = String(name || '').toLowerCase()
  return PROTECTED_SUFFIX.some(s => lower.endsWith(s))
}

/**
 * Look for a protected cache artefact anywhere inside a `rapid-*` entry before
 * removing it — a directory whose *name* starts with `rapid-` must not take an
 * `.ocr.md` sitting inside it with it. Bounded so a pathological tree cannot
 * stall the sweep; returns the path found, or null.
 */
function findCacheInside(dir, maxDepth = 3, maxEntries = 2000) {
  const queue = [{ d: dir, depth: 1 }]
  let seen = 0
  while (queue.length) {
    const { d, depth } = queue.shift()
    let ents = []
    try { ents = readdirSync(d, { withFileTypes: true }) } catch { continue }
    for (const e of ents) {
      if (++seen > maxEntries) return null
      const full = join(d, e.name)
      if (e.isDirectory()) { if (depth < maxDepth) queue.push({ d: full, depth: depth + 1 }) }
      else if (isProtectedName(e.name)) return full
    }
  }
  return null
}

/**
 * Startup sweep (guard ported from office-docs): remove stale engine scratch
 * entries (basename starts with `rapid-`) older than `maxAgeHours` under the
 * OCR temp roots. Only `rapid-` prefixed entries are ever candidates; anything
 * holding (or named like) an `.ocr.md` / `.ocr.json` cache is protected.
 */
export function sweepStale({ roots = [join(tmpdir(), 'dsh-office-ocr')], maxAgeHours = 1, now = Date.now(), dryRun = false } = {}) {
  const cutoff = now - maxAgeHours * 3600e3
  const report = { roots, maxAgeHours, scanned: 0, removed: [], keptFresh: [], protected: [], errors: [] }
  for (const root of roots) {
    if (!root || !existsSync(root)) continue
    let entries = []
    try { entries = readdirSync(root, { withFileTypes: true }) } catch (e) {
      report.errors.push({ root, err: String(e.code || e) })
      continue
    }
    for (const ent of entries) {
      const full = join(root, ent.name)
      report.scanned++
      if (!ent.name.startsWith(SWEEP_PREFIX)) continue
      if (isProtectedName(ent.name)) { report.protected.push(full); continue }
      if (ent.isDirectory()) {
        const hit = findCacheInside(full)
        if (hit) { report.protected.push(hit); continue }
      }
      let st = null
      try { st = statSync(full) } catch { /* ignore */ }
      if (!st || st.mtimeMs > cutoff) { report.keptFresh.push(full); continue }
      if (dryRun) { report.removed.push({ path: full, dryRun: true, size: st.size }); continue }
      try {
        rmSync(full, { recursive: true, force: true })
        report.removed.push({ path: full, size: st.size, gone: !existsSync(full) })
      } catch (e) {
        report.errors.push({ path: full, err: String(e.code || e) })
      }
    }
  }
  return report
}

/**
 * Is the pid alive? Signal-0 probe first: it needs no external process, so it
 * also works on sandboxed hosts where tasklist is denied. EPERM is treated as
 * "alive but not signallable".
 */
function pidAlive(pid) {
  try { process.kill(pid, 0); return true } catch (e) { return e && e.code === 'EPERM' }
}

/** tasklist cross-check; null when the host denies it (sandboxed runners). */
function tasklistHasPid(pid) {
  try {
    const r = spawnSync('tasklist', ['/FI', `PID eq ${pid}`, '/NH'], { windowsHide: true, timeout: 8000 })
    if (r.error) return null
    return String(r.stdout || '').includes(String(pid))
  } catch { return null }
}

function tasklistEngineRunning() {
  try {
    const r = spawnSync('tasklist', ['/FI', 'IMAGENAME eq RapidOCR-json.exe', '/NH'], { windowsHide: true, timeout: 8000 })
    if (r.error) return null
    return String(r.stdout || '').includes('RapidOCR-json.exe')
  } catch { return null }
}

/** Wait until the pid really disappears (signal-0). Returns elapsed ms, or -1 on timeout. */
async function waitForExit(pid, timeoutMs) {
  const t0 = Date.now()
  while (Date.now() - t0 < timeoutMs) {
    if (!pidAlive(pid)) return Date.now() - t0
    await sleep(150)
  }
  return -1
}

/** taskkill one pid; stdio ignored so it also works where pipes are forbidden. */
function taskkillTree(pid) {
  try {
    return spawnSync('taskkill', ['/PID', String(pid), '/T', '/F'], { windowsHide: true, timeout: 15000, stdio: 'ignore' })
  } catch { return null }
}

/** Delete a scratch file, retrying while the killed engine's handles wind down. */
async function rmWithRetry(file, tries = 6) {
  const errs = []
  for (let i = 0; i < tries; i++) {
    try {
      rmSync(file, { force: true })
      if (!existsSync(file)) return { ok: true, tries: i + 1 }
    } catch (e) { errs.push(String((e && e.code) || e)) }
    await sleep(120 * (i + 1))
  }
  return { ok: false, tries, errs }
}

async function cleanupScratch(dir, inJ, outT, meta) {
  meta.deleted = { in: await rmWithRetry(inJ), out: await rmWithRetry(outT) }
  meta.residual = existsSync(inJ) || existsSync(outT)
  try { rmdirSync(dir); meta.workDirRemoved = true } catch { meta.workDirRemoved = false }
  try { meta.leftovers = existsSync(dir) ? readdirSync(dir).filter(nm => nm.startsWith(SWEEP_PREFIX)) : [] } catch { meta.leftovers = [] }
}

let sweptOnce = false

/**
 * Run the engine over a batch of PNGs (method A: direct exe spawn, fds for
 * stdio). Resolves to `{ pages, meta }`; `pages[i]` aligns with `pngs` and is a
 * parseResultLine() result or null when the page produced nothing (timeout /
 * crash / unreadable file). `meta` records the kill/verify/delete evidence:
 * taskkill status, pid-liveness confirmation, scratch deletion retries and any
 * residue — the caller only needs `pages`.
 */
export async function ocrImages(pngs, engine, { timeoutMs = 0, workDir, scale } = {}) {
  const n = pngs.length
  const empty = new Array(n).fill(null)
  const meta = {
    enginePid: null, images: n, ms: 0, bytes: 0, stopReason: 'not-started',
    kill: null, deleted: null, residual: false, workDirRemoved: false,
    leftovers: [], engineStillRunning: null, swept: null,
  }
  if (!engine || !n) { meta.stopReason = 'no-engine-or-images'; return { pages: empty, meta } }

  // 启动清扫：上次异常退出留下的 rapid-* 残留（>1 小时）在跑引擎前清掉；只做一次/进程
  if (!sweptOnce) {
    sweptOnce = true
    try { meta.swept = sweepStale() } catch { meta.swept = null }
  }

  const dir = workDir || join(tmpdir(), 'dsh-office-ocr', `rapid-${Date.now()}`)
  try { mkdirSync(dir, { recursive: true }) } catch { /* temp may be read-only */ }
  const stamp = `${process.pid}-${Date.now()}`
  const inJ = join(dir, `rapid-in-${stamp}.json`)
  const outT = join(dir, `rapid-out-${stamp}.txt`)
  const out = new Array(n).fill(null)
  try {
    const body = pngs.map(p => `{"image_path":"${String(p).replace(/\\/g, '\\\\')}"}`).join('\n') + '\n'
    writeFileSync(inJ, body, { encoding: 'utf8' }) // writeFileSync emits no BOM for utf8 strings
  } catch {
    meta.stopReason = 'write-in-failed'
    return { pages: out, meta }
  }

  const flags = engineArgs(scale)
  const started = Date.now()
  const budget = timeoutMs || (30000 + 15000 * n)
  let child
  try {
    const fdIn = openSync(inJ, 'r')
    const fdOut = openSync(outT, 'w')
    try {
      // stdio are file descriptors (NOT pipes): sandboxed hosts that block named
      // pipes accept fds, and child.pid is the engine itself.
      child = spawn(engine.exe, flags, { cwd: engine.dir, stdio: [fdIn, fdOut, fdOut], windowsHide: true })
    } finally {
      // parent's copies are dead weight once the child inherited them; close
      // immediately so they never keep the scratch files locked
      try { closeSync(fdIn) } catch { /* ignore */ }
      try { closeSync(fdOut) } catch { /* ignore */ }
    }
  } catch (e) {
    meta.stopReason = 'spawn-failed: ' + (e && e.message)
    await cleanupScratch(dir, inJ, outT, meta)
    return { pages: out, meta }
  }
  meta.enginePid = child.pid

  let readFd = -1
  let pos = 0
  let carry = ''
  let bytes = 0
  let stopReason = 'unknown'
  const lines = []
  try {
    while (true) {
      if (Date.now() - started > budget) { stopReason = 'budget'; break }
      if (readFd < 0) {
        if (!existsSync(outT)) { await sleep(POLL_MS); continue }
        try { readFd = openSync(outT, 'r') } catch { await sleep(POLL_MS); continue }
      }
      const buf = Buffer.alloc(1 << 20)
      let read = 0
      try { read = readSync(readFd, buf, 0, buf.length, pos) } catch { await sleep(POLL_MS); continue }
      if (read > 0) {
        pos += read
        bytes += read
        const chunk = carry + buf.toString('utf8', 0, read)
        const parts = chunk.split('\n')
        carry = parts.pop() ?? ''
        for (const raw of parts) {
          const line = raw.replace(/\r$/, '')
          if (!line.startsWith('{"code":')) continue       // banner / stray log
          if (line.includes('"code":299')) continue         // post-EOF spew
          lines.push(line)
        }
      }
      if (lines.length >= n) { stopReason = 'complete'; break }
      // no growth AND the engine process itself exited: it died early
      if (read === 0 && child.exitCode !== null) { stopReason = 'child-exit'; break }
      if (bytes > 96 << 20) { stopReason = 'byte-cap'; break }   // runaway guard
      await sleep(POLL_MS)
    }
    meta.ocrDoneMs = Date.now() - started   // pure engine+detect time, before kill/verify/cleanup
  } finally {
    if (readFd >= 0) { try { closeSync(readFd) } catch { /* ignore */ } }
    meta.bytes = bytes

    // --- kill the engine itself (the whole point of method A) ---
    const tk = taskkillTree(child.pid)
    try { child.kill('SIGKILL') } catch { /* already gone */ }
    let waited = await waitForExit(child.pid, KILL_WAIT_MS)
    if (waited < 0) {
      taskkillTree(child.pid)                                // one defensive retry
      waited = await waitForExit(child.pid, KILL_WAIT_MS)
    }
    meta.kill = {
      pid: child.pid,
      taskkillStatus: tk ? tk.status : null,
      goneAfterMs: waited,
      gone: waited >= 0,
      // independent confirmation when the host allows tasklist; null = denied
      tasklistConfirm: waited >= 0 ? (tasklistHasPid(child.pid) === false ? 'tasklist-gone' : null) : 'timeout',
    }

    // only after the pid is confirmed gone do we touch the scratch files
    await cleanupScratch(dir, inJ, outT, meta)
    meta.engineStillRunning = tasklistEngineRunning()
  }

  for (let i = 0; i < lines.length && i < n; i++) out[i] = parseResultLine(lines[i])
  meta.ms = Date.now() - started
  meta.stopReason = stopReason
  meta.linesSeen = lines.length
  return { pages: out, meta }
}

/** One-line description for notes and the cache header. */
export function engineLabel(engine = findEngine()) {
  if (!engine) return ''
  try {
    const v = statSync(join(engine.dir, 'RapidOCR-json.exe'))
    return `RapidOCR-json · ${process.env.DSH_OFFICE_OCR_REC || DEFAULT_MODELS.rec} (${(v.size / 1048576).toFixed(0)}MB)`
  } catch {
    return `RapidOCR-json · ${process.env.DSH_OFFICE_OCR_REC || DEFAULT_MODELS.rec}`
  }
}
