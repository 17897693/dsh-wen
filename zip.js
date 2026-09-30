// Minimal, dependency-free ZIP reader/writer sufficient for OOXML/ODF
// packages (deflate + stored entries, UTF-8 names, zip64 EOCD tolerated).
//
// 任务三（R16）：**资源限制与结构校验**。旧读取器只解压、不设防，实测后果：
//   · 33 KB 的"docx 形态"文件能让 office_read 吃进 +129 MiB，261 KB 的能吃 +512 MiB；
//   · 分散形态的炸弹甚至**不报错、静默成功**（只有集中形态才会撞上 V8 的栈上限）；
//   · 越界 / 说谎的字段一律 `RangeError: Offset is outside the bounds of the DataView`，
//     而不是可读错误；stored 条目还能静默返回「负载 + 中央目录 + EOCD」的拼接物。
// 现在：中央目录 / local header / 偏移 / 长度 / ZIP64 全部做边界校验（越界一律可读中文错误），
// 解压前按**声明尺寸**挡两道上限、解压时用 zlib 的 `maxOutputLength` 早停、解压后核对实际长度，
// 并按需校验 CRC32。上限是**惰性**的：只读 `[Content_Types].xml` 或只看 `names` 的调用不受影响。
import { deflateRawSync, inflateRawSync } from 'node:zlib'

const CRC_TABLE = /* @__PURE____ */ (() => {
  const table = new Int32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[i] = c
  }
  return table
})()

function crc32(buf) {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

const enc = new TextEncoder()
const decUtf8 = new TextDecoder('utf-8', { fatal: false })
const decUtf8Strict = new TextDecoder('utf-8', { fatal: true })

// WHATWG TextDecoder has no cp437 label, so the OEM code page used by legacy
// zip writers is decoded locally (0x00-0x7f is ASCII, the rest is the table).
const CP437_HIGH = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u00a0'

function decodeCp437(bytes) {
  let out = ''
  for (const b of bytes) out += b < 0x80 ? String.fromCharCode(b) : CP437_HIGH[b - 0x80] ?? '?'
  return out
}

function decodeName(bytes, utf8Flag) {
  if (utf8Flag) return decUtf8.decode(bytes)
  // Heuristic: try UTF-8 strictly first (many writers omit the flag), else cp437.
  try { return decUtf8Strict.decode(bytes) } catch { return decodeCp437(bytes) }
}

// ---------------------------------------------------------------------------
// 体积护栏（可环境变量覆盖，0 = 关闭该上限；风格对齐 DSH_OFFICE_MAX_INLINE_CHARS）
// ---------------------------------------------------------------------------
//
// 阈值依据（r16 实测，见 work/r16-run/probe-zip/AUDIT-zip.md）：
//   · 合法大型 Office 文档：docx 内嵌 2000×2000 不可压缩 PNG → 单条目 11.45 MiB（stored，
//     压缩比 1.00×，说明**压缩比对图片完全无效**）；xlsx 20000 行×8 列 → `sheet1.xml`
//     17.31 MiB（放大 27.8×，≈908 B/行）。
//   · 输入侧 `MAX_FILE_BYTES` 是 200 MiB，但实测放大比可达 294×（合法 XML）~1026×（零串），
//     只能用"解压后绝对体积"兜底，**不能按压缩比拒绝**（那会误杀正常的大图/大表文档）。
//   · 单条目 256 MiB ＝ 实测最大合法条目的 14.8× / 最大图片条目的 22×，按 908 B/行覆盖约
//     29.6 万行工作表；Excel 理论上限的 950 MiB 大表刻意不覆盖 —— 那种包解压后近 1 GiB，
//     模型层必然先耗尽内存，不如给一条点名环境变量的可读错误。
//   · 累计 512 MiB ＝ 合法整包实测最大值（17.31 MiB）的 ~30×；`office_edit` 的 zipToMap 会把
//     每个条目都解压、再全部重新 deflate，累计才是编辑链路的真实约束点。
const MIB = 1024 * 1024

/** 单条目解压上限，默认 256 MiB。`DSH_OFFICE_ZIP_MAX_ENTRY_BYTES=0` 关闭该上限。 */
export function maxEntryBytes(raw = process.env.DSH_OFFICE_ZIP_MAX_ENTRY_BYTES) {
  return envBytes(raw, 256 * MIB)
}

/** 整归档累计解压上限，默认 512 MiB。`DSH_OFFICE_ZIP_MAX_TOTAL_BYTES=0` 关闭该上限。 */
export function maxTotalBytes(raw = process.env.DSH_OFFICE_ZIP_MAX_TOTAL_BYTES) {
  return envBytes(raw, 512 * MIB)
}

function envBytes(raw, dflt) {
  if (raw === undefined || raw === null || String(raw).trim() === '') return dflt
  const n = Number(raw)
  if (n === 0) return Infinity
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : dflt
}

const fmtBytes = n => (n === Infinity ? '不限' : n >= MIB ? `${(n / MIB).toFixed(n % MIB ? 1 : 0)} MiB` : `${n} 字节`)

// ---------------------------------------------------------------------------
// 有界读取器：任何越界都变成可读中文错误，绝不 RangeError
// ---------------------------------------------------------------------------
class ZipView {
  constructor(buf) {
    this.buf = buf
    this.dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
    this.len = buf.byteLength
  }
  need(off, size, what) {
    if (!Number.isSafeInteger(off) || off < 0 || !Number.isSafeInteger(size) || size < 0
      || !Number.isSafeInteger(off + size) || off + size > this.len) {
      throw new Error(`zip 结构损坏：${what} 越界（偏移 ${off} + ${size} > 文件长度 ${this.len}）`)
    }
  }
  u16(off, what) { this.need(off, 2, what); return this.dv.getUint16(off, true) }
  u32(off, what) { this.need(off, 4, what); return this.dv.getUint32(off, true) }
  u64(off, what) {
    this.need(off, 8, what)
    const v = this.dv.getBigUint64(off, true)
    if (v > 0x7fffffffffffffffn) throw new Error(`zip 结构损坏：${what} 数值过大（0x${v.toString(16)}）`)
    return Number(v)
  }
  sigAt(off, want) {
    if (off < 0 || off + 4 > this.len) return false
    return this.dv.getUint32(off, true) === want
  }
  bytes(off, size, what) { this.need(off, size, what); return this.buf.subarray(off, off + size) }
}

const SIG_LOCAL = 0x04034b50
const SIG_CD = 0x02014b50
const SIG_EOCD = 0x06054b50
const SIG_Z64_LOCATOR = 0x07064b50
const SIG_Z64_EOCD = 0x06064b50

/**
 * 定位 EOCD：**先严格后宽容**两轮。
 * 严格轮要求"注释长度与文件末尾自洽"（`i + 22 + commentLen === len`），这样注释里混进的
 * 伪 EOCD 签名会被跳过（旧实现在这种文件上直接 RangeError）；一条都不满足才回落到旧行为
 * （EOCD 之后还有附加字节的自解压包）。反扫起点保持 `len - 22`：它保证 `eocd + 22 <= len`。
 */
function findEocd(z) {
  const min = Math.max(0, z.len - 22 - 65536)
  let lenient = -1
  for (let i = z.len - 22; i >= min; i--) {
    if (!(z.buf[i] === 0x50 && z.buf[i + 1] === 0x4b && z.buf[i + 2] === 0x05 && z.buf[i + 3] === 0x06)) continue
    if (lenient < 0) lenient = i
    const commentLen = z.buf[i + 20] | (z.buf[i + 21] << 8)
    if (i + 22 + commentLen === z.len) return i
  }
  if (lenient >= 0) return lenient
  throw new Error('zip 结构损坏：未找到中央目录结尾记录（EOCD）；文件不是 zip 或已被截断')
}

/** 解析 ZIP64 扩展字段（ID 0x0001）；字段不足、声明越界、记录缺失都抛可读错误。 */
function readZip64Extra(z, off, len, want) {
  let q = 0
  while (q + 4 <= len) {
    const id = z.u16(off + q, '中央目录条目扩展字段 ID')
    const sz = z.u16(off + q + 2, '中央目录条目扩展字段长度')
    if (q + 4 + sz > len) {
      throw new Error(`zip 结构损坏：扩展字段（ID 0x${id.toString(16)}，声明 ${sz} 字节）超出中央目录条目范围（可用 ${len - q - 4} 字节）`)
    }
    if (id === 0x0001) {
      const end = off + q + 4 + sz
      let r = off + q + 4
      const out = {}
      const take = what => {
        if (r + 8 > end) throw new Error(`zip 结构损坏：ZIP64 扩展字段不足，缺少「${what}」（需要 8 字节，仅剩 ${end - r} 字节）`)
        const v = z.u64(r, `ZIP64 ${what}`)
        r += 8
        return v
      }
      // APPNOTE 规定这三个字段只在对应值取哨兵时出现，顺序固定
      if (want.uncSize) out.uncSize = take('未压缩长度')
      if (want.compSize) out.compSize = take('压缩长度')
      if (want.localOff) out.localOff = take('本地头偏移')
      return out
    }
    q += 4 + sz
  }
  throw new Error('zip 结构损坏：条目使用了 ZIP64 哨兵尺寸（0xffffffff），但扩展字段里没有 ZIP64（ID 0x0001）记录')
}

/**
 * Read a zip archive from a Uint8Array. Returns { names, get(name), getText(name) }.
 *
 * @param {Uint8Array} buf
 * @param {{maxEntryBytes?:number, maxTotalBytes?:number, checkCrc?:boolean}} [opts] 测试用覆盖
 */
export function openZip(buf, opts = {}) {
  const z = new ZipView(buf)
  const eocd = findEocd(z)
  z.need(eocd, 22, 'EOCD 记录')
  let cdCount = z.u16(eocd + 10, 'EOCD 条目总数')
  let cdSize = z.u32(eocd + 12, 'EOCD 中央目录大小')
  let cdOffset = z.u32(eocd + 16, 'EOCD 中央目录偏移')
  const commentLen = z.u16(eocd + 20, 'EOCD 注释长度')
  if (eocd + 22 + commentLen > z.len) {
    throw new Error(`zip 结构损坏：EOCD 声明注释 ${commentLen} 字节，但文件只剩 ${z.len - eocd - 22} 字节`)
  }
  // ZIP64：只要出现哨兵值就必须有定位记录 + EOCD64。
  // 旧实现把条件写成 `(…) && cdCount === 0`，于是**真实 ZIP64 永远进不来** → cdOffset 保持
  // 0xffffffff → RangeError（文件头注释里"zip64 EOCD tolerated"对真 ZIP64 是假的）。
  const sentinel = cdCount === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff
  const loc = eocd - 20
  const hasLocator = loc >= 0 && z.sigAt(loc, SIG_Z64_LOCATOR)
  if (hasLocator && (sentinel || cdCount === 0)) {
    const z64off = z.u64(loc + 8, 'ZIP64 定位记录指向的 EOCD64 偏移')
    z.need(z64off, 56, 'ZIP64 EOCD64 记录')
    if (!z.sigAt(z64off, SIG_Z64_EOCD)) {
      throw new Error(`zip 结构损坏：ZIP64 定位记录指向的偏移 ${z64off} 不是 EOCD64 记录`)
    }
    const declaredLen = z.u64(z64off + 4, 'ZIP64 EOCD64 记录长度')
    if (declaredLen < 44) throw new Error(`zip 结构损坏：ZIP64 EOCD64 记录长度 ${declaredLen} 小于最小值 44`)
    cdCount = z.u64(z64off + 32, 'ZIP64 条目总数')
    cdSize = z.u64(z64off + 40, 'ZIP64 中央目录大小')
    cdOffset = z.u64(z64off + 48, 'ZIP64 中央目录偏移')
  } else if (cdOffset === 0xffffffff || cdSize === 0xffffffff) {
    throw new Error('zip 结构损坏：EOCD 使用了 ZIP64 哨兵值，但缺少 ZIP64 定位记录')
  }
  // —— 中央目录与条目数必须真的落在文件里（旧实现两样都不查）——
  z.need(cdOffset, 0, '中央目录偏移')
  if (cdSize > 0 && cdOffset + cdSize > z.len) {
    throw new Error(`zip 结构损坏：中央目录偏移 ${cdOffset} + 大小 ${cdSize} 超出文件长度 ${z.len}`)
  }
  // 每个条目头部至少 46 字节 → 这是**严格成立**的上界，绝不会误拒正常归档
  if (cdCount > Math.floor(z.len / 46) + 1) {
    throw new Error(`zip 结构损坏：EOCD 声称 ${cdCount} 个条目，但文件仅 ${z.len} 字节（每个条目头部至少 46 字节，最多 ${Math.floor(z.len / 46)} 个）`)
  }
  // 0 = 关闭该上限（与 `DSH_OFFICE_ZIP_*=0` 同一语义，opts 覆盖也遵循）
  const normCap = v => (v === 0 ? Infinity : v)
  const maxEntry = normCap(opts.maxEntryBytes) ?? maxEntryBytes()
  const maxTotal = normCap(opts.maxTotalBytes) ?? maxTotalBytes()
  const checkCrc = opts.checkCrc ?? process.env.DSH_OFFICE_ZIP_CRC !== '0'

  const entries = new Map()
  let p = cdOffset
  for (let n = 0; n < cdCount; n++) {
    z.need(p, 46, `第 ${n + 1} 个中央目录条目头`)
    if (!z.sigAt(p, SIG_CD)) {
      throw new Error(`zip 结构损坏：第 ${n + 1} 个中央目录条目签名不是 0x02014b50（偏移 ${p}）`)
    }
    const flags = z.u16(p + 8, '条目通用标志')
    const method = z.u16(p + 10, '压缩方法')
    let compSize = z.u32(p + 20, '压缩长度')
    let uncSize = z.u32(p + 24, '未压缩长度')
    const crc = z.u32(p + 16, 'CRC32')
    const nameLen = z.u16(p + 28, '条目名长度')
    const extraLen = z.u16(p + 30, '扩展字段长度')
    const entryCommentLen = z.u16(p + 32, '条目注释长度')
    let localOff = z.u32(p + 42, '本地头偏移')
    // 变长区必须整段在文件内：旧实现用 subarray 静默截断，名字会变成垃圾
    z.need(p + 46, nameLen + extraLen + entryCommentLen, `第 ${n + 1} 个条目的变长区`)
    const nameBytes = z.bytes(p + 46, nameLen, `第 ${n + 1} 个条目的名字`)
    const name = decodeName(nameBytes, (flags & 0x800) !== 0)
    if (!name) throw new Error(`zip 结构损坏：第 ${n + 1} 个条目名字为空`)
    if (method !== 0 && method !== 8) throw new Error(`不支持的 zip 压缩方法 ${method}（条目 ${name}）`)
    if (compSize === 0xffffffff || uncSize === 0xffffffff || localOff === 0xffffffff) {
      const x = readZip64Extra(z, p + 46 + nameLen, extraLen, {
        uncSize: uncSize === 0xffffffff,
        compSize: compSize === 0xffffffff,
        localOff: localOff === 0xffffffff,
      })
      if (x.uncSize !== undefined) uncSize = x.uncSize
      if (x.compSize !== undefined) compSize = x.compSize
      // 旧实现在这里只写了注释 `/* localOff = */` —— 赋值丢了，ZIP64 本地头偏移的包必崩
      if (x.localOff !== undefined) localOff = x.localOff
    }
    if (entries.has(name)) throw new Error(`zip 结构损坏：条目名 "${name}" 重复出现`)
    entries.set(name, { name, method, compSize, uncSize, crc, flags, localOff })
    p += 46 + nameLen + extraLen + entryCommentLen
  }

  const dataCache = new Map()
  let totalOut = 0
  function read(name) {
    if (dataCache.has(name)) return dataCache.get(name)
    const e = entries.get(name)
    if (!e) return undefined
    // —— 解压前：先按中央目录的**声明尺寸**挡两道上限，绝不"先解压再检查" ——
    if (e.compSize > z.len) {
      throw new Error(`zip 结构损坏：条目 "${e.name}" 声明压缩长度 ${e.compSize} 超过文件长度 ${z.len}`)
    }
    if (e.uncSize > maxEntry) {
      throw new Error(`zip 条目解压后过大：${e.name} 声明 ${fmtBytes(e.uncSize)}，单条目上限 ${fmtBytes(maxEntry)}（可用 DSH_OFFICE_ZIP_MAX_ENTRY_BYTES 调整）`)
    }
    if (totalOut + e.uncSize > maxTotal) {
      throw new Error(`zip 累计解压体积超限：已解压 ${fmtBytes(totalOut)}，再加 ${e.name} 的 ${fmtBytes(e.uncSize)} 将超过全归档上限 ${fmtBytes(maxTotal)}（可用 DSH_OFFICE_ZIP_MAX_TOTAL_BYTES 调整）`)
    }
    const lh = e.localOff
    z.need(lh, 30, `条目 "${e.name}" 的本地头`)
    if (!z.sigAt(lh, SIG_LOCAL)) {
      throw new Error(`zip 结构损坏：条目 "${e.name}" 的本地头签名不对（偏移 ${lh}）`)
    }
    const lhNameLen = z.u16(lh + 26, `条目 "${e.name}" 本地头的名字长度`)
    const lhExtraLen = z.u16(lh + 28, `条目 "${e.name}" 本地头的扩展字段长度`)
    const start = lh + 30 + lhNameLen + lhExtraLen
    const raw = z.bytes(start, e.compSize, `条目 "${e.name}" 的压缩数据`)
    let data
    if (e.method === 0) {
      if (e.compSize !== e.uncSize) {
        throw new Error(`zip 结构损坏：条目 "${e.name}" 是未压缩存储，但声明压缩长度 ${e.compSize} ≠ 未压缩长度 ${e.uncSize}`)
      }
      data = raw
    } else {
      // 前置校验已保证 uncSize ≤ maxEntry 且 totalOut + uncSize ≤ maxTotal，所以上限就是声明值；
      // `Math.max(1, …)` 是因为 zlib 不接受 maxOutputLength=0（会抛 ERR_OUT_OF_RANGE）。
      const cap = Math.max(1, e.uncSize)
      try {
        data = new Uint8Array(inflateRawSync(raw, { maxOutputLength: cap }))
      } catch (err) {
        if (err?.code === 'ERR_BUFFER_TOO_LARGE') {
          throw new Error(`zip 炸弹防护：条目 "${e.name}" 的实际解压体积超过声明的 ${fmtBytes(e.uncSize)}（本次上限 ${fmtBytes(cap)}），已中止`)
        }
        if (err?.code === 'Z_BUF_ERROR') {
          throw new Error(`zip 结构损坏：条目 "${e.name}" 的 deflate 数据被截断（声明压缩长度 ${e.compSize}）`)
        }
        if (err?.code === 'Z_DATA_ERROR') {
          throw new Error(`zip 结构损坏：条目 "${e.name}" 的 deflate 数据非法`)
        }
        throw err
      }
      // —— 解压后：实际长度必须等于声明值（旧实现从不比对）——
      if (data.length !== e.uncSize) {
        throw new Error(`zip 结构损坏：条目 "${e.name}" 声明解压后 ${e.uncSize} 字节，实际得到 ${data.length} 字节`)
      }
    }
    // CRC32：raw deflate 自带无校验和，尺寸相符 + CRC 相符才能同时抓住"负载损坏但仍可解压"
    // 与"尺寸声明与真实流不符"。`crc === 0` 一律放行（部分流式写包器不填 CD 的 CRC）。
    // 成本实测约 440 MB/s（64 MiB / 146 ms），相对 OCR/渲染可忽略。
    if (checkCrc && e.crc !== 0 && crc32(data) !== e.crc) {
      throw new Error(`zip 数据损坏：条目 "${e.name}" 的 CRC32 不匹配（声明 0x${e.crc.toString(16)}，实际 0x${crc32(data).toString(16)}）；如确认是无 CRC 的旧包可设 DSH_OFFICE_ZIP_CRC=0 跳过`)
    }
    totalOut += data.length
    dataCache.set(name, data)
    return data
  }
  return {
    names: [...entries.keys()],
    has: name => entries.has(name),
    get: read,
    getText(name) { const d = read(name); return d === undefined ? undefined : decUtf8.decode(d) },
  }
}

// ---------------------------------------------------------------------------
// 写侧：数据源与长度字段必须真的装得下 —— 装不下就写 **ZIP64**，绝不静默 wrap
// ---------------------------------------------------------------------------
const MAX_U16 = 0xffff
const MAX_U32 = 0xffffffff

// ZIP64（R18 任务 B）：`0xffff` / `0xffffffff` 是**哨兵值**而不是普通取值，所以字段一旦
// ≥ 哨兵就必须把字段改写成哨兵、并把真值放进 ZIP64 扩展字段 / ZIP64 EOCD（APPNOTE 6.3.6、
// 4.5.3、4.4.1.4）。旧实现遇到这些输入一律报错（"当前实现不写 ZIP64"）。
const ZIP64_EXTRA_ID = 0x0001
const Z64_SENTINEL = MAX_U32
const Z64_FIXED_SIZE = 56        // EOCD64 固定 56 字节（含 4+8 头）
const Z64_TAIL_SIZE = 44         // EOCD64 "size of record" 字段的值（56-12）
/** ZIP64 版本号：4.5 = 45（写 zip64 结构的包必须声明它）。 */
const VERSION_ZIP64 = 45
const VERSION_DEFAULT = 20

/**
 * 组装一个 ZIP64 扩展字段（APPNOTE 4.5.3）：ID 0x0001 + 长度 + 若干个 8 字节小端值。
 * `keys` 的**顺序必须是规范钉死的那一个**（未压缩尺寸 → 压缩尺寸 → 本地头偏移 → 起始磁盘号），
 * 且只允许出现"对应字段已写成哨兵"的那些值 —— 读侧是按顺序 + 哨兵位置反推的。
 * 导出只为让单测能直接核对记录布局（4 GiB 级条目无法在自动化里真造出来）。
 */
export function zip64ExtraField(keys, vals) {
  const body = Buffer.alloc(keys.length * 8)
  keys.forEach((k, i) => body.writeBigUInt64LE(BigInt(vals[k]), i * 8))
  const out = Buffer.alloc(4 + body.length)
  out.writeUInt16LE(ZIP64_EXTRA_ID, 0)
  out.writeUInt16LE(body.length, 2)
  body.copy(out, 4)
  return out
}

/** ZIP64 EOCD 记录（56 字节）+ ZIP64 定位记录（20 字节）。定位记录里指向的偏移 = EOCD64 自身位置。 */
function zip64EndRecords({ count, cdSize, cdOffset }) {
  const z = Buffer.alloc(Z64_FIXED_SIZE)
  z.writeUInt32LE(SIG_Z64_EOCD, 0)
  z.writeBigUInt64LE(BigInt(Z64_TAIL_SIZE), 4)
  z.writeUInt16LE(VERSION_ZIP64, 12)   // version made by
  z.writeUInt16LE(VERSION_ZIP64, 14)   // version needed to extract
  z.writeUInt32LE(0, 16)               // 本磁盘号
  z.writeUInt32LE(0, 20)               // 中央目录所在磁盘号
  z.writeBigUInt64LE(BigInt(count), 24)      // 本磁盘上的条目数
  z.writeBigUInt64LE(BigInt(count), 32)      // 条目总数
  z.writeBigUInt64LE(BigInt(cdSize), 40)
  z.writeBigUInt64LE(BigInt(cdOffset), 48)
  const loc = Buffer.alloc(20)
  loc.writeUInt32LE(SIG_Z64_LOCATOR, 0)
  loc.writeUInt32LE(0, 4)                                  // ZIP64 EOCD 所在磁盘号
  loc.writeBigUInt64LE(BigInt(cdOffset + cdSize), 8)       // EOCD64 的偏移
  loc.writeUInt32LE(1, 16)                                 // 磁盘总数
  return concatBytes([z, loc])
}

/**
 * Write entries [{name, data(string|Uint8Array), store?}] into a zip Uint8Array.
 *
 * R18 任务 B：**写出端支持真实 ZIP64**（旧实现遇到 >65535 条目 / >4 GiB 数据或偏移直接报错，
 * 因为 `Math.min(entries.length, 0xffff)` / `>= 0xffffffff ? 0 :` 会静默丢条目、产出废包）。
 * 现在：条目数 ≥ 0xffff、尺寸 / 偏移 ≥ 0xffffffff 时按规范写 ZIP64 扩展字段 + EOCD64 + 定位记录。
 * **不需要 ZIP64 的包（也就是所有真实的 ODF / OOXML 包）字节布局与旧版逐字一致**：
 * 不写扩展字段、version needed 仍是 20、EOCD 仍是 22 字节（探针
 * `work/r18-probe/probe-zip64.mjs` 逐字节对照过旧实现）。
 *
 * `opts.baseOffset`（可选，默认 0）：**测试接缝，不是稳定 API（`@internal`）** ——
 * 把整包当成从 `baseOffset` 开始的片段来记账。本仓**没有任何生产调用点**：`office_create` /
 * `office_edit` / `office_convert` 写出的包永远自包含、从 0 开始。保留它只为一件事 ——
 * 在自动化里覆盖">4 GiB 本地头偏移 → ZIP64"这条分支（真造一个 4 GiB 偏移的包不可行）。
 * 若将来真要"把 zip 嵌进别的容器"，它就是那个接缝；届时再把它提升为公开参数，
 * 并补"跨容器读回"的用例。
 *
 * @internal `opts.baseOffset` 仅供测试与本模块内部使用（见上一段）。
 *
 * 唯一的"装不下"是**条目名**：nameLen 在本地头与中央目录里都是 16 位，ZIP64 只扩尺寸与偏移，
 * **没有任何扩展名字长度的记录**，所以 >65535 字节的条目名仍然显式报错（说清为什么没有出路）。
 */
export function makeZip(entries, opts = {}) {
  const baseOffset = Number.isSafeInteger(opts.baseOffset) && opts.baseOffset >= 0 ? opts.baseOffset : 0
  const now = dosDateTime(new Date())
  const parts = []
  const central = []
  let offset = baseOffset
  let entryNeedsZip64 = false
  for (const entry of entries) {
    const localOffset = offset
    const nameBytes = enc.encode(entry.name)
    if (nameBytes.length > MAX_U16) {
      throw new Error(`makeZip：条目名过长（${nameBytes.length} 字节 > 65535）—— nameLen 在本地头与中央目录里都是 16 位字段，`
        + 'ZIP64 只扩尺寸与偏移、**没有**扩展条目名长度的记录（APPNOTE 4.5.3），所以这条没有 ZIP64 出路；请缩短条目名')
    }
    const data = typeof entry.data === 'string' ? enc.encode(entry.data) : entry.data
    const crc = crc32(data)
    let stored = data
    let method = 0
    if (!entry.store) {
      const deflated = new Uint8Array(deflateRawSync(data, { level: 9 }))
      if (deflated.length < data.length) { stored = deflated; method = 8 }
    }
    // 三个哨兵判据：尺寸 / 压缩尺寸 / 本地头偏移（都按 ≥ 哨兵算，因为 0xffffffff 本身不能当普通值）。
    const unc64 = data.length >= Z64_SENTINEL
    const comp64 = stored.length >= Z64_SENTINEL
    const off64 = localOffset >= Z64_SENTINEL
    // 本地头的 ZIP64 扩展：任一尺寸装不下时，把两个尺寸字段都写哨兵并带上两个 8 字节值
    //（只带一半会让"按哨兵位置反推字段顺序"的读侧无法自解释）。
    const lhKeys = unc64 || comp64 ? ['uncSize', 'compSize'] : []
    const lhExtra = lhKeys.length ? zip64ExtraField(lhKeys, { uncSize: data.length, compSize: stored.length }) : null
    // 中央目录的 ZIP64 扩展：只列出**本条目真的写成哨兵**的那些字段，顺序固定。
    const cdKeys = []
    if (unc64) cdKeys.push('uncSize')
    if (comp64) cdKeys.push('compSize')
    if (off64) cdKeys.push('localOffset')
    const cdExtra = cdKeys.length
      ? zip64ExtraField(cdKeys, { uncSize: data.length, compSize: stored.length, localOffset })
      : null
    if (lhExtra || cdExtra) entryNeedsZip64 = true
    const version = lhExtra ? VERSION_ZIP64 : VERSION_DEFAULT
    const flags = 0x0800
    const local = Buffer.alloc(30)
    local.writeUInt32LE(SIG_LOCAL, 0)
    local.writeUInt16LE(version, 4)
    local.writeUInt16LE(flags, 6)
    local.writeUInt16LE(method, 8)
    local.writeUInt16LE(now.time, 10)
    local.writeUInt16LE(now.date, 12)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(comp64 ? MAX_U32 : stored.length, 18)
    local.writeUInt32LE(unc64 ? MAX_U32 : data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    local.writeUInt16LE(lhExtra ? lhExtra.length : 0, 28)
    parts.push(local, nameBytes)
    if (lhExtra) parts.push(lhExtra)
    parts.push(stored)

    const cd = Buffer.alloc(46)
    cd.writeUInt32LE(SIG_CD, 0)
    cd.writeUInt16LE(cdExtra ? VERSION_ZIP64 : VERSION_DEFAULT, 4)   // version made by
    cd.writeUInt16LE(cdExtra ? VERSION_ZIP64 : VERSION_DEFAULT, 6)   // version needed
    cd.writeUInt16LE(flags, 8)
    cd.writeUInt16LE(method, 10)
    cd.writeUInt16LE(now.time, 12)
    cd.writeUInt16LE(now.date, 14)
    cd.writeUInt32LE(crc, 16)
    cd.writeUInt32LE(comp64 ? MAX_U32 : stored.length, 20)
    cd.writeUInt32LE(unc64 ? MAX_U32 : data.length, 24)
    cd.writeUInt16LE(nameBytes.length, 28)
    cd.writeUInt16LE(cdExtra ? cdExtra.length : 0, 30)
    cd.writeUInt16LE(0, 32)   // comment length
    cd.writeUInt16LE(0, 34)   // disk number start
    cd.writeUInt16LE(0, 36)   // internal attributes
    cd.writeUInt32LE(0, 38)   // external attributes
    cd.writeUInt32LE(off64 ? MAX_U32 : localOffset, 42)
    central.push(cd, nameBytes)
    if (cdExtra) central.push(cdExtra)
    offset += local.length + nameBytes.length + (lhExtra ? lhExtra.length : 0) + stored.length
  }
  let cdSize = 0
  for (const c of central) cdSize += c.length
  const cdOffset = offset
  // EOCD 侧的三个哨兵：条目数（≥ 0xffff —— 0xffff 本身就读作"见 ZIP64 记录"，见 openZip）、
  // 中央目录大小、中央目录偏移。任一处要 ZIP64（含任一条目要 ZIP64）就整包写 ZIP64 记录。
  const needCount = entries.length >= MAX_U16
  const needCdSize = cdSize >= Z64_SENTINEL
  const needCdOffset = cdOffset >= Z64_SENTINEL
  const useZip64 = entryNeedsZip64 || needCount || needCdSize || needCdOffset
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(SIG_EOCD, 0)
  eocd.writeUInt16LE(0, 4)   // this disk
  eocd.writeUInt16LE(0, 6)   // disk with central directory
  eocd.writeUInt16LE(needCount ? MAX_U16 : entries.length, 8)
  eocd.writeUInt16LE(needCount ? MAX_U16 : entries.length, 10)
  eocd.writeUInt32LE(needCdSize ? MAX_U32 : cdSize, 12)
  eocd.writeUInt32LE(needCdOffset ? MAX_U32 : cdOffset, 16)
  eocd.writeUInt16LE(0, 20)  // comment length
  if (!useZip64) return concatBytes([...parts, ...central, eocd])
  return concatBytes([...parts, ...central, zip64EndRecords({ count: entries.length, cdSize, cdOffset }), eocd])
}

function dosDateTime(d) {
  const year = Math.max(1980, d.getFullYear())
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | (Math.floor(d.getSeconds() / 2)),
    date: ((year - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  }
}

function concatBytes(list) {
  let total = 0
  const views = list.map(item => {
    if (item instanceof Uint8Array) {
      const view = new Uint8Array(item.buffer, item.byteOffset, item.byteLength)
      total += view.length
      return view
    }
    const buf = Buffer.from(item)
    total += buf.length
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
  })
  const out = new Uint8Array(total)
  let at = 0
  for (const view of views) { out.set(view, at); at += view.length }
  return out
}