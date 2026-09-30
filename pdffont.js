// pdffont.js — 系统字体嵌入（阶段一）：零依赖 TTF/TTC 解析 + 恒等 GID 子集化。
//
// 目标：writePdf 产出的中文/emoji PDF 内嵌真实字形（Identity-H + FontFile2 +
// 完整 ToUnicode），在 WinRT/Edge/Chrome 里正常渲染、office_read 提取不丢字符。
//
// 设计要点（与旧 STSong-Light 预定义 CMap 路径的关系）：
//   * 子集**不重排 GID**：内容流里的 2 字节码就是原字体 GID（CIDToGIDMap/Identity）。
//     因此 composite 字形的组件引用、hmtx/cmap/kern 等按 GID 索引的表**原样保留即有效**，
//     只需把未用到的 glyf 清零并重建 loca——工程量与风险都远小于真正的 GID 重排子集。
//   * .ttc（simsun.ttc 是集合）先抽 font[0] 成独立 sfnt；EBDT/EBLC（点阵 strikes）
//     与 DSIG 等对 PDF 无用或子集后必失效的表剔除，控制体积。
//   * head.indexToLocFormat 统一改成长格式（32 位偏移），checkSumAdjustment 收尾重算。
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

// ---------------------------------------------------------------------------
// 字体源：正文宋体 -> 备选黑体/雅黑；emoji -> Segoe UI Symbol 兜底。
// 每类按顺序懒加载：第一个能解析的胜出。
// ---------------------------------------------------------------------------

const FONT_DIRS = (() => {
  const windir = process.env.WINDIR || process.env.SystemRoot || 'C:\\Windows'
  return [join(windir, 'Fonts'), 'C:/Windows/Fonts']
})()

const SOURCES = {
  cjk: [['simsun.ttc', 'SimSun'], ['simhei.ttf', 'SimHei'], ['msyh.ttc', 'MicrosoftYaHei']],
  emoji: [['seguiemj.ttf', 'SegoeUIEmoji']],
  symbol: [['seguisym.ttf', 'SegoeUISymbol']],
}

const parsedCache = new Map() // path -> SfntFont | null（解析失败缓存 null，不重试）

// ---------------------------------------------------------------------------
// sfnt / TTC 头解析
// ---------------------------------------------------------------------------

/** 从 buf 的 off 处读出 sfnt 目录；'ttcf' 时递归到第 0 个字体。 */
function sfntAt(buf, off = 0) {
  if (off + 12 > buf.length) throw new Error('sfnt: truncated header')
  const tag = buf.toString('latin1', off, off + 4)
  if (tag === 'ttcf') {
    const numFonts = buf.readUInt32BE(off + 8)
    if (!numFonts) throw new Error('ttcf: empty collection')
    return sfntAt(buf, buf.readUInt32BE(off + 12)) // font[0] = 正文宋体
  }
  const scaler = buf.readUInt32BE(off)
  const ok = scaler === 0x00010000 || tag === 'true' || tag === 'typ1'
  if (!ok) throw new Error(`sfnt: unsupported scaler tag ${JSON.stringify(tag)}`)
  const numTables = buf.readUInt16BE(off + 4)
  const tables = new Map()
  for (let i = 0; i < numTables; i++) {
    const p = off + 12 + i * 16
    if (p + 16 > buf.length) throw new Error('sfnt: truncated table directory')
    const name = buf.toString('latin1', p, p + 4)
    tables.set(name, { checksum: buf.readUInt32BE(p + 4), offset: buf.readUInt32BE(p + 8), length: buf.readUInt32BE(p + 12) })
  }
  return { tag, numTables, tables }
}

// ---------------------------------------------------------------------------
// cmap：码点 -> GID（format 0/4/6/12；优先 Windows UCS-2/UCS-4 子表）
// ---------------------------------------------------------------------------

function readCmapSubtable(buf, sub) {
  const fmt = buf.readUInt16BE(sub)
  const map = new Map()
  if (fmt === 0) {
    for (let c = 0; c < 256; c++) { const g = buf.readUInt8(sub + 6 + c); if (g) map.set(c, g) }
  } else if (fmt === 4) {
    const segX2 = buf.readUInt16BE(sub + 6)
    const seg = segX2 / 2
    const endO = sub + 14
    const startO = endO + segX2 + 2
    const deltaO = startO + segX2
    const rangeO = deltaO + segX2
    for (let s = 0; s < seg; s++) {
      const start = buf.readUInt16BE(startO + s * 2)
      const end = buf.readUInt16BE(endO + s * 2)
      if (start === 0xffff || start > end) continue
      const delta = buf.readInt16BE(deltaO + s * 2)
      const rangeOff = buf.readUInt16BE(rangeO + s * 2)
      for (let c = start; c <= end; c++) {
        let g
        if (rangeOff === 0) g = (c + delta) & 0xffff
        else {
          const gi = rangeO + s * 2 + rangeOff + (c - start) * 2
          if (gi + 1 >= buf.length) break
          g = buf.readUInt16BE(gi)
          if (g) g = (g + delta) & 0xffff
        }
        if (g) map.set(c, g)
      }
    }
  } else if (fmt === 6) {
    const first = buf.readUInt16BE(sub + 6)
    const count = buf.readUInt16BE(sub + 8)
    for (let i = 0; i < count; i++) { const g = buf.readUInt16BE(sub + 10 + i * 2); if (g) map.set(first + i, g) }
  } else if (fmt === 12) {
    const n = buf.readUInt32BE(sub + 12)
    for (let i = 0; i < n; i++) {
      const p = sub + 16 + i * 12
      if (p + 12 > buf.length) break
      const sc = buf.readUInt32BE(p)
      const ec = buf.readUInt32BE(p + 4)
      const sg = buf.readUInt32BE(p + 8)
      const span = Math.min(ec - sc, 0xffff)
      for (let c = 0; c <= span; c++) map.set(sc + c, sg + c)
    }
  }
  return map
}

function readCmap(buf, rec) {
  const base = rec.offset
  const num = buf.readUInt16BE(base + 2)
  let best = null
  let bestScore = -1
  for (let i = 0; i < num; i++) {
    const p = base + 4 + i * 8
    const pid = buf.readUInt16BE(p)
    const eid = buf.readUInt16BE(p + 2)
    const off = buf.readUInt32BE(p + 4)
    let score = -1
    if (pid === 3 && eid === 10) score = 5      // Windows UCS-4
    else if (pid === 3 && eid === 1) score = 4  // Windows BMP
    else if (pid === 0) score = 3               // Unicode 全平台
    else if (pid === 3 && eid === 0) score = 2  // Windows Symbol（0xF000 区，见下）
    if (score > bestScore) { bestScore = score; best = { off, symbol: pid === 3 && eid === 0 } }
  }
  if (!best) return new Map()
  let map = readCmapSubtable(buf, base + best.off)
  if (best.symbol && map.size) {
    // Symbol 字体常把字形放在 0xF000-0xF0FF：补一份 ASCII 映射
    const extra = new Map()
    for (const [c, g] of map) if (c >= 0xf000 && c <= 0xf0ff) extra.set(c - 0xf000, g)
    map = new Map([...extra, ...map])
  }
  return map
}

// ---------------------------------------------------------------------------
// SfntFont：一个可查询、可子集化的字体
// ---------------------------------------------------------------------------

/**
 * 子集化时剔除的表：
 *  - EBDT/EBLC/EBSC、CBLC/CBDT、sbix：点阵 strike（体积大户，且子集后必然失效）
 *  - COLR/CPAL/SVG ：彩色字形层。PDF 的 CIDFontType2 只吃 glyf 轮廓，这些表渲染器不会读，
 *    却被整表拷贝——Segoe UI Emoji 的 COLR 单项就有 7.4MB，是中文 PDF 体积失控的主因。
 *  - DSIG/LTSH/VDMX/hdmx：签名与设备度量，子集后无意义。
 */
const DROP_TABLES = new Set(['EBDT', 'EBLC', 'EBSC', 'CBLC', 'CBDT', 'sbix', 'DSIG', 'LTSH', 'VDMX', 'hdmx',
  'COLR', 'CPAL', 'SVG '])

class SfntFont {
  constructor(buf, fallbackName) {
    this.buf = buf
    const dir = sfntAt(buf, 0)
    if (dir.tables.has('CFF ') || dir.tables.has('CFF2')) throw new Error('sfnt: CFF outlines unsupported (FontFile3 path not implemented)')
    this.dir = dir
    const head = dir.tables.get('head')
    const hhea = dir.tables.get('hhea')
    const maxp = dir.tables.get('maxp')
    const hmtx = dir.tables.get('hmtx')
    const loca = dir.tables.get('loca')
    const glyf = dir.tables.get('glyf')
    const os2 = dir.tables.get('OS/2')
    const post = dir.tables.get('post')
    if (!head || !hhea || !maxp || !hmtx || !loca || !glyf) throw new Error('sfnt: missing required tables')
    this.head = head
    this.upem = buf.readUInt16BE(head.offset + 18) || 1000
    this.indexToLocFormat = buf.readInt16BE(head.offset + 50)
    this.xMin = buf.readInt16BE(head.offset + 36)
    this.yMin = buf.readInt16BE(head.offset + 38)
    this.xMax = buf.readInt16BE(head.offset + 40)
    this.yMax = buf.readInt16BE(head.offset + 42)
    this.numGlyphs = buf.readUInt16BE(maxp.offset + 4)
    this.numHMetrics = buf.readUInt16BE(hhea.offset + 34)
    this.hheaAscent = buf.readInt16BE(hhea.offset + 4)
    this.hheaDescent = buf.readInt16BE(hhea.offset + 6)
    this.italicAngle = post ? buf.readInt32BE(post.offset + 4) : 0
    this.ascent = (os2 && os2.length >= 74 ? buf.readInt16BE(os2.offset + 68) : 0) || this.hheaAscent
    this.descent = (os2 && os2.length >= 74 ? buf.readInt16BE(os2.offset + 70) : 0) || this.hheaDescent
    this.capHeight = (os2 && os2.length >= 90) ? buf.readInt16BE(os2.offset + 88) : 0
    this.hmtx = hmtx
    this.loca = loca
    this.glyf = glyf
    this.cmapRec = dir.tables.get('cmap')
    this._cmap = null
    this.name = this._postScriptName() || fallbackName || 'EmbeddedFont'
  }

  get cmap() {
    if (!this._cmap && this.cmapRec) this._cmap = readCmap(this.buf, this.cmapRec)
    return this._cmap
  }

  /** 码点 -> GID（0 = 没有字形）。 */
  gid(cp) {
    const g = this.cmap.get(cp)
    return typeof g === 'number' && g > 0 && g < this.numGlyphs ? g : 0
  }

  /** GID 的前进宽度（font unit）。 */
  advance(gid) {
    if (gid < 0 || gid >= this.numGlyphs) return 0
    const nh = this.numHMetrics || 1
    const i = gid < nh ? gid : nh - 1
    return this.buf.readUInt16BE(this.hmtx.offset + i * 4)
  }

  _locaRange(gid) {
    const b = this.buf
    const base = this.loca.offset
    if (this.indexToLocFormat === 1) {
      return [b.readUInt32BE(base + gid * 4), b.readUInt32BE(base + gid * 4 + 4)]
    }
    return [b.readUInt16BE(base + gid * 2) * 2, b.readUInt16BE(base + gid * 2 + 2) * 2]
  }

  glyphBytes(gid) {
    const [s, e] = this._locaRange(gid)
    if (!(e > s)) return null
    const from = this.glyf.offset + s
    const to = Math.min(this.glyf.offset + e, this.buf.length)
    return to > from ? this.buf.subarray(from, to) : null
  }

  /** composite 字形引用的组件 GID 列表（simple 字形返回 []）。 */
  componentGids(glyph) {
    if (!glyph || glyph.length < 10) return []
    const nContours = glyph.readInt16BE(0)
    if (nContours >= 0) return []
    const out = []
    let p = 10
    for (;;) {
      if (p + 4 > glyph.length) break
      const flags = glyph.readUInt16BE(p)
      const gi = glyph.readUInt16BE(p + 2)
      p += 4
      out.push(gi)
      p += (flags & 0x0001) ? 4 : 2                      // ARG_1_AND_2_ARE_WORDS
      if (!(flags & 0x0008)) break                        // MORE_COMPONENTS
      if (flags & 0x0004) p += 2                          // WE_HAVE_A_SCALE
      else if (flags & 0x0040) p += 4                     // X_AND_Y_SCALE（x/y 各 2 字节）
      else if (flags & 0x0080) p += 8                     // TWO_BY_TWO（a b c d 各 2 字节）
    }
    return out
  }

  _postScriptName() {
    const rec = this.dir.tables.get('name')
    if (!rec) return null
    try {
      const b = this.buf
      const count = b.readUInt16BE(rec.offset + 2)
      const strOff = rec.offset + b.readUInt16BE(rec.offset + 4)
      let fallback = null
      for (let i = 0; i < count; i++) {
        const p = rec.offset + 6 + i * 12
        const pid = b.readUInt16BE(p)
        const nid = b.readUInt16BE(p + 6)
        const len = b.readUInt16BE(p + 8)
        const off = b.readUInt16BE(p + 10)
        if (nid !== 6 && nid !== 4) continue
        let s
        if (pid === 3 || pid === 0) {
          s = ''
          for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(b.readUInt16BE(strOff + off + k))
        } else s = b.toString('latin1', strOff + off, strOff + off + len)
        if (!s) continue
        if (nid === 6) return s
        fallback = fallback || s
      }
      return fallback
    } catch { return null }
  }

  /**
   * 恒等 GID 子集：保留 numGlyphs 不变，未用到的 glyf 清零、loca 重建为长格式，
   * 其余表原样拷贝（按 GID 索引的引用依然有效），最后重算 checkSumAdjustment。
   *
   * @param {Set<number>} used 显式用到的 GID（组件在内部递归补全）
   * @returns {Buffer} 独立 sfnt 字节
   */
  subset(used) {
    const keep = new Set([0])
    const stack = [...used]
    while (stack.length) {
      const g = stack.pop()
      if (!(g > 0) || g >= this.numGlyphs || keep.has(g)) continue
      keep.add(g)
      const glyph = this.glyphBytes(g)
      if (!glyph) continue
      for (const comp of this.componentGids(glyph)) stack.push(comp)
    }
    // 重建 glyf + loca（长格式，每字形 4 字节对齐）
    const loca = Buffer.alloc((this.numGlyphs + 1) * 4)
    const parts = []
    let cursor = 0
    for (let g = 0; g < this.numGlyphs; g++) {
      loca.writeUInt32BE(cursor, g * 4)
      const glyph = keep.has(g) ? this.glyphBytes(g) : null
      if (glyph && glyph.length) {
        parts.push(glyph)
        cursor += glyph.length
        const pad = (4 - (cursor % 4)) % 4
        if (pad) { parts.push(Buffer.alloc(pad)); cursor += pad }
      }
    }
    loca.writeUInt32BE(cursor, this.numGlyphs * 4)
    const glyf = Buffer.concat(parts)

    // 表集合：剔除 DROP_TABLES 与旧 glyf/loca/head，head 单独重写
    const outTables = new Map()
    for (const [tag, rec] of this.dir.tables) {
      if (tag === 'glyf' || tag === 'loca' || tag === 'head' || DROP_TABLES.has(tag)) continue
      const from = rec.offset
      const to = Math.min(from + rec.length, this.buf.length)
      if (to <= from) continue
      outTables.set(tag, this.buf.subarray(from, to))
    }
    const headCopy = Buffer.from(this.buf.subarray(this.head.offset, this.head.offset + this.head.length))
    headCopy.writeInt16BE(1, 50)              // indexToLocFormat = 长格式
    headCopy.writeUInt32BE(0, 8)              // checkSumAdjustment = 0（收尾再算）
    outTables.set('head', headCopy)
    outTables.set('loca', loca)
    outTables.set('glyf', glyf)

    // 组装：表目录按 tag 字典序（sfnt 规范），表体 4 字节对齐
    const tags = [...outTables.keys()].sort()
    const numTables = tags.length
    let searchRange = 16
    let entrySelector = 0
    while (searchRange * 2 <= numTables * 16) { searchRange *= 2; entrySelector++ }
    const header = Buffer.alloc(12)
    header.writeUInt32BE(0x00010000, 0)
    header.writeUInt16BE(numTables, 4)
    header.writeUInt16BE(searchRange, 6)
    header.writeUInt16BE(entrySelector, 8)
    header.writeUInt16BE(numTables * 16 - searchRange, 10)
    const dirBuf = Buffer.alloc(numTables * 16)
    const chunks = []
    let offset = 12 + numTables * 16
    let headFileOffset = -1
    tags.forEach((tag, i) => {
      const data = outTables.get(tag)
      const p = i * 16
      dirBuf.write(tag.slice(0, 4).padEnd(4, ' '), p, 'latin1')
      dirBuf.writeUInt32BE(tableChecksum(data), p + 4)
      dirBuf.writeUInt32BE(offset, p + 8)
      dirBuf.writeUInt32BE(data.length, p + 12)
      if (tag === 'head') headFileOffset = offset
      chunks.push(data)
      offset += data.length
      const pad = (4 - (offset % 4)) % 4
      if (pad) { chunks.push(Buffer.alloc(pad)); offset += pad }
    })
    const file = Buffer.concat([header, dirBuf, ...chunks])
    // checkSumAdjustment：整个文件 uint32 和 == 0xB1B2B3B4
    const adj = (0xb1b2b3b4 - tableChecksum(file)) >>> 0
    if (headFileOffset >= 0) file.writeUInt32BE(adj, headFileOffset + 8)
    return file
  }
}

function tableChecksum(data) {
  let sum = 0
  const n = data.length
  for (let i = 0; i < n; i += 4) {
    const b0 = data[i] ?? 0
    const b1 = data[i + 1] ?? 0
    const b2 = data[i + 2] ?? 0
    const b3 = data[i + 3] ?? 0
    sum = (sum + ((b0 << 24 | b1 << 16 | b2 << 8 | b3) >>> 0)) >>> 0
  }
  return sum >>> 0
}

function loadFont(kind) {
  for (const [file, fallbackName] of SOURCES[kind]) {
    for (const dir of FONT_DIRS) {
      const path = join(dir, file)
      if (parsedCache.has(path)) {
        if (parsedCache.get(path)) return parsedCache.get(path)
        continue
      }
      let font = null
      try { font = new SfntFont(readFileSync(path), fallbackName) } catch { font = null }
      parsedCache.set(path, font)
      if (font) return font
    }
  }
  return null
}

// ---------------------------------------------------------------------------
// FontChain：writePdf 布局期的逐字决策 + 落盘期的字体对象装配
// ---------------------------------------------------------------------------

export class FontChain {
  constructor() {
    this.fonts = []          // 已加载的字体（cjk -> emoji -> symbol 按需补）
    this.byKind = {}
    this.notes = []
    this._charCache = new Map()
    this._usedByFont = new Map() // SfntFont -> { gids:Set, chars:Map(gid->char) }
  }

  _ensure(kind) {
    if (kind in this.byKind) return this.byKind[kind]
    const font = loadFont(kind)
    this.byKind[kind] = font
    if (font && !this.fonts.includes(font)) {
      this.fonts.push(font)
      this._usedByFont.set(font, { gids: new Set(), chars: new Map() })
    } else if (!font) {
      this.notes.push(`${kind} font missing`)
    }
    return font
  }

  /**
   * 码点分类（memo 化；每次 writePdf 一个 chain，缓存不会串）：
   *   'latin' — WinAnsi 可表达 -> 继续走 base-14（F1/F2/F3）
   *   'embed' — 命中内嵌链 -> {f, g}（f=字体，g=GID），同时记入 used
   *   'gb'    — 内嵌链没有该字形、且 cp>0x2e80 -> 回落 STSong-Light（旧路径）
   *   'miss'  — 都不行 -> 旧语义下走 WinAnsi 的 '?' 行
   */
  classify(ch) {
    const hit = this._charCache.get(ch)
    if (hit) return hit
    const cp = ch.codePointAt(0)
    let out
    if (winAnsiOk(ch)) {
      out = { k: 'latin' }
    } else {
      let resolved = null
      for (const f of this.fonts) {
        const g = f.gid(cp)
        if (g) { resolved = { f, g }; break }
      }
      if (!resolved) {
        for (const kind of ['cjk', 'emoji', 'symbol']) {
          const f = this._ensure(kind)
          if (!f) continue
          const g = f.gid(cp)
          if (g) { resolved = { f, g }; break }
        }
      }
      if (resolved) {
        out = { k: 'embed', f: resolved.f, g: resolved.g }
        const u = this._usedByFont.get(resolved.f)
        if (u) {
          u.gids.add(resolved.g)
          if (!u.chars.has(resolved.g)) u.chars.set(resolved.g, ch)
        }
      } else if (cp > 0x2e80) out = { k: 'gb' }
      else out = { k: 'miss' }
    }
    this._charCache.set(ch, out)
    return out
  }

  /** 某内嵌字体显式用到的 GID/字符（ToUnicode 与 W 数组用）。 */
  used(font) {
    return this._usedByFont.get(font) || { gids: new Set(), chars: new Map() }
  }

  /** 字形宽度（em 小数，如 1 / 0.5）。 */
  emWidth(font, gid) {
    return font.advance(gid) / (font.upem || 1000)
  }
}

/** WinAnsi（含 cp1252 别名）能否无损表达该字符——能就继续走 base-14。 */
const WINANSI_OK_EXTRA = new Set(['€', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', 'Ž',
  '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', 'ž', 'Ÿ'])
export function winAnsiOk(ch) {
  const cp = ch.codePointAt(0)
  if (cp > 0xffff) return false
  if (cp <= 0xff) return cp < 0x7f || cp > 0x9f // C0/C1 控制符不承载正文
  return WINANSI_OK_EXTRA.has(ch)
}

/** 供测试/诊断：清空字体解析缓存。 */
export function resetFontCache() {
  parsedCache.clear()
}
