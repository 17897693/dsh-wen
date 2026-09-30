// Minimal PNG support for OCR banding: decode a rendered page image and cut
// it into horizontal strips, so a dense page can be recognised in several
// passes when the vision model's output length is capped.
// Handles the non-interlaced 8-bit greyscale / RGB / RGBA images that the
// Windows PDF renderer emits (color type 0, 2, 4, 6) plus — 第二轮需求 4a —
// 8-bit indexed images (color type 3) via their PLTE/tRNS chunks.
import { inflateSync, deflateSync } from 'node:zlib'

const CRC_TABLE = (() => {
  const t = new Int32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[i] = c
  }
  return t
})()

function crc32(buf) {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

const CHANNELS = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }

export function readPngInfo(buf) {
  if (buf.length < 33 || buf.readUInt32BE(0) !== 0x89504e47) throw new Error('not a PNG')
  return {
    width: buf.readUInt32BE(16),
    height: buf.readUInt32BE(20),
    bitDepth: buf[24],
    colorType: buf[25],
    interlace: buf[28],
  }
}

function chunk(type, data) {
  const out = Buffer.alloc(12 + data.length)
  out.writeUInt32BE(data.length, 0)
  out.write(type, 4, 'latin1')
  data.copy(out, 8)
  out.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, 'latin1'), data])), 8 + data.length)
  return out
}

/**
 * Decode a PNG into `{info, channels, rows}` (unfiltered scanlines).
 * Exported for the插图链路：PDF 写出端要把 PNG 解成原始采样（`/FlateDecode`），
 * docx 插图要读像素尺寸 —— OCR 那条路只需要裁带。
 * 注意：只支持 8-bit 非隔行 colorType 0/2/4/6（Windows 渲染器与常见导出器的形态）。
 */
export function decodePng(buf) {
  return decodeRows(buf)
}

/**
 * Decode the image into unfiltered scanlines (one Buffer per row).
 * `palette`/`trns` 是索引图（colorType 3）的 `PLTE`/`tRNS` 原字节，供调用方展开成 RGB/RGBA。
 */
function decodeRows(buf) {
  const info = readPngInfo(buf)
  if (info.bitDepth !== 8) throw new Error(`unsupported PNG bit depth ${info.bitDepth}`)
  if (info.interlace !== 0) throw new Error('interlaced PNG not supported')
  const channels = CHANNELS[info.colorType]
  if (!channels) throw new Error(`unsupported PNG color type ${info.colorType}`)
  const idat = []
  let palette = null
  let trns = null
  for (let p = 8; p + 8 <= buf.length;) {
    const len = buf.readUInt32BE(p)
    const type = buf.toString('latin1', p + 4, p + 8)
    if (type === 'IDAT') idat.push(buf.subarray(p + 8, p + 8 + len))
    else if (type === 'PLTE') palette = Buffer.from(buf.subarray(p + 8, p + 8 + len))
    else if (type === 'tRNS') trns = Buffer.from(buf.subarray(p + 8, p + 8 + len))
    if (type === 'IEND') break
    p += 12 + len
  }
  // 索引图（colorType 3）必须有 PLTE，否则解出来的是无意义索引 —— 宁可报错也不猜
  if (info.colorType === 3 && (!palette || palette.length < 3)) throw new Error('indexed PNG 缺少 PLTE 调色板')
  const raw = inflateSync(Buffer.concat(idat))
  const rowBytes = info.width * channels
  const rows = []
  let prev = Buffer.alloc(rowBytes)
  for (let y = 0, at = 0; y < info.height; y++) {
    const filter = raw[at++]
    const cur = Buffer.from(raw.subarray(at, at + rowBytes))
    at += rowBytes
    if (filter === 1) for (let i = channels; i < rowBytes; i++) cur[i] = (cur[i] + cur[i - channels]) & 0xff
    else if (filter === 2) for (let i = 0; i < rowBytes; i++) cur[i] = (cur[i] + prev[i]) & 0xff
    else if (filter === 3) for (let i = 0; i < rowBytes; i++) cur[i] = (cur[i] + Math.floor(((i >= channels ? cur[i - channels] : 0) + prev[i]) / 2)) & 0xff
    else if (filter === 4) {
      for (let i = 0; i < rowBytes; i++) {
        const a = i >= channels ? cur[i - channels] : 0
        const b = prev[i]
        const c = i >= channels ? prev[i - channels] : 0
        const p = a + b - c
        const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c)
        cur[i] = (cur[i] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 0xff
      }
    }
    rows.push(cur)
    prev = cur
  }
  return { info, channels, rows, palette, trns }
}

/** Encode rows back into an 8-bit PNG of the same color type. */
function encodeRows(info, channels, rows, palette = null, trns = null) {
  const rowBytes = info.width * channels
  const raw = Buffer.alloc((rowBytes + 1) * rows.length)
  rows.forEach((row, y) => {
    raw[y * (rowBytes + 1)] = 0
    row.copy(raw, y * (rowBytes + 1) + 1, 0, rowBytes)
  })
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(info.width, 0)
  ihdr.writeUInt32BE(rows.length, 4)
  ihdr[8] = 8
  ihdr[9] = info.colorType
  ihdr[10] = 0
  ihdr[11] = 0
  ihdr[12] = 0
  const extra = []
  if (info.colorType === 3) {
    // 索引图重打包必须带上 PLTE（否则产物是非法 PNG，读图端会拒绝）
    if (!palette) throw new Error('indexed PNG 重打包缺少 PLTE')
    extra.push(chunk('PLTE', palette))
    if (trns) extra.push(chunk('tRNS', trns))
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    ...extra,
    chunk('IDAT', deflateSync(raw, { level: 6 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

/**
 * Cut the image into `count` horizontal bands and return one band as a PNG.
 * `index` is 0-based. Bands overlap by `overlap` rows so a text line split
 * across the boundary is fully visible in one of them.
 */
export function cropPngBand(buf, index, count, overlap = 24) {
  const { info, channels, rows, palette, trns } = decodeRows(buf)
  if (count <= 1) return buf
  const per = Math.ceil(info.height / count)
  let from = Math.max(0, index * per - (index > 0 ? overlap : 0))
  let to = Math.min(info.height, (index + 1) * per + (index < count - 1 ? overlap : 0))
  if (to <= from) { from = Math.min(from, Math.max(0, info.height - 1)); to = Math.min(info.height, from + 1) }
  return encodeRows(info, channels, rows.slice(from, to), palette, trns)
}

/**
 * 阶段二（PDF 产出质量门）：渲染级启发——黑像素（亮度 < 192）覆盖率。
 * 只吃字节，不 OCR、不消耗视觉额度。colorType 0/2/3/4/6 的 8-bit 非隔行 PNG
 * （Windows PDF 渲染器的输出形态，外加索引图）都可解；其他形态抛错由调用方降级。
 * @returns {{width:number, height:number, dark:number, ink:number}} ink = 黑像素占比 0..1，dark = 黑像素个数
 */
export function pngInkCoverage(buf) {
  const { info, channels, rows, palette } = decodeRows(buf)
  let dark = 0
  const total = Math.max(1, info.width * info.height)
  const hasAlpha = channels === 2 || channels === 4
  const indexed = info.colorType === 3
  for (const row of rows) {
    for (let x = 0; x < info.width; x++) {
      const i = x * channels
      if (hasAlpha && row[i + channels - 1] < 16) continue // 透明 = 空白
      let lum
      if (indexed && palette) {
        const o = Math.min(row[i], Math.floor(palette.length / 3) - 1) * 3
        lum = (palette[o] + palette[o + 1] + palette[o + 2]) / 3
      } else {
        lum = channels >= 3 ? (row[i] + row[i + 1] + row[i + 2]) / 3 : row[i]
      }
      if (lum < 192) dark++
    }
  }
  return { width: info.width, height: info.height, dark, ink: dark / total }
}

