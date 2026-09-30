// image.js — 零依赖图片嗅探/解码，供插图链路使用（docx `w:drawing` 的尺寸、PDF `/XObject` 的采样数据）。
//
// 只做两件必要的事：
//   ① 读出**像素尺寸**（docx 需要 EMU 换算，PDF 需要 `/Width` `/Height`）；
//   ② 解出**未滤波的原始采样**（PDF 走 `FlateDecode`；JPEG 直接 `DCTDecode` 原样内嵌）。
// 支持的原始采样来源（第二轮需求 4a/4b）：PNG（灰度/RGB/RGBA/**索引图**）、
// BMP（8-bit 调色板 / 24-bit / 32-bit）、GIF（**首帧**，含透明索引与交错行序）。
// 不引入任何第三方依赖；不认识的格式返回 null，由调用方显式报"跳过"。
import { readFileSync } from 'node:fs'
import { decodePng } from './png.js'

/** 从 `data:<mime>;base64,<payload>` 里解出字节（不是 data URL 时返回 null）。 */
export function decodeDataUrl(s) {
  const m = /^data:([\w.+-]+\/[\w.+-]+)?(;charset=[\w-]+)?(;base64)?,(.*)$/s.exec(String(s ?? '').trim())
  if (!m) return null
  try {
    return m[3] ? Buffer.from(m[4], 'base64') : Buffer.from(decodeURIComponent(m[4]), 'latin1')
  } catch { return null }
}

/**
 * 读取图片字节：既接受**文件路径**，也接受 `base64:…` / `data:image/…;base64,…` 内联串。
 * @returns {{buf:Buffer, from:'path'|'base64', path:string}}
 */
export function readImageBytes(input) {
  const raw = String(input ?? '').trim()
  if (!raw) throw new Error('图片参数为空')
  const dataUrl = decodeDataUrl(raw)
  if (dataUrl) return { buf: dataUrl, from: 'base64', path: '' }
  const b64 = /^base64:(.*)$/s.exec(raw)
  if (b64) {
    try { return { buf: Buffer.from(b64[1].replace(/\s+/g, ''), 'base64'), from: 'base64', path: '' } }
    catch (e) { throw new Error(`base64 图片解不开：${e?.message || e}`) }
  }
  return { buf: readFileSync(raw), from: 'path', path: raw }
}

function pngInfo(buf) {
  if (buf.length < 33 || buf.readUInt32BE(0) !== 0x89504e47) return null
  return { kind: 'png', width: buf.readUInt32BE(16), height: buf.readUInt32BE(20), colorType: buf[25], bitDepth: buf[24] }
}

/** JPEG：扫段找 SOFn（C0–CF，去掉 C4/C8/CC）读尺寸与分量数。 */
function jpegInfo(buf) {
  if (buf.length < 4 || buf[0] !== 0xff || buf[1] !== 0xd8) return null
  let p = 2
  while (p + 4 <= buf.length) {
    if (buf[p] !== 0xff) { p++; continue }
    let marker = buf[p + 1]
    while (marker === 0xff && p + 2 < buf.length) { p++; marker = buf[p + 1] }
    p += 2
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue
    if (marker === 0xd9 || marker === 0xda) break          // EOI / SOS：后面是熵编码数据
    if (p + 2 > buf.length) break
    const len = buf.readUInt16BE(p)
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (p + 7 > buf.length) break
      return { kind: 'jpeg', height: buf.readUInt16BE(p + 3), width: buf.readUInt16BE(p + 5), components: buf[p + 7] }
    }
    p += len
  }
  return { kind: 'jpeg', width: 0, height: 0, components: 3 }
}

function gifInfo(buf) {
  if (buf.length < 10 || buf.toString('latin1', 0, 3) !== 'GIF') return null
  return { kind: 'gif', width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) }
}

function bmpInfo(buf) {
  if (buf.length < 26 || buf.toString('latin1', 0, 2) !== 'BM') return null
  return { kind: 'bmp', width: buf.readInt32LE(18), height: Math.abs(buf.readInt32LE(22)) }
}

/** 嗅探图片：返回 `{kind, width, height, ...}`，不认识时 null。 */
export function sniffImage(buf) {
  return pngInfo(buf) || jpegInfo(buf) || gifInfo(buf) || bmpInfo(buf)
}

/** 图片文件的目标扩展名（写 docx media 部件用）。 */
export function imageExt(info) {
  return info?.kind === 'jpeg' ? 'jpeg' : (info?.kind || 'bin')
}

/** MIME（[Content_Types].xml 的 Default 用）。 */
export function imageMime(kind) {
  return ({ png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' })[kind] || 'application/octet-stream'
}

/**
 * PNG → 未滤波的原始采样数据（PDF `/FlateDecode` 用）。
 * @returns {null|{width:number,height:number,channels:number,colorSpace:string,data:Buffer,alpha:Buffer|null}}
 */
export function pngToRaw(buf) {
  const info = pngInfo(buf)
  if (!info) return null
  let decoded
  try { decoded = decodePng(buf) } catch { return null }
  if (info.colorType === 3) return indexedPngToRaw(decoded)   // 调色板图：PLTE/tRNS 展开（需求 4a）
  const { channels, rows, info: meta } = decoded
  const w = meta.width
  const h = rows.length
  const cs = channels === 1 || channels === 2 ? 'DeviceGray' : 'DeviceRGB'
  const hasAlpha = channels === 2 || channels === 4
  const data = Buffer.alloc(w * h * (hasAlpha ? channels - 1 : channels))
  let alpha = hasAlpha ? Buffer.alloc(w * h) : null
  let at = 0
  let aAt = 0
  for (const row of rows) {
    for (let x = 0; x < w; x++) {
      const o = x * channels
      if (channels === 1) data[at++] = row[o]
      else if (channels === 2) { data[at++] = row[o]; alpha[aAt++] = row[o + 1] }
      else if (channels === 3) { data[at++] = row[o]; data[at++] = row[o + 1]; data[at++] = row[o + 2] }
      else { data[at++] = row[o]; data[at++] = row[o + 1]; data[at++] = row[o + 2]; alpha[aAt++] = row[o + 3] }
    }
  }
  return { width: w, height: h, channels: hasAlpha ? channels - 1 : channels, colorSpace: cs, data, alpha }
}

/** 调色板 PNG（colorType 3）→ RGB/RGBA：`PLTE` 展开 + `tRNS` 透明表。 */
function indexedPngToRaw({ info, rows, palette, trns }) {
  if (!palette || palette.length < 3) return null
  const count = Math.floor(palette.length / 3)
  const w = info.width
  const h = rows.length
  let hasAlpha = false
  if (trns) for (let i = 0; i < Math.min(trns.length, count); i++) if (trns[i] < 255) { hasAlpha = true; break }
  const data = Buffer.alloc(w * h * 3)
  const alpha = hasAlpha ? Buffer.alloc(w * h) : null
  let at = 0
  let aAt = 0
  for (const row of rows) {
    for (let x = 0; x < w; x++) {
      const i = Math.min(row[x], count - 1)
      const o = i * 3
      data[at++] = palette[o]
      data[at++] = palette[o + 1]
      data[at++] = palette[o + 2]
      if (alpha) alpha[aAt++] = trns && i < trns.length ? trns[i] : 255
    }
  }
  return { width: w, height: h, channels: 3, colorSpace: 'DeviceRGB', data, alpha }
}

/**
 * BMP → 未压缩原始采样（PDF `/FlateDecode` 用）。需求 4b。
 * 支持 BITMAPINFOHEADER（≥40 字节）的 8-bit 调色板 / 24-bit / 32-bit（BI_RGB、BI_BITFIELDS），
 * 自上而下（负高度）与自下而上两种行序都处理；1/4-bit 与 RLE 压缩**返回 null**，
 * 由调用方记 `imagesSkipped`（写明原因），绝不猜。
 */
export function bmpToRaw(buf) {
  if (buf.length < 54 || buf.toString('latin1', 0, 2) !== 'BM') return null
  const dataOff = buf.readUInt32LE(10)
  const dibSize = buf.readUInt32LE(14)
  if (dibSize < 40) return null
  const w = buf.readInt32LE(18)
  const rawH = buf.readInt32LE(22)
  const h = Math.abs(rawH)
  const bpp = buf.readUInt16LE(28)
  const compression = buf.readUInt32LE(30)
  if (w <= 0 || h <= 0) return null
  if (compression !== 0 && compression !== 3) return null
  if (bpp !== 8 && bpp !== 24 && bpp !== 32) return null
  let palette = null
  if (bpp === 8) {
    const palStart = 14 + dibSize
    const count = Math.max(0, Math.floor((dataOff - palStart) / 4))
    if (!count) return null
    palette = Buffer.alloc(count * 3)
    for (let i = 0; i < count; i++) {
      const o = palStart + i * 4
      palette[i * 3] = buf[o + 2]
      palette[i * 3 + 1] = buf[o + 1]
      palette[i * 3 + 2] = buf[o]
    }
  }
  const rowSize = Math.floor((bpp * w + 31) / 32) * 4
  if (dataOff + rowSize * h > buf.length) return null
  const topDown = rawH < 0
  const hasAlpha = bpp === 32
  const data = Buffer.alloc(w * h * 3)
  const alpha = hasAlpha ? Buffer.alloc(w * h) : null
  for (let y = 0; y < h; y++) {
    const src = dataOff + (topDown ? y : h - 1 - y) * rowSize
    for (let x = 0; x < w; x++) {
      const d = (y * w + x) * 3
      if (bpp === 24) {
        const o = src + x * 3
        data[d] = buf[o + 2]; data[d + 1] = buf[o + 1]; data[d + 2] = buf[o]
      } else if (bpp === 32) {
        const o = src + x * 4
        data[d] = buf[o + 2]; data[d + 1] = buf[o + 1]; data[d + 2] = buf[o]
        alpha[y * w + x] = buf[o + 3]
      } else {
        const o = Math.min(buf[src + x], palette.length / 3 - 1) * 3
        data[d] = palette[o]; data[d + 1] = palette[o + 1]; data[d + 2] = palette[o + 2]
      }
    }
  }
  return { width: w, height: h, channels: 3, colorSpace: 'DeviceRGB', data, alpha }
}

/** 交错 GIF 的 4 个 pass 展开成实际行号序列（0,8,16… → 4,12,… → 2,6,… → 1,3,…）。 */
function gifInterlaceOrder(height) {
  const rows = []
  for (const [start, step] of [[0, 8], [4, 8], [2, 4], [1, 2]]) {
    for (let y = start; y < height; y += step) rows.push(y)
  }
  return rows
}

/** GIF 的 LZW 变长码解码（首帧用）。畸形数据有护栏，不会无限增长。 */
function lzwDecode(minCodeSize, data, expected) {
  const clear = 1 << minCodeSize
  const eoi = clear + 1
  let codeSize = minCodeSize + 1
  let dict = []
  const reset = () => {
    dict = new Array(clear + 2)
    for (let i = 0; i < clear; i++) dict[i] = [i]
    dict[clear] = null
    dict[eoi] = null
  }
  reset()
  const out = []
  let bitPos = 0
  let prev = null
  const readCode = () => {
    let code = 0
    for (let i = 0; i < codeSize; i++) {
      const byte = data[bitPos >> 3]
      if (byte === undefined) return -1
      code |= ((byte >> (bitPos & 7)) & 1) << i
      bitPos++
    }
    return code
  }
  for (;;) {
    const code = readCode()
    if (code < 0) break
    if (code === clear) { reset(); codeSize = minCodeSize + 1; prev = null; continue }
    if (code === eoi) break
    let entry
    if (dict[code]) entry = dict[code]
    else if (prev) entry = prev.concat(prev[0])
    else break
    for (const v of entry) out.push(v)
    if (prev) {
      dict.push(prev.concat(entry[0]))
      if (dict.length >= (1 << codeSize) && codeSize < 12) codeSize++
    }
    prev = entry
    if (expected && out.length > expected * 4) break
  }
  return out
}

/**
 * GIF（**首帧**）→ 未压缩原始采样。需求 4b。
 * 支持全局/局部调色板、Graphic Control Extension 的透明索引、交错行序；
 * 只取第一帧（动画 GIF 的后续帧不合成），返回 null 的情形由调用方记账。
 */
export function gifToRaw(buf) {
  if (buf.length < 13 || buf.toString('latin1', 0, 3) !== 'GIF') return null
  const w = buf.readUInt16LE(6)
  const h = buf.readUInt16LE(8)
  if (!w || !h) return null
  const packed = buf[10]
  let p = 13
  let gct = null
  if (packed & 0x80) {
    const size = 3 * (2 << (packed & 7))
    if (p + size > buf.length) return null
    gct = buf.subarray(p, p + size)
    p += size
  }
  const skipBlocks = at => {
    while (at < buf.length) {
      const len = buf[at++]
      if (!len) break
      at += len
    }
    return at
  }
  let transparent = -1
  let frame = null
  while (p < buf.length) {
    const marker = buf[p++]
    if (marker === 0x3b) break
    if (marker === 0x21) {
      const label = buf[p++]
      if (label === 0xf9 && p + 6 <= buf.length && (buf[p + 1] & 1)) transparent = buf[p + 4]
      p = skipBlocks(p)
      continue
    }
    if (marker !== 0x2c) return null
    if (p + 9 > buf.length) return null
    const left = buf.readUInt16LE(p)
    const top = buf.readUInt16LE(p + 2)
    const fw = buf.readUInt16LE(p + 4)
    const fh = buf.readUInt16LE(p + 6)
    const fpacked = buf[p + 8]
    p += 9
    let table = gct
    if (fpacked & 0x80) {
      const size = 3 * (2 << (fpacked & 7))
      if (p + size > buf.length) return null
      table = buf.subarray(p, p + size)
      p += size
    }
    if (!table || p >= buf.length) return null
    const minCodeSize = buf[p++]
    const chunks = []
    while (p < buf.length) {
      const len = buf[p++]
      if (!len) break
      if (p + len > buf.length) break
      chunks.push(buf.subarray(p, p + len))
      p += len
    }
    let indices = lzwDecode(minCodeSize, Buffer.concat(chunks), fw * fh)
    if (fpacked & 0x40) {
      const order = gifInterlaceOrder(fh)
      const re = new Array(fw * fh).fill(0)
      order.forEach((row, k) => {
        for (let x = 0; x < fw; x++) re[row * fw + x] = indices[k * fw + x]
      })
      indices = re
    }
    frame = { left, top, fw, fh, table, indices }
    break
  }
  if (!frame || !frame.indices.length) return null
  const colors = Math.floor(frame.table.length / 3)
  const hasAlpha = transparent >= 0
  const data = Buffer.alloc(w * h * 3)
  const alpha = hasAlpha ? Buffer.alloc(w * h) : null
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const d = (y * w + x) * 3
      const inside = x >= frame.left && x < frame.left + frame.fw && y >= frame.top && y < frame.top + frame.fh
      const idx = inside ? frame.indices[(y - frame.top) * frame.fw + (x - frame.left)] : -1
      if (idx === undefined || idx < 0 || idx >= colors || idx === transparent) {
        data[d] = 255; data[d + 1] = 255; data[d + 2] = 255
        if (alpha) alpha[y * w + x] = 0
      } else {
        data[d] = frame.table[idx * 3]
        data[d + 1] = frame.table[idx * 3 + 1]
        data[d + 2] = frame.table[idx * 3 + 2]
        if (alpha) alpha[y * w + x] = 255
      }
    }
  }
  return { width: w, height: h, channels: 3, colorSpace: 'DeviceRGB', data, alpha }
}

/**
 * 统一入口：任意已嗅探出的图片字节 → 未压缩原始采样（JPEG 返回 null：它走 `/DCTDecode` 原样内嵌）。
 * @returns {null|{width:number,height:number,channels:number,colorSpace:string,data:Buffer,alpha:Buffer|null}}
 */
export function imageToRaw(buf) {
  const info = sniffImage(buf)
  if (!info) return null
  if (info.kind === 'png') return pngToRaw(buf)
  if (info.kind === 'bmp') return bmpToRaw(buf)
  if (info.kind === 'gif') return gifToRaw(buf)
  return null
}
