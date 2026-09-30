// PDF standard security handler (ISO 32000-1 §7.6 + ISO 32000-2 §7.6.4): enough to open
// files that are permission-restricted with an EMPTY user password — the common case for
// exam/ebook PDFs that other readers open without asking for a password.
// Supports R2/R3/R4 (RC4-40, RC4-128, AES-128) and — 第十二轮 — **R5/R6 (AES-256)**.
// A real user password is never guessed or bypassed: when the empty password fails we
// report that instead.
import { createHash, createDecipheriv, createCipheriv } from 'node:crypto'

const PAD = Buffer.from([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
])

const md5 = buf => createHash('md5').update(buf).digest()

/** RC4 is not available in modern OpenSSL providers, so it lives here. */
export function rc4(key, data) {
  const S = new Uint8Array(256)
  for (let i = 0; i < 256; i++) S[i] = i
  let j = 0
  for (let i = 0; i < 256; i++) {
    j = (j + S[i] + key[i % key.length]) & 0xff
    const t = S[i]; S[i] = S[j]; S[j] = t
  }
  const out = Buffer.alloc(data.length)
  let a = 0
  let b = 0
  for (let k = 0; k < data.length; k++) {
    a = (a + 1) & 0xff
    b = (b + S[a]) & 0xff
    const t = S[a]; S[a] = S[b]; S[b] = t
    out[k] = data[k] ^ S[(S[a] + S[b]) & 0xff]
  }
  return out
}

function padPassword(password) {
  const p = Buffer.from(String(password ?? ''), 'latin1')
  if (p.length >= 32) return p.subarray(0, 32)
  return Buffer.concat([p, PAD.subarray(0, 32 - p.length)])
}

/** Algorithm 2: file encryption key from a candidate user password. */
function fileKey(password, o, p, id0, r, lengthBits, encryptMetadata) {
  const pbuf = Buffer.alloc(4)
  pbuf.writeInt32LE(p | 0)
  let input = Buffer.concat([padPassword(password), o, pbuf, id0])
  if (r >= 4 && encryptMetadata === false) input = Buffer.concat([input, Buffer.from([0xff, 0xff, 0xff, 0xff])])
  let key = md5(input)
  const n = Math.max(5, Math.min(16, (lengthBits || 40) / 8))
  if (r >= 3) for (let i = 0; i < 50; i++) key = md5(key.subarray(0, n))
  return key.subarray(0, n)
}

/** Algorithms 4/5: does this key match /U for the candidate password? */
function userKeyMatches(key, u, id0, r) {
  if (r === 2) return rc4(key, PAD).equals(u.subarray(0, 16))
  let x = md5(Buffer.concat([PAD, id0]))
  x = rc4(key, x)
  for (let i = 1; i <= 19; i++) {
    const k = Buffer.from(key.map(b => b ^ i))
    x = rc4(k, x)
  }
  return x.subarray(0, 16).equals(u.subarray(0, 16))
}

/** Algorithm 1: per-object key. */
function objectKey(key, num, gen, aes) {
  const ext = Buffer.from([num & 0xff, (num >> 8) & 0xff, (num >> 16) & 0xff, gen & 0xff, (gen >> 8) & 0xff])
  let input = Buffer.concat([key, ext])
  if (aes) input = Buffer.concat([input, Buffer.from('sAlT', 'latin1')])
  return md5(input).subarray(0, Math.min(key.length + 5, 16))
}

function aesDecrypt(key, data) {
  if (data.length <= 16) return Buffer.alloc(0)
  const iv = data.subarray(0, 16)
  const body = data.subarray(16)
  try {
    const d = createDecipheriv(key.length === 32 ? 'aes-256-cbc' : 'aes-128-cbc', key, iv)
    return Buffer.concat([d.update(body), d.final()])
  } catch {
    return Buffer.alloc(0)
  }
}

const CFM_AES = new Set(['AESV2', 'AESV3'])

// ---------------------------------------------------------------------------
// AES-256（R5 / R6）—— 第十二轮需求 2
// ---------------------------------------------------------------------------
//
// R5（Adobe ExtensionLevel 3）与 R6（ISO 32000-2）结构相同，只差**哈希函数**：
//   · R5：一次 SHA-256
//   · R6：Algorithm 2.B —— 迭代哈希（SHA-256/384/512，至少 64 轮，末字节 > i−32 时继续）
// 两者都由同一份 48 字节 `U`/`O` 承载：[0:32] 验证哈希，[32:40] 验证盐，[40:48] 密钥盐；
// 用户口令的密钥盐哈希解出 `UE` → 32 字节文件密钥；所有者侧同理用 `OE`。
//
// 约定差异只在**附加数据 udata** 上：R5 的用户验证不带文件 /ID，R6 带；所有者侧两边都带 `U`。
// 因此这里**逐个规范变体试**，并用 `/Perms` 的已知明文做**独立校验** —— 只有校验通过才认账，
// 所以既不会误接受错误密钥，也不靠"猜"。口令本身**永远只试调用方给的候选**（默认空串）。

const shaOf = (alg, buf) => createHash(alg).update(buf).digest()

/** Algorithm 2.B（ISO 32000-2）：迭代哈希。 */
export function hash2B(password, salt, udata = Buffer.alloc(0)) {
  const pw = Buffer.from(String(password ?? ''), 'utf8').subarray(0, 127)
  let K = shaOf('sha256', Buffer.concat([pw, salt, udata]))
  let E = Buffer.alloc(0)
  for (let i = 0; i < 64 || (E.length && E[E.length - 1] > i - 32); i++) {
    const unit = Buffer.concat([pw, K, udata])
    const block = Buffer.alloc(unit.length * 64)
    for (let r = 0; r < 64; r++) unit.copy(block, r * unit.length)
    const cipher = createCipheriv('aes-128-cbc', K.subarray(0, 16), K.subarray(16, 32))
    cipher.setAutoPadding(false)
    E = Buffer.concat([cipher.update(block), cipher.final()])
    let mod = 0
    for (let b = 0; b < 16; b++) mod += E[b]
    mod %= 3
    K = shaOf(mod === 0 ? 'sha256' : mod === 1 ? 'sha384' : 'sha512', E)
  }
  return K.subarray(0, 32)
}

/** R5 的哈希 = 一次 SHA-256；R6 = Algorithm 2.B。 */
function hashFor(R, password, salt, udata) {
  if (R >= 6) return hash2B(password, salt, udata)
  const pw = Buffer.from(String(password ?? ''), 'utf8').subarray(0, 127)
  return shaOf('sha256', Buffer.concat([pw, salt, udata]))
}

/** AES-256-CBC 解密，IV 固定 0（`UE`/`OE` 用）。 */
function aesCbcZeroIvDecrypt(key, data) {
  try {
    const d = createDecipheriv('aes-256-cbc', key, Buffer.alloc(16))
    d.setAutoPadding(false)
    return Buffer.concat([d.update(data), d.final()])
  } catch { return Buffer.alloc(0) }
}

/** `/Perms` 已知明文校验：AES-256-ECB 解密后应得到 P / 0xFFFFFFFF / T|F / adb。 */
export function checkPerms(perms, fileKey) {
  if (!perms || perms.length < 16) return true          // 没有 Perms 就不校验（老文件可能不写）
  try {
    const d = createDecipheriv('aes-256-ecb', fileKey, null)
    d.setAutoPadding(false)
    const p = Buffer.concat([d.update(perms.subarray(0, 16)), d.final()])
    if (p.length < 16) return false
    const allF = p[4] === 0xff && p[5] === 0xff && p[6] === 0xff && p[7] === 0xff
    const meta = p[8] === 0x54 || p[8] === 0x46               // 'T' / 'F'
    const adb = p[9] === 0x61 && p[10] === 0x64 && p[11] === 0x62
    return allF && meta && adb
  } catch { return false }
}

/** R5/R6 的候选密钥派生（规范里存在"带/不带文件 ID"两种约定，逐个试）。 */
function collectV5Keys(R, pw, u, o, ue, oe, perms, id0) {
  const out = []
  const variants = [id0, Buffer.alloc(0)]
  for (const udata of variants) {
    // 用户口令：验证盐在 U[32:40]，密钥盐在 U[40:48]
    if (u.length >= 48 && ue.length >= 32) {
      const key = aesCbcZeroIvDecrypt(hashFor(R, pw, u.subarray(40, 48), udata), ue.subarray(0, 32))
      if (key.length === 32) out.push({ key, role: 'user', udata })
    }
    // 所有者口令：附加数据恒为 U 的前 48 字节
    if (o.length >= 48 && oe.length >= 32) {
      const key = aesCbcZeroIvDecrypt(hashFor(R, pw, o.subarray(40, 48), u.subarray(0, 48)), oe.subarray(0, 32))
      if (key.length === 32) out.push({ key, role: 'owner', udata })
    }
  }
  // /Perms 校验通过的最优先；没有 Perms 就退回"验证哈希对得上"
  const ok = out.filter(c => checkPerms(perms, c.key))
  if (ok.length) return ok
  if (!perms || perms.length < 16) {
    return out.filter(c => {
      try {
        const expect = c.role === 'user' ? u.subarray(0, 32) : o.subarray(0, 32)
        const salt = c.role === 'user' ? u.subarray(32, 40) : o.subarray(32, 40)
        const udata = c.role === 'user' ? c.udata : u.subarray(0, 48)
        return hashFor(R, pw, salt, udata).equals(expect)
      } catch { return false }
    })
  }
  return []
}

/**
 * Build a decryptor for one /Encrypt dictionary.
 * @returns {{supported:boolean, needsPassword:boolean, note?:string, decrypt?:(bytes:Buffer,num:number,gen:number,isStream?:boolean)=>Buffer}}
 */
export function createDecryptor(enc, id0, options = {}) {
  if (!enc || typeof enc !== 'object') return { supported: false, needsPassword: false, note: '缺少 /Encrypt 字典' }
  const filter = String(enc.Filter ?? '')
  if (filter !== 'Standard') return { supported: false, needsPassword: false, note: `不支持的加密过滤器 /${filter || '?'}` }
  const R = Number(enc.R ?? 0)
  const V = Number(enc.V ?? 0)
  const lengthBits = Number(enc.Length ?? 40)
  const P = Number(enc.P ?? 0)
  const encryptMetadata = enc.EncryptMetadata !== false
  const o = Buffer.from(String(enc.O ?? ''), 'latin1')
  const u = Buffer.from(String(enc.U ?? ''), 'latin1')
  const candidates = options.passwords ?? ['']
  const v5 = R >= 5
  if (!v5 && (!o.length || !u.length)) return { supported: false, needsPassword: false, note: '加密字典缺少 O/U 字段' }
  if (v5) {
    const ue = Buffer.from(String(enc.UE ?? ''), 'latin1')
    const oe = Buffer.from(String(enc.OE ?? ''), 'latin1')
    const perms = Buffer.from(String(enc.Perms ?? ''), 'latin1')
    if (u.length < 48 || ue.length < 32) {
      return { supported: false, needsPassword: false, note: `AES-256（R${R}）的 /U /UE 字段不完整（U=${u.length} UE=${ue.length} 字节）` }
    }
    for (const pw of candidates) {
      const keys = collectV5Keys(R, pw, u, o, ue, oe, perms, id0 || Buffer.alloc(0))
      if (!keys.length) continue
      const { key, role } = keys[0]
      return {
        supported: true,
        needsPassword: false,
        aes: true,
        v5: true,
        method: `AES-256（R${R}，打开密码为空，口令角色=${role === 'user' ? '用户' : '所有者'}，Perms 校验通过）`,
        decrypt(bytes, _num, _gen, _isStream = true) {
          if (!bytes || !bytes.length) return bytes
          return aesDecrypt(key, bytes)
        },
      }
    }
    return {
      supported: false,
      needsPassword: true,
      note: `该 PDF 使用 AES-256（R${R}）加密，且**不是**空口令（/Perms 校验未通过），无法自动打开；`
        + '请在支持 AES-256 的阅读器（Acrobat Reader / Chrome / Edge）里输入密码后另存为未加密版本',
    }
  }
  const cfm = String(enc?.CF?.StdCF?.CFM ?? (V === 4 ? 'AESV2' : 'V2'))
  const aes = CFM_AES.has(cfm)
  // Only the empty user password is ever tried: a file that opens without a
  // prompt is legitimate to read; a password-protected file is reported.
  for (const pw of candidates) {
    const key = fileKey(pw, o, P, id0, R, lengthBits, encryptMetadata)
    if (userKeyMatches(key, u, id0, R)) {
      const cache = new Map()
      return {
        supported: true,
        needsPassword: false,
        aes,
        method: `${aes ? 'AES-128' : 'RC4-' + (R === 2 ? 40 : lengthBits)}（R${R}/V${V}，打开密码为空）`,
        decrypt(bytes, num, gen, isStream = true) {
          if (!bytes || !bytes.length) return bytes
          const ck = `${num}:${gen}`
          let k = cache.get(ck)
          if (!k) { k = objectKey(key, num, gen, aes); cache.set(ck, k) }
          return aes ? aesDecrypt(k, bytes) : rc4(k, bytes)
        },
      }
    }
  }
  return {
    supported: false,
    needsPassword: true,
    note: '该 PDF 设置了打开密码（不是仅限制复制/打印的权限密码），无法自动打开；请在阅读器中输入密码后另存为未加密版本',
  }
}
