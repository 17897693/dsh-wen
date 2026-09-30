// Compact, faithful XML parser/serializer used by every OOXML/ODF module.
// Produces a light tree {name, attrs, children} with text nodes as strings,
// comments and CDATA preserved as raw nodes so read→edit→write is lossless
// for untouched content.

const VOIDSELF = new Set()

const TAG_WS = new Set([' ', '\t', '\n', '\r', '\f', '\v'])
const isTagWs = ch => TAG_WS.has(ch)

/**
 * R19 任务 D：把"标签体 → {name, attrs}"从**复合体量词**正则改成手写扫描。
 * 旧正则 `/^([^\s/>]+)((?:[\s\/]+[^=>\s\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))*)\s*$/`
 * 的 `(?:…)*` 在"单个 start tag 里有上百万个属性"时抛 `RangeError: Maximum call stack size exceeded`
 * （探针 `work/r19-probe/probe-biginput-sites.mjs`；`parseXML` 是所有 OOXML / ODF 部件的入口）。
 * 逐条对齐旧语义：name = `[^\s/>]+`；分隔符 = 一个或多个空白 / `/`；
 * 属性名 = `[^=\s/>]+`；值 = `"…"` / `'…'` / `[^\s>]+`（无引号值可含 `/` 与 `=`）；尾部允许空白。
 * **任何一处不合法 ⇒ 返回 null**，调用方按旧版行为把整条标签跳过（不建节点、不落文本）。
 */
function parseTagBody(body) {
  const n = body.length
  let p = 0
  while (p < n && !isTagWs(body[p]) && body[p] !== '/' && body[p] !== '>') p++
  if (p === 0) return null
  const name = body.slice(0, p)
  const attrs = {}
  while (p < n) {
    const sepAt = p
    while (p < n && (isTagWs(body[p]) || body[p] === '/')) p++
    if (p >= n) break
    if (p === sepAt) return null            // 属性之间必须有分隔符
    const ks = p
    while (p < n && !isTagWs(body[p]) && body[p] !== '/' && body[p] !== '=' && body[p] !== '>') p++
    if (p === ks) return null               // 属性名不得为空
    const key = body.slice(ks, p)
    while (p < n && isTagWs(body[p])) p++
    if (body[p] !== '=') return null        // 旧正则要求每个属性都带 `=`
    p++
    while (p < n && isTagWs(body[p])) p++
    let val
    if (body[p] === '"' || body[p] === "'") {
      const q = body[p]
      p++
      const vs = p
      while (p < n && body[p] !== q) p++
      if (p >= n) return null               // 引号未闭合（旧正则的 `"[^"]*"` 要求闭合）
      val = body.slice(vs, p)
      p++
    } else {
      const vs = p
      while (p < n && !isTagWs(body[p]) && body[p] !== '>') p++
      if (p === vs) return null
      val = body.slice(vs, p)
    }
    attrs[key] = decodeEntities(val)
  }
  return { name, attrs }
}

export function parseXML(text) {
  const root = { name: '#document', attrs: {}, children: [] }
  const stack = [root]
  let i = 0
  const n = text.length
  const top = () => stack[stack.length - 1]

  while (i < n) {
    const lt = text.indexOf('<', i)
    if (lt < 0) { addText(text.slice(i)); break }
    if (lt > i) addText(text.slice(i, lt))
    if (text.startsWith('<!--', lt)) {
      const end = text.indexOf('-->', lt + 4)
      const stop = end < 0 ? n : end + 3
      top().children.push({ raw: text.slice(lt, stop) })
      i = stop
      continue
    }
    if (text.startsWith('<![CDATA[', lt)) {
      const end = text.indexOf(']]>', lt + 9)
      const stop = end < 0 ? n : end + 3
      top().children.push({ cdata: text.slice(lt + 9, end < 0 ? n : end) })
      i = stop
      continue
    }
    if (text.startsWith('<?', lt)) {
      const end = text.indexOf('?>', lt + 2)
      i = end < 0 ? n : end + 2
      continue
    }
    if (text.startsWith('<!', lt)) { // DOCTYPE etc.
      const end = text.indexOf('>', lt + 2)
      i = end < 0 ? n : end + 1
      continue
    }
    if (text[lt + 1] === '/') {
      const end = text.indexOf('>', lt)
      const name = text.slice(lt + 2, end).trim()
      // pop to matching open tag (tolerate mismatch gracefully)
      for (let s = stack.length - 1; s > 0; s--) {
        if (stack[s].name === name) { stack.length = s; break }
      }
      i = end + 1
      continue
    }
    // open (or self-closing) tag: scan to its '>' respecting quotes
    let j = lt + 1
    while (j < n) {
      const c = text[j]
      if (c === '"' || c === "'") { const q = c; j++; while (j < n && text[j] !== q) j++ }
      else if (c === '>') break
      j++
    }
    const tagText = text.slice(lt + 1, j)
    i = j + 1
    if (j >= n) { addText(text.slice(lt)); continue }
    const selfClose = tagText.endsWith('/')
    const body = selfClose ? tagText.slice(0, -1) : tagText
    // R19 任务 D：这里原先是一条吃下**整个属性序列**的复合体量词正则 + 第二遍全局 `exec`，
    // 现在合成一次手写扫描（见 parseTagBody）。
    const parsed = parseTagBody(body)
    if (!parsed) continue
    const node = { name: parsed.name, attrs: parsed.attrs, children: [] }
    top().children.push(node)
    if (!selfClose && !VOIDSELF.has(node.name)) stack.push(node)
  }
  function addText(t) { if (t) top().children.push(decodeEntities(t)) }
  return root
}

export function serializeXML(node) {
  if (typeof node === 'string') return encodeEntities(node)
  if (node.raw) return node.raw
  if (node.cdata !== undefined) return `<![CDATA[${node.cdata}]]>`
  if (node.name === '#document') {
    let out = ''
    for (const c of node.children) out += serializeXML(c)
    return out
  }
  let attrs = ''
  for (const [k, v] of Object.entries(node.attrs)) attrs += ` ${k}="${encodeEntities(v, true)}"`
  if (!node.children.length) return `<${node.name}${attrs}/>`
  let inner = ''
  for (const c of node.children) inner += serializeXML(c)
  return `<${node.name}${attrs}>${inner}</${node.name}>`
}

const NAMED = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' }
export function decodeEntities(s) {
  if (!s.includes('&')) return s
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (all, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10)
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all
    }
    return NAMED[body] ?? all
  })
}

export function encodeEntities(s, isAttr = false) {
  let out = String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
  if (isAttr) out = out.replace(/"/g, '&quot;').replace(/\r?\n/g, '&#10;')
  return out
}

/** All descendant elements with the given tag (exact, or array of names). */
export function findAll(node, name) {
  const wanted = Array.isArray(name) ? new Set(name) : new Set([name])
  const out = []
  const walk = n => {
    if (typeof n === 'string' || n.raw || n.cdata) return
    if (wanted.has(n.name)) out.push(n)
    for (const c of n.children) walk(c)
  }
  for (const c of node.children) walk(c)
  return out
}

/** First matching descendant. */
export function find(node, name) {
  const wanted = Array.isArray(name) ? new Set(name) : new Set([name])
  const walk = n => {
    if (typeof n === 'string' || n.raw || n.cdata) return undefined
    if (wanted.has(n.name)) return n
    for (const c of n.children) { const hit = walk(c); if (hit) return hit }
    return undefined
  }
  for (const c of node.children) { const hit = walk(c); if (hit) return hit }
  return undefined
}

/** Direct children (optionally filtered by tag). */
export function children(node, name) {
  const out = []
  for (const c of node.children) {
    if (typeof c === 'string' || c.raw || c.cdata) continue
    if (!name || c.name === name) out.push(c)
  }
  return out
}

/** Concatenated text of all descendant text nodes (entity-decoded). */
export function textOf(node) {
  if (typeof node === 'string') return node
  if (node.cdata !== undefined) return node.cdata
  if (node.raw) return ''
  let t = ''
  for (const c of node.children) t += textOf(c)
  return t
}
