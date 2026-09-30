// dsh-office — zero-dependency HTML reader (需求 1：HTML 读取与转换).
//
// 为什么不用正则剥标签：实测中正则会丢嵌套列表语义（<ul> 套 <ol> 的层级一旦压平，
// 语义顺序就断了）。这里实现一个"规范精简版"的 HTML 分词器 + 树构造器 +
// 文档模型遍历器，三段都只依赖 Node 内置（零第三方依赖，许可证 = 本插件同款私有）：
//
//   1. tokenize()      —— 词法：标签 / 注释 / DOCTYPE / 原始文本元素（script/style/textarea/title）
//   2. buildTree()     —— 语法：隐式闭标签（li/p/tr/td/option…）、void 元素、大小写无关
//   3. toDocumentModel() —— 语义：h1-h6→heading、ul/ol→list(嵌套→level)、table→table、
//                         blockquote→quote、pre→code、a→link run、strong/em→样式 run；
//                         style/script/表单控件/隐藏子树跳过；HTML 实体（含 named/数字/
//                         Windows-1252 数字别名）统一解码；空白合并但保留段落边界。
//
// 输出是插件统一的 document 内容模型（与 docx/odt/md 同构），因此 office_read 的
// as="markdown"/"json"/"meta" 与 office_convert 的全部目标格式（.md/.txt/.docx/.pdf/…）
// 都自动继承。as="text" 用 plainTextOf() 给真正的纯文本（表格→制表符行，不残留 | 管道）。
//
// 已知边界（有意不做，避免误伤）：CSS 伪元素（content:attr(data-no) 之类）生成的内容
// 不在 DOM 里，转换不包含——那是排版层而不是内容层；colspan/rowspan 只取单元格文本，
// 不做网格展开；<ol start="N"> 的起始编号在 markdown 渲染层无法表达（list 模型只有
// ordered 布尔）。

// ---------------------------------------------------------------------------
// HTML 实体解码：named（HTML4 常用集）+ 数字（含 Windows-1252 数字别名）。
// 不用 xml.js 的 decodeEntities —— 那张 named 表只有 5 项，而网页笔记里
// &nbsp; &mdash; &ldquo; “中文页面常用符号” 出现频率很高，解码不全就是语义损失。
// ---------------------------------------------------------------------------

const ENTITIES = {
  // 结构与基础
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0', shy: '\u00ad',
  // Latin-1 补充（印刷符号）
  iexcl: '¡', cent: '¢', pound: '£', curren: '¤', yen: '¥', brvbar: '¦', sect: '§',
  uml: '¨', copy: '©', ordf: 'ª', laquo: '«', not: '¬', reg: '®', macr: '¯',
  deg: '°', plusmn: '±', sup2: '²', sup3: '³', acute: '´', micro: 'µ', para: '¶',
  middot: '·', cedil: '¸', sup1: '¹', ordm: 'º', raquo: '»', frac14: '¼',
  frac12: '½', frac34: '¾', iquest: '¿', times: '×', divide: '÷',
  // Latin-1 大写字母
  Agrave: 'À', Aacute: 'Á', Acirc: 'Â', Atilde: 'Ã', Auml: 'Ä', Aring: 'Å',
  AElig: 'Æ', Ccedil: 'Ç', Egrave: 'È', Eacute: 'É', Ecirc: 'Ê', Euml: 'Ë',
  Igrave: 'Ì', Iacute: 'Í', Icirc: 'Î', Iuml: 'Ï', ETH: 'Ð', Ntilde: 'Ñ',
  Ograve: 'Ò', Oacute: 'Ó', Ocirc: 'Ô', Otilde: 'Õ', Ouml: 'Ö', Oslash: 'Ø',
  Ugrave: 'Ù', Uacute: 'Ú', Ucirc: 'Û', Uuml: 'Ü', Yacute: 'Ý', THORN: 'Þ',
  szlig: 'ß',
  // Latin-1 小写字母
  agrave: 'à', aacute: 'á', acirc: 'â', atilde: 'ã', auml: 'ä', aring: 'å',
  aelig: 'æ', ccedil: 'ç', egrave: 'è', eacute: 'é', ecirc: 'ê', euml: 'ë',
  igrave: 'ì', iacute: 'í', icirc: 'î', iuml: 'ï', eth: 'ð', ntilde: 'ñ',
  ograve: 'ò', oacute: 'ó', ocirc: 'ô', otilde: 'õ', ouml: 'ö', oslash: 'ø',
  ugrave: 'ù', uacute: 'ú', ucirc: 'û', uuml: 'ü', yacute: 'ý', thorn: 'þ',
  yuml: 'ÿ',
  // Latin Extended-A（常用）
  OElig: 'Œ', oelig: 'œ', Scaron: 'Š', scaron: 'š', Yuml: 'Ÿ', fnof: 'ƒ',
  // 间距与标点
  ensp: '\u2002', emsp: '\u2003', thinsp: '\u2009', zwnj: '\u200c', zwj: '\u200d',
  lrm: '\u200e', rlm: '\u200f', ndash: '–', mdash: '—', lsquo: '‘', rsquo: '’',
  sbquo: '‚', ldquo: '“', rdquo: '”', bdquo: '„', dagger: '†', Dagger: '‡',
  bull: '•', hellip: '…', permil: '‰', prime: '′', Prime: '″', lsaquo: '‹',
  rsaquo: '›', oline: '‾', frasl: '⁄', euro: '€', trade: '™',
  // 字母符号
  image: 'ℑ', weierp: '℘', real: 'ℜ', alefsym: 'ℵ',
  // 箭头
  larr: '←', uarr: '↑', rarr: '→', darr: '↓', harr: '↔', crarr: '↵',
  lArr: '⇐', uArr: '⇑', rArr: '⇒', dArr: '⇓', hArr: '⇔',
  // 数学
  forall: '∀', part: '∂', exist: '∃', empty: '∅', nabla: '∇', isin: '∈',
  notin: '∉', ni: '∋', prod: '∏', sum: '∑', minus: '−', lowast: '∗',
  radic: '√', prop: '∝', infin: '∞', ang: '∠', and: '∧', or: '∨',
  cap: '∩', cup: '∪', int: '∫', there4: '∴', sim: '∼', cong: '≅',
  asymp: '≈', ne: '≠', equiv: '≡', le: '≤', ge: '≥',
  // 希腊字母
  Alpha: 'Α', Beta: 'Β', Gamma: 'Γ', Delta: 'Δ', Epsilon: 'Ε', Zeta: 'Ζ',
  Eta: 'Η', Theta: 'Θ', Iota: 'Ι', Kappa: 'Κ', Lambda: 'Λ', Mu: 'Μ',
  Nu: 'Ν', Xi: 'Ξ', Omicron: 'Ο', Pi: 'Π', Rho: 'Ρ', Sigma: 'Σ', Tau: 'Τ',
  Upsilon: 'Υ', Phi: 'Φ', Chi: 'Χ', Psi: 'Ψ', Omega: 'Ω',
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ',
  eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ',
  nu: 'ν', xi: 'ξ', omicron: 'ο', pi: 'π', rho: 'ρ', sigmaf: 'ς',
  sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ',
  omega: 'ω', thetasym: 'ϑ', upsih: 'ϒ', piv: 'ϖ',
  // 形状
  loz: '◊', spades: '♠', clubs: '♣', hearts: '♥', diams: '♦',
}

/** 浏览器行为：数字实体 0x80–0x9F 按 Windows-1252 解释（&#150; → – 而不是控制符）。 */
const CP1252_ALIASES = {
  0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡',
  0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ', 0x8e: 'Ž',
  0x91: '‘', 0x92: '’', 0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—',
  0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›', 0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ',
}

export function decodeEntities(s) {
  if (!s || !s.includes('&')) return s
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (all, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10)
      if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return all
      if (CP1252_ALIASES[code]) return CP1252_ALIASES[code]
      if ((code >= 0xd800 && code <= 0xdfff) || (code >= 0x7f && code <= 0x9f)) return '\ufffd'
      return String.fromCodePoint(code)
    }
    return ENTITIES[body] ?? all
  })
}

// ---------------------------------------------------------------------------
// 1. 词法：tokenize
// ---------------------------------------------------------------------------

/** 内容当"原始文本"处理的元素：内部不识别标签，只找匹配的闭标签。 */
const RAW_TEXT = new Set(['script', 'style', 'textarea', 'title', 'xmp', 'noembed', 'noframes'])

const VOID = new Set([
  'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta',
  'param', 'source', 'track', 'wbr', 'basefont', 'bgsound', 'frame', 'keygen',
])

function tokenize(html) {
  const tokens = []
  const n = html.length
  let i = 0
  while (i < n) {
    const lt = html.indexOf('<', i)
    if (lt < 0) { if (i < n) tokens.push({ type: 'text', data: decodeEntities(html.slice(i)) }); break }
    if (lt > i) tokens.push({ type: 'text', data: decodeEntities(html.slice(i, lt)) })
    if (html.startsWith('<!--', lt)) {
      const end = html.indexOf('-->', lt + 4)
      tokens.push({ type: 'comment', data: end < 0 ? html.slice(lt + 4) : html.slice(lt + 4, end) })
      i = end < 0 ? n : end + 3
      continue
    }
    if (html.startsWith('<!', lt) || html.startsWith('<?', lt)) {   // DOCTYPE / 处理指令 / 假注释
      const end = html.indexOf('>', lt + 2)
      i = end < 0 ? n : end + 1
      continue
    }
    const isClose = html[lt + 1] === '/'
    // 标签体：扫到 '>'，引号里的 '>' 不算结束
    let j = lt + (isClose ? 2 : 1)
    while (j < n) {
      const c = html[j]
      if (c === '"' || c === "'") { const q = c; j++; while (j < n && html[j] !== q) j++ }
      else if (c === '>') break
      j++
    }
    const body = html.slice(lt + (isClose ? 2 : 1), j)
    i = j + 1
    const nameMatch = /^[a-zA-Z][a-zA-Z0-9:._-]*/.exec(body)
    if (!nameMatch) {
      // `<3` 这类不是标签：按文本回收（宁可留噪不删正文）
      tokens.push({ type: 'text', data: html.slice(lt, Math.min(j + 1, n)) })
      continue
    }
    const name = nameMatch[0].toLowerCase()
    if (isClose) { tokens.push({ type: 'close', name }); continue }
    const attrs = {}
    const attrRe = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]*)))?/g
    let rest = body.slice(nameMatch[0].length)
    if (rest.endsWith('/')) rest = rest.slice(0, -1)
    let am
    while ((am = attrRe.exec(rest))) {
      const key = am[1].toLowerCase()
      if (key in attrs) continue
      attrs[key] = decodeEntities(am[2] ?? am[3] ?? am[4] ?? '')
    }
    tokens.push({ type: 'open', name, attrs })
    if (RAW_TEXT.has(name)) {
      // 原始文本元素：找到大小写无关的 </name 才算结束（内容不做实体解码、不识别标签）
      const closeRe = new RegExp(`</${name}\\s*>`, 'i')
      const m = closeRe.exec(html.slice(i))
      const raw = m ? html.slice(i, i + m.index) : html.slice(i)
      tokens.push({ type: 'rawtext', name, data: raw })
      i += raw.length
      const cm = /^<\/([a-zA-Z][a-zA-Z0-9:._-]*)/.exec(html.slice(i))
      if (cm) {
        tokens.push({ type: 'close', name: cm[1].toLowerCase() })
        const gt = html.indexOf('>', i)
        i = gt < 0 ? n : gt + 1
      }
    } else if (VOID.has(name)) {
      tokens.push({ type: 'close', name, implied: true })
    }
  }
  return tokens
}

// ---------------------------------------------------------------------------
// 2. 语法：buildTree（隐式闭标签 + 大小写无关的配对）
// ---------------------------------------------------------------------------

/** 这些块级元素开始时，会自动关掉还开着的 <p>（HTML5 "close a p element" 规则的精简版）。 */
const CLOSES_P = new Set([
  'address', 'article', 'aside', 'blockquote', 'details', 'dialog', 'dir', 'div',
  'dl', 'fieldset', 'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3',
  'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'main', 'menu', 'nav', 'ol', 'p',
  'pre', 'section', 'table', 'ul', 'li', 'dt', 'dd',
])

/** 新开的元素会自动关掉的同名/同类开元素（列容器为界）。 */
const IMPLIED_END = {
  li: { same: ['li'], scope: ['ul', 'ol'] },
  dt: { same: ['dt', 'dd'], scope: ['dl'] },
  dd: { same: ['dt', 'dd'], scope: ['dl'] },
  tr: { same: ['tr'], scope: ['table', 'thead', 'tbody', 'tfoot'] },
  td: { same: ['td', 'th'], scope: ['tr', 'table'] },
  th: { same: ['td', 'th'], scope: ['tr', 'table'] },
  option: { same: ['option'], scope: ['select', 'datalist'] },
  optgroup: { same: ['optgroup', 'option'], scope: ['select'] },
  thead: { same: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th'], scope: ['table'] },
  tbody: { same: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th'], scope: ['table'] },
  tfoot: { same: ['thead', 'tbody', 'tfoot', 'tr', 'td', 'th'], scope: ['table'] },
  a: { same: ['a'], scope: [] },
}

function buildTree(tokens) {
  const root = { name: '#root', attrs: {}, children: [] }
  const stack = [root]
  const top = () => stack[stack.length - 1]
  // <p> 的作用域边界：越过这些容器还没找到开着的 <p>，就当没有（HTML5 "button scope" 精简版）
  const P_SCOPE = new Set(['ul', 'ol', 'table', 'td', 'th', 'caption', 'select', 'button', 'object', 'template', 'html', 'body', '#root'])
  const findOpenP = () => {
    for (let s = stack.length - 1; s > 0; s--) {
      const nm = stack[s].name
      if (nm === 'p') return s
      if (P_SCOPE.has(nm)) return -1
    }
    return -1
  }
  for (const t of tokens) {
    if (t.type === 'text') { top().children.push(t.data); continue }
    if (t.type === 'rawtext') {
      // script/style 的内容整体丢弃；textarea/title 是真实正文（解码实体）
      if (t.name === 'script' || t.name === 'style') continue
      top().children.push(decodeEntities(t.data))
      continue
    }
    if (t.type === 'comment') continue
    if (t.type === 'close') {
      if (t.implied) continue                       // void 元素的合成闭标签：不入栈，也无事可做
      if (t.name === 'p') {
        const s = findOpenP()
        if (s > 0) stack.length = s
        continue
      }
      for (let s = stack.length - 1; s > 0; s--) {
        if (stack[s].name === t.name) { stack.length = s; break }
      }
      continue
    }
    // open
    const implied = IMPLIED_END[t.name]
    if (implied) {
      for (let s = stack.length - 1; s > 0; s--) {
        const nm = stack[s].name
        if (implied.same.includes(nm)) { stack.length = s; break }
        if (implied.scope.includes(nm)) break
      }
    }
    if (CLOSES_P.has(t.name)) {
      const s = findOpenP()
      if (s > 0) stack.length = s
    }
    const node = { name: t.name, attrs: t.attrs, children: [] }
    top().children.push(node)
    if (!VOID.has(t.name)) stack.push(node)
  }
  return root
}

// ---------------------------------------------------------------------------
// 3. 语义：文档模型遍历
// ---------------------------------------------------------------------------

/** 整个子树跳过（噪声/非正文/不可渲染）。 */
const SKIP_SUBTREE = new Set([
  'script', 'style', 'noscript', 'template', 'svg', 'canvas', 'iframe',
  'object', 'embed', 'audio', 'video', 'select', 'option', 'optgroup',
  'input', 'textarea', 'button', 'label', 'datalist', 'output', 'dialog',
  'datalist', 'map', 'head', 'colgroup', 'col',
])

const INLINE_STYLE = {
  strong: { bold: true }, b: { bold: true },
  em: { italic: true }, i: { italic: true }, cite: { italic: true },
  var: { italic: true }, dfn: { italic: true },
  code: { code: true }, kbd: { code: true }, samp: { code: true }, tt: { code: true },
  u: { underline: true }, ins: { underline: true },
  s: { strike: true }, strike: { strike: true }, del: { strike: true },
}

const isHidden = node => node.attrs && (
  node.attrs.hidden !== undefined
  || /display\s*:\s*none|visibility\s*:\s*hidden/i.test(String(node.attrs.style || ''))
  || node.attrs['aria-hidden'] === 'true'
)

const collapse = s => String(s ?? '').replace(/[ \t\r\n\f]+/g, ' ')

/** 子树的纯文本（块边界转空格；quote 用换行版本由调用方处理）。 */
function innerTextOf(node, sep = ' ') {
  let out = ''
  const walk = n => {
    if (typeof n === 'string') { out += n; return }
    if (n.name === 'br') { out += ' '; return }
    for (const c of n.children) walk(c)
    if (BLOCK_LEVEL.has(n.name)) out += sep
  }
  walk(node)
  return collapse(out).trim()
}

/** <pre> 子树文本：保留原样（含嵌套 <code>/<span> 内的文本与 <br> 换行），不折叠空白。 */
function preTextOf(node) {
  let out = ''
  const walk = n => {
    if (typeof n === 'string') { out += n; return }
    if (n.name === 'br') { out += '\n'; return }
    for (const c of n.children) walk(c)
  }
  walk(node)
  return out
}

const BLOCK_LEVEL = new Set([
  'address', 'article', 'aside', 'blockquote', 'details', 'div', 'dl', 'dt',
  'dd', 'fieldset', 'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3',
  'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'li', 'main', 'nav', 'ol', 'p',
  'pre', 'section', 'summary', 'table', 'ul',
])

/**
 * HTML 树 → 统一 document 内容模型。
 * 结构：blocks 数组 + "当前段落"累积器。任何块级元素出现时先冲刷段落；
 * 文本/行内元素累积进当前段落（空白合并、样式 run 合并），段落边界由块级元素保证。
 */
function toDocumentModel(tree) {
  const meta = {}
  const blocks = []
  let para = null                       // { runs: [...] } 或 null

  const flushPara = () => {
    if (!para) return
    const runs = para.runs
    para = null
    // 掐头去尾的普通空格；全空白段落丢弃
    if (runs.length) {
      runs[0] = { ...runs[0], text: String(runs[0].text ?? '').replace(/^ +/, '') }
      const last = runs[runs.length - 1]
      runs[runs.length - 1] = { ...last, text: String(last.text ?? '').replace(/ +$/, '') }
    }
    // 掐完可能出空 run（换行 run 除外——那是硬换行，有语义）
    const kept = runs.filter(r => String(r.text ?? '').length && (String(r.text).trim() || String(r.text).includes('\n')))
    if (kept.length) blocks.push({ type: 'paragraph', runs: kept })
  }

  const appendText = (text, style, runs) => {
    const t = collapse(text)
    if (!t.trim() && (!runs || !runs.length)) return   // 块间的纯空白不产生段落/空 run
    const prev = runs[runs.length - 1]
    const same = prev && ['bold', 'italic', 'code', 'underline', 'strike', 'link']
      .every(k => (prev[k] ?? null) === (style[k] ?? null))
    if (same) {
      // 合并连续同款 run；前文结尾是空格/换行而新文以空格开头时不重复
      prev.text += (/[ \n]$/.test(prev.text) && t.startsWith(' ') ? t.slice(1) : t)
    } else {
      runs.push({ ...style, text: t })
    }
  }

  const hrefOf = node => {
    const href = String(node.attrs?.href || '').trim()
    return /^https?:/i.test(href) ? href : ''   // 锚点/mailto/javascript 属导航噪声，不进正文
  }

  /** 当前段落的 runs（保证 para 存在）。行内遍历每个文本节点都要重新解析——
   *  段落内的 <img> 会把段落冲刷掉并新起一段，持有旧数组引用会把图后正文写进孤儿数组。 */
  const curRuns = () => {
    if (!para) para = { runs: [] }
    return para.runs
  }

  /** 行内遍历：把子树累积进目标 runs（缺省 = 当前段落，逐节点解析）。 */
  const walkInline = (node, style, runs = null) => {
    for (const c of node.children) {
      if (typeof c === 'string') { appendText(c, style, runs ?? curRuns()); continue }
      if (c.name === 'br') { appendText('\n', style, runs ?? curRuns()); continue }
      if (c.name === 'img') { pushImage(c); continue }
      if (SKIP_SUBTREE.has(c.name) || isHidden(c)) continue
      const st = { ...style }
      if (c.name === 'a') {
        const href = hrefOf(c)
        if (href) st.link = href
      }
      const extra = INLINE_STYLE[c.name]
      if (extra) Object.assign(st, extra)
      walkInline(c, st, runs)
    }
  }

  /** 图片块：先把段落冲刷掉（图片是独立块），图后正文另起一段继续累积。 */
  const pushImage = node => {
    flushPara()
    const alt = collapse(node.attrs?.alt || '').trim()
    const src = String(node.attrs?.src || '').trim()
    blocks.push({ type: 'image', alt: alt || '图片', name: /^data:/i.test(src) ? '' : src })
    if (!para) para = { runs: [] }
  }

  /** 列表：嵌套 ul/ol 压平进同一个 list 块（level 记层级，语义顺序=文档顺序）。 */
  const parseList = (node, level, ordered) => {
    const items = []
    for (const c of node.children) {
      if (typeof c === 'string' || c.name === 'script' || c.name === 'style') continue
      if (c.name === 'li') {
        let text = ''
        for (const g of c.children) {
          if (typeof g === 'string') { text += g; continue }
          if (g.name === 'ul' || g.name === 'ol') continue   // 嵌套列表递归处理
          if (g.name === 'img') { text += ` [图片${collapse(g.attrs?.alt || '') ? '：' + collapse(g.attrs.alt).trim() : ''}] `; continue }
          text += ' ' + innerTextOf(g)
        }
        text = collapse(text).trim()
        if (text) items.push({ level, text })
        for (const g of c.children) {
          if (g.name === 'ul' || g.name === 'ol') items.push(...parseList(g, level + 1, ordered))
        }
      } else if (c.name === 'ul' || c.name === 'ol') {
        items.push(...parseList(c, level, ordered))          // 容错：ul 直套 ul
      }
    }
    return items
  }

  const parseTable = node => {
    const rows = []
    let caption = ''
    for (const c of node.children) {
      if (typeof c === 'string') continue
      if (c.name === 'caption') { caption = innerTextOf(c); continue }
      if (c.name === 'colgroup' || c.name === 'col') continue
      collectRows(c, rows)
    }
    const header = rows.length > 0 && rows[0].every(cell => cell.th)
    return {
      block: { type: 'table', header, rows: rows.map(r => r.map(c => c.text)) },
      caption,
    }
  }
  /** 行收集：node 可能是 tr 本身，也可能是 thead/tbody/tfoot/table（或包裹它们的容器）。 */
  const collectRows = (node, rows) => {
    if (node.name === 'tr') { addRow(node, rows); return }
    for (const c of node.children) {
      if (typeof c === 'string') continue
      collectRows(c, rows)             // tr → 加行；thead/tbody/tfoot/包裹容器 → 递归
    }
  }
  const addRow = (trNode, rows) => {
    const cells = []
    for (const g of trNode.children) {
      if (typeof g === 'string') continue
      if (g.name === 'td' || g.name === 'th') cells.push({ th: g.name === 'th', text: cellText(g) })
      else if (g.name === 'tr') collectRows(g, rows)   // 容错：td 直套 tr
    }
    if (cells.length) rows.push(cells)
  }
  const cellText = node => {
    let out = ''
    for (const c of node.children) {
      if (typeof c === 'string') { out += c; continue }
      if (c.name === 'table') { out += ' ' + tablePlainText(c) + ' '; continue }   // 嵌套表：压成文本
      if (c.name === 'img') { out += ` [图片${collapse(c.attrs?.alt || '') ? '：' + collapse(c.attrs.alt).trim() : ''}] `; continue }
      if (c.name === 'br') { out += ' '; continue }
      out += cellText(c)
      if (BLOCK_LEVEL.has(c.name)) out += ' '
    }
    return collapse(out).replace(/^ +| +$/g, '').trim()
  }
  const tablePlainText = node => parseTable(node).block.rows.map(r => r.join(' / ')).join(' / ')

  /** 块级遍历。 */
  const walkBlock = node => {
    for (const c of node.children) {
      if (typeof c === 'string') { appendText(c, {}, curRuns()); continue }
      if (isHidden(c)) continue
      const name = c.name
      if (name === 'head') {
        // <head> 子树只回收 <title>（进 meta），其余（style/script/meta/link）跳过
        for (const g of c.children) {
          if (typeof g !== 'string' && g.name === 'title') {
            meta.title = collapse(g.children.filter(x => typeof x === 'string').join('')).trim()
          }
        }
        continue
      }
      if (SKIP_SUBTREE.has(c.name)) continue
      if (/^h[1-6]$/.test(name)) {
        flushPara()
        const runs = []
        walkInline(c, {}, runs)
        blocks.push({ type: 'heading', level: Number(name[1]), text: collapse(runs.map(r => r.text).join('')).trim() })
        continue
      }
      if (name === 'p' || name === 'figcaption' || name === 'legend' || name === 'summary') {
        flushPara()
        walkInline(c, {})
        flushPara()
        continue
      }
      if (name === 'ul' || name === 'ol') {
        flushPara()
        const items = parseList(c, 0, name === 'ol')
        if (items.length) blocks.push({ type: 'list', ordered: name === 'ol', items })
        continue
      }
      if (name === 'dl') {
        flushPara()
        const items = []
        let level = 0
        for (const g of c.children) {
          if (typeof g === 'string') continue
          if (g.name === 'dt') { items.push({ level: 0, text: innerTextOf(g) }) ; level = 1 }
          else if (g.name === 'dd') items.push({ level, text: innerTextOf(g) })
        }
        if (items.length) blocks.push({ type: 'list', ordered: false, items: items.filter(i => i.text) })
        continue
      }
      if (name === 'li') {  // 容错：游离 li
        flushPara()
        appendText(innerTextOf(c), {}, curRuns())
        flushPara()
        continue
      }
      if (name === 'table') {
        flushPara()
        const { block, caption } = parseTable(c)
        if (caption) blocks.push({ type: 'paragraph', runs: [{ text: caption, bold: true }] })
        if (block.rows.length) blocks.push(block)
        continue
      }
      if (name === 'blockquote' || name === 'q') {
        flushPara()
        const text = innerTextOf(c, '\n')
        if (text) blocks.push({ type: 'quote', text })
        continue
      }
      if (name === 'pre') {
        flushPara()
        blocks.push({ type: 'code', text: preTextOf(c).replace(/^\n+|\s+$/g, '') })
        continue
      }
      if (name === 'hr') { flushPara(); blocks.push({ type: 'hr' }); continue }
      if (name === 'img') { pushImage(c); continue }
      if (name === 'br') { appendText('\n', {}, curRuns()); continue }
      if (INLINE_STYLE[name] || name === 'a' || name === 'span' || name === 'font'
        || name === 'small' || name === 'big' || name === 'abbr' || name === 'time'
        || name === 'mark' || name === 'sub' || name === 'sup' || name === 'bdi'
        || name === 'bdo' || name === 'data' || name === 'wbr') {
        // 块级语境里撞见行内元素：按行内处理（罕见但合法）
        walkInline({ children: [c] }, {})
        continue
      }
      // 其余（div/section/article/main/header/footer/aside/nav/figure/body…）：
      // 透明容器 —— 先冲刷（块边界），递归子树。
      flushPara()
      walkBlock(c)
      flushPara()
    }
  }

  walkBlock(tree)
  flushPara()
  return { meta, blocks }
}

// ---------------------------------------------------------------------------
// 对外入口
// ---------------------------------------------------------------------------

/** HTML 源文本 → 统一 document 内容模型（office_read / office_convert 的读取端）。 */
export function htmlToDocument(html) {
  const tree = buildTree(tokenize(String(html ?? '')))
  const { meta, blocks } = toDocumentModel(tree)
  return { kind: 'document', meta, blocks }
}

/** HTML 源文本 → 纯文本（as="text" 用）：块边界换行、列表缩进、表格制表符分隔。 */
export function htmlPlainText(html) {
  return plainTextOf(htmlToDocument(html))
}

/** document 模型 → 纯文本（与 documentToMarkdown 不同：不残留 #、|、** 等标记）。 */
export function plainTextOf(doc) {
  const out = []
  for (const b of doc?.blocks || []) {
    switch (b.type) {
      case 'heading': out.push(String(b.text ?? '')); break
      case 'paragraph': out.push((b.runs || []).map(r => String(r.text ?? '')).join('')); break
      case 'list': {
        // 列表是一个块：条目之间只用换行（不空行），嵌套用缩进表达
        const lines = (b.items || []).map(it =>
          `${'    '.repeat(Math.max(0, Math.min(6, it.level || 0)))}- ${it.text}`)
        if (lines.length) out.push(lines.join('\n'))
        break
      }
      case 'table':
        out.push((b.rows || []).map(r => (r || []).join('\t')).join('\n'))
        break
      case 'code': out.push(String(b.text ?? '')); break
      case 'quote': out.push(String(b.text ?? '')); break
      case 'hr': out.push('──────────'); break
      case 'image': out.push(`[图片：${b.alt || ''}]`); break
      default: if (b.text) out.push(String(b.text))
    }
  }
  return out.join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

// ---------------------------------------------------------------------------
// 写出端：统一 document 内容模型 → 语义化 HTML5
// （office_create / office_convert 的 .html 目标；读取端 htmlToDocument 与它互为往返）
// ---------------------------------------------------------------------------

const HTML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }

/** 文本 → HTML 转义（中文/emoji 一律原样输出，不转数字实体）。 */
export function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => HTML_ESCAPES[c])
}

const clampLevel = lv => Math.max(0, Math.min(8, Number(lv) || 0))

/** 行内 run 序列 → HTML（强调用语义标签，链接用 <a>，换行用 <br>）。 */
function runsToHtml(runs) {
  return (runs || []).map(r => {
    const run = typeof r === 'string' ? { text: r } : r
    let s = escapeHtml(run.text).replace(/\n/g, '<br>')
    if (run.code) s = `<code>${s}</code>`
    if (run.bold) s = `<strong>${s}</strong>`
    if (run.italic) s = `<em>${s}</em>`
    if (run.underline) s = `<u>${s}</u>`
    if (run.strike) s = `<s>${s}</s>`
    if (run.link) s = `<a href="${escapeHtml(run.link)}">${s}</a>`
    return s
  }).join('')
}

/**
 * 扁平 items（带 level）→ 嵌套 <ul>/<ol>。
 * 层级上升时新列表开在**当前未闭合的 <li> 内部**；下降时逐层收掉 </li></列表>。
 */
function listToHtml(items, ordered) {
  const tag = ordered ? 'ol' : 'ul'
  let html = ''
  let cur = -1
  for (const it of items || []) {
    const lvl = clampLevel(it.level)
    if (lvl > cur) {
      for (let d = cur + 1; d <= lvl; d++) html += `<${tag}>`
    } else {
      if (cur >= 0) html += '</li>'
      for (let d = cur; d > lvl; d--) html += `</${tag}>`
    }
    html += `<li>${runsToHtml([{ text: it.text }])}`
    cur = lvl
  }
  while (cur >= 0) { html += `</li></${tag}>`; cur-- }
  return html
}

function tableToHtml(rows, header) {
  const grid = Math.max(1, ...(rows || []).map(r => (r || []).length))
  const cell = (v, th) => {
    const body = runsToHtml([{ text: String(v ?? '') }])
    return `<${th ? 'th' : 'td'}>${body}</${th ? 'th' : 'td'}>`
  }
  const row = (r, th) => `<tr>${Array.from({ length: grid }, (_, i) => cell((r || [])[i], th)).join('')}</tr>`
  const head = header && (rows || []).length ? `<thead>${row(rows[0], true)}</thead>` : ''
  const body = (header ? (rows || []).slice(1) : (rows || [])).map(r => row(r, false)).join('')
  return `<table>${head}<tbody>${body}</tbody></table>`
}

/**
 * 统一 document 内容模型 → 完整 HTML5 文档（UTF-8、<meta charset>、中文原样）。
 * 说明：HTML 没有分页概念，`pagebreak` 写成带 class 的 div（本插件的读取端会忽略它）；
 * 图片块写成占位文本（该写入器不内嵌二进制）。
 */
export function documentToHtml(doc) {
  const meta = doc?.meta || {}
  const body = []
  for (const b of doc?.blocks || []) {
    switch (b.type) {
      case 'heading': body.push(`<h${Math.min(6, Math.max(1, b.level || 1))}>${runsToHtml([{ text: b.text }])}</h${Math.min(6, Math.max(1, b.level || 1))}>`); break
      case 'paragraph': body.push(`<p>${runsToHtml(b.runs)}</p>`); break
      case 'list': {
        const html = listToHtml(b.items, b.ordered)
        if (html) body.push(html)
        break
      }
      case 'table': body.push(tableToHtml(b.rows, b.header !== false)); break
      case 'quote': body.push(`<blockquote>${String(b.text ?? '').split('\n').map(l => `<p>${escapeHtml(l)}</p>`).join('')}</blockquote>`); break
      case 'code': body.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`); break
      case 'hr': body.push('<hr>'); break
      case 'pagebreak': body.push('<div class="page-break" style="page-break-after: always"></div>'); break
      // 插图链路（第十二轮需求 1a 的写出侧）：有路径就写成真 `<img>`（往返后仍是 image 块），
      // 没有路径（如 `![]( )`）保留占位文本 —— 不静默丢语义。
      case 'image': {
        const src = String(b.name ?? '')
        body.push(src
          ? `<figure><img src="${escapeHtml(src)}" alt="${escapeHtml(b.alt || '图片')}"></figure>`
          : `<p><em>[图片：${escapeHtml(b.alt || '')}]</em></p>`)
        break
      }
      default: if (b.text) body.push(`<p>${escapeHtml(b.text)}</p>`)
    }
  }
  const title = meta.title ? `<title>${escapeHtml(meta.title)}</title>\n` : ''
  const desc = meta.description ? `<meta name="description" content="${escapeHtml(meta.description)}">\n` : ''
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
${title}${desc}</head>
<body>
${body.join('\n')}
</body>
</html>
`
}
