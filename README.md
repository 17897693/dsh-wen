# dsh-wen — DSH 系统级 Office 全能插件

## 关于这个项目

考公考编闲暇之余写的 DSH 办公文档插件，**自用为主，顺手开源**。
项目原名 `dsh-office`，后更名为 **`dsh-wen`**。

### 维护状态：不承诺维护

- 不保证跟进 DSH 的版本变更；作者可能随时停止更新或归档本仓库
- 不承诺回复 issue / PR
- **欢迎 fork 自行维护、修改、再分发**（见 [LICENSE](LICENSE)）

### 平台

**Windows x64** 为完整能力平台。OCR（内嵌 RapidOCR-json.exe）与 PDF 产出的渲染级质量门
（PowerShell + WinRT）依赖 Windows；其他平台自动降级为纯 JS 能力
（docx / xlsx / pptx / pdf 文本层 / odf / csv / html / md 的读写与互转）。

### 安装

推荐作为 DSH 组合包安装（profile 内一条命令）：

```sh
dsh plugin --profile web add ./dsh-wen                # 本地目录
# 或从 git：dsh plugin --profile web add github:17897693/dsh-wen
```

包内的 [`cordis.patch.yml`](cordis.patch.yml) 会随 `dsh.bundle` 声明自动作为一个配置层生效。
若你不安装为包、而是手动挂载，请把该文件里两行的 `name` 改成 `index.js` / `skill.js`
的绝对路径（home 级 `${DSH_HOME}/cordis.patch.yml` 同样可用）。

```sh
dsh --profile web --dump-config      # 应能看到 "# == dsh-wen" 层
```

### 兼容性说明

- **工具名仍是 `office_read` / `office_create` / `office_edit` / `office_convert`**，
  环境变量仍是 `DSH_OFFICE_*`：这些是既有调用契约，更名时刻意未动。
- 仅**插件标识**更名（包名、skill 名、cordis 插件名）。
- 内部缓存目录名与批间标记仍沿用 `dsh-office`，以保证既有缓存与输出格式不变。

### 文档

- [`SKILL.md`](SKILL.md) —— 面向调用方的操作手册（工具选择 / 参数 / 已知边界）
- [`CHANGELOG.md`](CHANGELOG.md) —— 逐轮修复史与实测数据
- [`DEVELOPMENT.md`](DEVELOPMENT.md) —— 内部实现细节与诊断过程
  （上述两份文档写于更名前，故内文仍记作 `dsh-office`；其中样本文件名与环境路径已代号化）

### 第三方组件

`vendor/` 内的 OCR 引擎与模型属于第三方作品，各自遵循其许可 —— 见
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)。不需要 OCR 时可以不打包 `vendor/`。

### 免责

本软件按「原样」提供，不附带任何形式的担保。

零第三方依赖（仅 Node 内置模块）的办公文档读写/编辑/转换能力，注册为 DSH 全局工具，
**所有 profile 的所有新会话自动可用**，无需手动启用或配置路径。

- 安装位置：`${DSH_HOME}/plugins/dsh-wen/`（本机为 `${DSH_HOME}/plugins/dsh-wen\`）
- 挂载方式：`${DSH_HOME}/cordis.patch.yml`（DSH_HOME 级 patch 层，对所有 profile 生效）
- 插件行（**绝对路径，两行**）：`id: tool-office` 与 `id: skill-office`，`name:` 分别是
  `${DSH_HOME}/plugins/dsh-wen/index.js` 与 `${DSH_HOME}/plugins/dsh-wen/skill.js`
  （裸包名 `dsh-wen` 在 desktop profile 下解析失败，patch 文件里已注明）

> **本文档只装**：快速入门 / 格式能力矩阵 / 常见用法 / 指向下面两份文档的入口。
> 「工具怎么选、按什么顺序操作、参数怎么传、哪些边界仍然生效、失败后下一步做什么」
> 的完整口径在 [SKILL.md](SKILL.md)（**面向调用方的唯一权威操作手册**）；
> 逐轮修复史、内部实现细节、诊断过程与当轮实测数据在 [DEVELOPMENT.md](DEVELOPMENT.md)。

## 提供的工具

| 工具 | 作用 |
| --- | --- |
| `office_read` | 读取办公文档 → Markdown / 纯文本 / JSON / 元数据。支持 `sheet`（工作表）、`pages`（PDF 页码范围）、`pageFrom`/`pageTo`（PDF 页级续读）、`offset`/`limit`（超长续读）、`paths`（批量盘点：一次返回逐文件轻量 stats）、`boundary`（分批读取的批间衔接内联开关）。⚠ `as="json"` / `as="meta"` 的 `content` 是 **JSON 文本字符串**（不是对象，调用方需自行 `JSON.parse`） |
| `office_create` | 按扩展名创建文件。内容来源任选其一：`markdown`、`document`、`slides`、`workbook`、`table`、`from`（转换式创建）；`encoding` 控制文本类目标写出 BOM |
| `office_edit` | 原地修改已有文件（zip/XML 级，保留其余内容；CSV/文本编辑保留原 BOM） |
| `office_convert` | 读取 source，按统一内容模型写出 target（扩展名决定目标格式）；`encoding` 控制文本类目标写出 BOM |

## 格式支持

| 类别 | 读取 | 创建 | 编辑 |
| --- | --- | --- | --- |
| Word | `.docx` `.docm` `.dotx`、`.doc`（OLE2 97-2003）、`.wps`（OLE2 变体） | `.docx` | `replace_text` / `append_markdown` / `set_meta` / `insert_image` / `append_image`（插图：`{path 或 base64, alt?, width?（磅；省略则按原图像素 × 72/96，上限 A4 可用宽 451.3pt 等比缩）, after?（文本锚点）}`，PNG/JPEG/GIF/BMP，内容 SHA-256 去重；zip 级新增 `word/media/*` + image 关系 + `w:drawing`） |
| Excel | `.xlsx` `.xlsm` `.xltx`、`.xls`（BIFF8）、`.et`（OLE2 变体）、`.csv` `.tsv` | `.xlsx` `.csv` `.tsv` | `set_cell` / `append_rows` / `replace_value` / `add_sheet` / `rename_sheet` / `delete_sheet` |
| PowerPoint | `.pptx` `.pptm` `.potx`、`.ppt` `.dps` `.pps` `.pot`（PPT97 记录流） | `.pptx` | `replace_text` / `add_slide` / `update_slide` / `delete_slide` |
| PDF | `.pdf`（文本层提取，含 CID/ToUnicode 中文、对象流、PNG/TIFF 预测器；**含 AES-256（R5/R6）空口令加密**） | `.pdf`（拉丁用 base-14；中文/emoji **内嵌系统字体子集**：Identity-H + FontFile2 + 完整 ToUnicode，`simsun.ttc`→`simhei`→`msyh` 字体链，写盘前过**产出质量门**；markdown 里的图片 → `/XObject`：JPEG 走 `DCTDecode` 原样内嵌、PNG 走 `FlateDecode`+`/SMask`） | 不支持原地编辑（先 `office_convert`） |
| 网页 | `.html` `.htm`（零依赖 DOM 解析：h1-h6/嵌套列表/表格/引用/代码块/外链保留，style/script/隐藏子树剥离，实体解码；`<img>` → image 块） | `.html`（语义化 HTML5：h1-h6/表格/嵌套列表/引用/代码块/行内强调/`<figure><img>`，UTF-8 + `<meta charset>`，中文不转实体。⚠ **`.htm` 只能读/编辑，不能作为创建或转换目标**） | `replace_text` / `append_text` / `prepend_text`（**文件级文本替换**，不改结构、不重新排版） |
| OpenDocument | `.odt` `.ods` `.odp` | `.odt` `.ods` `.odp` | `replace_text` / `append_markdown`（.odt） |
| 文本类 | `.md` `.txt` `.json` `.jsonl` `.rtf` | `.md` `.txt` `.json` | `replace_text` / `append_text` / `prepend_text` |

转换矩阵：任意可读格式 → 任意可创建格式（经统一内容模型，如 docx→pdf、xlsx→csv、pptx→md、md→docx/pdf/pptx/html）。

## 使用示例

```
office_read   path="D:\报告.docx"                  # 读 Word 为 Markdown
office_read   path="D:\台账.xlsx" sheet="明细"      # 只读某个工作表
office_read   path="D:\讲义.pdf" pages="1-5"        # 只读 PDF 前 5 页
office_read   path="D:\讲义.pdf" pageFrom=6 pageTo=20   # 页级续读（含端点，等价 pages="6-20"）
office_read   path="D:\数据.xlsx" as="meta"         # 只看工作表/行列统计
office_read   paths=["<corpus>/真题"]            # 批量盘点：一次返回目录内逐文件轻量 stats
office_read   path="D:\笔记.html" as="markdown"     # 读网页笔记（style/script 剥离、表格/嵌套列表保留）
office_read   path="D:\台账.xlsx" recalc=true        # 重算公式：公式格给计算值 + stats.recalc（不支持的函数标 unsupported，绝不猜）

office_create path="D:\汇总.docx" markdown="# 标题\n- 要点\n\n| A | B |\n| --- | --- |\n| 1 | 2 |"
office_create path="D:\成绩.xlsx" workbook={sheets:[{name:"S1",rows:[["姓名","分数"],["张三",95]]}]}
office_create path="D:\汇报.pptx" slides={slides:[{layout:"title",title:"标题",subtitle:"副标题"},{title:"要点",bullets:[{text:"一",level:0}]}]}
office_create path="D:\报告.pdf"  markdown="# 中文标题\n\n正文（中文默认内嵌系统字体子集；WinRT/Edge/Chrome 渲染正常、提取不丢字符）"
office_create path="D:\交付.csv" table={rows:[["城市","备注"],["唐山","含,逗号"]]}   # 默认 UTF-8 with BOM，Excel 双击不乱码
office_create path="D:\申论.docx" markdown="# 作答" grid="20x25"                   # 申论稿纸：Word 文档网格（每行 20 字 × 每页 25 行；也接受数字 grid=20）
office_create path="D:\图文.pdf" markdown="# 标题\n\n![示意图](D:\图.png)"          # 含图 markdown → PDF 真内嵌图片（JPEG DCTDecode / PNG FlateDecode+SMask）

office_edit   path="D:\报告.docx" operations=[{op:"replace_text",find:"旧",replace:"新"},{op:"append_markdown",markdown:"## 补充"}]
office_edit   path="D:\台账.xlsx" operations=[{op:"set_cell",sheet:"明细",cell:"B3",value:100},{op:"add_sheet",name:"新增"}]
office_edit   path="D:\汇报.pptx" operations=[{op:"add_slide",slide:{title:"答疑",bullets:["Q&A"]}}]
office_edit   path="D:\报告.docx" operations=[{op:"append_image",path:"D:\图.png",alt:"示意图",width:400}]
office_edit   path="D:\报告.docx" operations=[{op:"insert_image",base64:"data:image/png;base64,iVBOR…",alt:"插图",after:"第二章"}]

office_convert source="D:\台账.xlsx" target="D:\台账.csv"
office_convert source="D:\报告.docx" target="D:\报告.pdf"
office_convert source="D:\笔记.html" target="D:\笔记.md"     # 网页 → Markdown
```

## 常见用法

### 扫描件 / 图片型 PDF

`office_read` 内置扫描页处理，**无需截图、无需分步操作**：`ocr="auto"`（默认）在页面没有文本层时
自动识别（未指定 `pages` 时先做前 3 页预览；指定 `pages` 时单次上限 20 页），
默认走**本地 RapidOCR-json**（离线），只有它判为可疑的页才升级视觉模型复核。

```text
office_read path="D:\扫描件.pdf"                      # 自动 OCR 首页若干张
office_read path="D:\扫描件.pdf" as="meta"            # 先看哪些页是扫描页（stats.scannedPages）
office_read path="D:\扫描件.pdf" pages="3-8"          # 只 OCR 第 3–8 页
office_read path="D:\扫描件.pdf" pages="1-20" ocrEngine="local"   # 批量识别且不花视觉额度
```

参数语义（`ocr` / `ocrEngine` / `ocrEngine="vision"` 的适用场景）、
降级链与 sidecar（`.ocr.md`）、来源标注、换倍率重试、失败后怎么续读、
以及全部 OCR 相关环境变量 → **[SKILL.md](SKILL.md)「读取的正确顺序」与「调优环境变量」**。
栅格化与引擎发现的实现细节、缓存位置的历史坑 → [DEVELOPMENT.md](DEVELOPMENT.md)。

### 加密 / 权限受限 PDF

只设了**权限密码**（禁止复制/打印）而**打开密码为空**的文件（考试资料、电子书很常见）会
**透明解密后正常读取**，无需额外操作，也**不改写原文件**。真正需要**打开密码**的文件**不会被猜测或绕过**，
会明确提示"需要密码，请在阅读器中输入密码后另存为未加密版本"。

覆盖范围：RC4-40 / RC4-128 / AES-128 / **AES-256（R5/R6）**。
算法细节、判别实验结论与"与第三方逐字节一致仍未取得"这条自证风险 → [SKILL.md](SKILL.md)「已知边界」。

### 批量盘点大目录

```text
office_read paths=["<corpus>/真题"]     # 一次调用返回逐文件轻量 stats：format/pages/characters/
                                          # textLayerUsable/scannedPages/garbledPages/blocks/sheets/slides
                                          # + suggestedBatches（建议分批数），绝不返回正文
```

20+ 个 PDF 的盘点从 20 次 `as="meta"` 调用变成 1 次；目录里的插件缓存（`.ocr.md` / `.read.md` /
`.ocr.json`）自动排除。单文件失败进该行 `error`，不连累整批。
`suggestedBatches` 的口径与单文件 `meta.suggestion` 一致：乱码 PDF 按 20 页/批；
**干净 PDF 只有超过 12 页才按 15 页/批**（≤12 页算 1 批）；其他格式按内联上限估算（不超上限就是 1）。

### 中文 PDF 产出

`office_create` / `office_convert` 写到 `.pdf` 时**默认内嵌系统字体子集**（Identity-H + `FontFile2` +
完整 `ToUnicode`，`/BaseFont` 带 `XXXXXX+` 前缀），WinRT / Edge / Chrome 渲染正常、`office_read` 提取不丢字符。
写盘前还会过**产出质量门**（字节级 + 第 1 页渲染级抽检）：判定为"渲染为空白"时**拒绝落盘**并抛四要素错误。
返回值带 `stats.pdfQuality = { embedded, renderCheck, firstPageBytes, images, imagesSkipped?, notice? }`。

⚠ 中文 PDF 在非 `%TEMP%` 目录里**也能正常产出**（R18 修好了渲染副本的落点）—— 若你自定义了
`DSH_OFFICE_PDF_GATE_DIR`，请把它指向 `%TEMP%` 内的目录：本机 WinRT 的 PDF 渲染器
**只读得到 `%TEMP%` 下的文件**（详见 [SKILL.md](SKILL.md)「调优环境变量」）。

## 已知边界（速查）

每条都有一句话结论；**完整口径、判别方法与失败后的下一步在 [SKILL.md](SKILL.md)「已知边界」**。

- **扫描件无文本层不会报错**（默认自动 OCR）；只有显式 `ocr="never"` 且整本无文本层才抛错。
- **过门 ≠ 可直接引用**：质量门只说明"质量门放过"，数字/正负号密集的表格页**务必回原 PDF 核对**
  （典型症状是**负号被丢**，视觉复核升级治不了这类版面级噪声）。
- **中文 PDF 乱码有两种成因**（自家解析器错位已修 / 文档真缺 `ToUnicode`），修法不同，
  靠 `stats.textLayerUsable` + `qualityGate.reasons` 的措辞区分。
- **`.wps` / `.et` / `.dps`** 的金山新版专有二进制解析不了，请先另存为 OOXML；OLE2 变体可直接读。
- **PDF 不支持原地改文本**；要改就 `office_convert` 重出一份。
- **`pptx` 的 `update_slide` 会重写该页**（该页原有图片/复杂图形会丢），其他页不受影响。
- **图片写出端**：`docx` / `pdf` / `html` 真内嵌；`.odt` / `.pptx` / `.md` 等**显式降级**为字面
  `![alt](path)` 文本 + `stats.imageFallback` + `notice`（**绝不静默丢图**）。
- **`.html` 编辑是文件级文本替换**，不改结构、不重排版。
- **写盘一律原子**（目标同目录唯一临时件 → 成功后替换）；失败时**原文件逐字节不变** + 「写盘失败｜四要素」。
- **ZIP 解压有护栏**（单条目 256 MiB / 累计 512 MiB / CRC32 默认校验），越界与说谎的长度给可读中文错误。
- **ZIP 写出端支持真实 ZIP64**（≥65535 条目、尺寸或偏移 ≥4 GiB 时按规范写扩展字段 + EOCD64）；
  "**≥4 GiB 数据尺寸**"这条分支**只有布局级 + 声明级证据**（端到端峰值内存 ≈8.0 GB，本机跑不了）——
  但声明格式已被 `.NET ZipArchive` 正确接受（`Entry.Length` 读成 4 GiB+）。详见
  [DEVELOPMENT.md](DEVELOPMENT.md) R19 §A。
- **工具参数错误带 `operations[N]` 下标定位**（缺 `op` / `op` 不认识 / 字段类型不对都在参数层报错）。
- 未实现能力的清单与重开条件（PDF 合并拆分、OMML 公式写入、`.html` 结构化编辑）见
  [DEVELOPMENT.md](DEVELOPMENT.md)。

## 开发与修复史

逐轮修复注记、内部实现细节与诊断过程全部在 [DEVELOPMENT.md](DEVELOPMENT.md)：
那里有"历史上修过什么、为什么这么修"的全量资料与逐轮实测数据（含 R16 / R18 的交付报告）。

## 验证

`node test.mjs` 运行端到端检查（**具体数量看最后一行的 `N checks`，不要在文档里写死**）：
创建/读回/编辑/转换全格式闭环、真实 docx/xlsx/pptx/pdf/csv 读取、独立 zip 与 CRC 校验、
**包结构体检**（`checkOoxmlPackage()`：`[Content_Types].xml` 覆盖、关系可解析、全部 XML 可解析）、错误处理。

在任意目录都能跑：产物目录默认 `./test-out`，不可写时自动回退到系统临时目录；
`DSH_OFFICE_TEST_OUT` 可指定产物目录，`DSH_OFFICE_TEST_SCAN_PDF` 可钉住 OCR 样本
（默认取 `~/.dsh/attachments` 中最近一个 >5 MiB 的 PDF，并拷进产物目录使用，不会改动用户附件与其 sidecar）。

**慢 / 可选的三条用例**（默认跳过，发版前手工跑一次）：`DSH_OFFICE_TEST_EXPAND_ARCHIVE=1`
（真·`Expand-Archive` 解 65536 条目，约 20 分钟）、`DSH_OFFICE_TEST_REAL_4GIB=1`
（真实 >4 GiB 合法 ZIP64 条目，约 15 秒）、`DSH_OFFICE_TEST_NET_DIR=<网络盘目录>`
（网络盘上的原子替换）。全部测试变量见 [SKILL.md](SKILL.md)「排查」。

**原生 Office 冒烟**：`node test.mjs` 会调本机的 Word / PowerPoint / Excel COM 做真实开箱验证 ——
产出的 `.docx`（含**含图**、稿纸网格）、`.html` 能被 Word 16.0 打开并导出 PDF；
产出的 `.pptx` 能被 PowerPoint 16.0 打开（并带"把 notesMaster 主题改回共用 theme1 **必须仍 FAIL**"的负向控制）。
这些冒烟在本机 Office 不可用时会**明确 skip 并打印原因**（不静默当通过）；
真实 AES-256 样本的回归槽位见 `DSH_OFFICE_TEST_AES256_PDF`。

## 卸载

删除 `${DSH_HOME}/cordis.patch.yml` 中的 **`tool-office` 与 `skill-office` 两行**（或整个文件），
再删除 `${DSH_HOME}/plugins/dsh-wen/` 目录即可；也可只把这两行 `disabled: true` 临时停用。