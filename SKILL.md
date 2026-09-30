# dsh-wen — 办公文档读写与转换

四个全局工具由 `${DSH_HOME}/plugins/dsh-wen/` 注册，对**所有 profile 的所有会话**可用。
遇到 Word / Excel / PPT / PDF / ODF / CSV / RTF / **HTML** 的读取、创建、编辑、互转，**先用它们**，
不要绕道 Python 库、LibreOffice、`pdftotext`、`npx anydoc` 之类外部转换器（那些在本机多半没装，
且受沙箱限制装不上）。

## 工具选择

| 要做的事 | 用 |
| --- | --- |
| 看内容 / 抽取文本 / 列结构 | `office_read` |
| 新建文件（md→docx/pdf/pptx、对象→xlsx 等） | `office_create` |
| 改已有文件（保留其余内容与样式） | `office_edit` |
| 格式互转（xlsx→csv、docx→pdf、pptx→md …） | `office_convert` |

扩展名决定目标格式；能力矩阵：任意可读格式 → 任意可创建格式（经统一内容模型）。

## 读取的正确顺序

1. **先 `as="meta"`**（几 KB，不要一上来灌正文）。返回 `pages` / `pagesWithText` /
   `scannedPages` / `blocks` / `characters`，据此判断该读哪几页、要不要 OCR。

   **`as="meta"` 还给出文字层质量画像**（纯 CPU：不渲染、不触发 OCR，meta 仍是几 KB、仍是"第一步"）：

   | 字段 | 含义 |
   | --- | --- |
   | `stats.textLayerUsable` | 文字层**能不能用**。注意与"有没有"（`pagesWithText` / `scannedPages`）是两件事——CID 乱码书 `pagesWithText` 很高但整本不可读 |
   | `stats.garbledPages` | 命中乱码的页，`"3,5-33"` 形态；干净时为 `[]` |
   | `stats.qualityGate` | `{ testedPages, garbledPages, reasons }`，`reasons` 只给前 3 条摘要 |
   | `stats.suggestion` | 命中乱码时为 `{ note, copy }`，`copy` 是**可直接复制**的整本 OCR 参数串（含分批页码）；干净且大文档时是分批计划 |

   非 PDF 的 meta 分支同样带 `stats.textLayerUsable`（乱码时附 `qualityGate.reasons`）。
   既有字段**不改名、不改值**。

2. 再按 `pages` / `sheet` / `offset` + `limit` **分段取**，不要整本读进上下文。
   **续读协议（已收敛）**：`content` 只放**纯前缀**，一个字的截断说明都不掺；说明在 `notice`（字符串）
   与 `stats.truncateNote` 里。于是硬不变式恒成立：**`offset + content.length === nextOffset`**。
   续读直接用 `nextOffset`，分段拼接**直接** `content += r.content` 即可
   （旧版把"已达上限"后缀掺进 content，逼得消费方手工 `slice(0, nextOffset - offset)`，写错就"又重又漏"）。
   `truncated` 语义不变。

   **PDF 页级续读（第四轮新增）**：`pageFrom` / `pageTo`（含端点）= 从第 N 页读到第 M 页，
   与 `pages="N-M"` 等价但更符合"按页分批"的直觉；两者**互斥**（同给报错）。
   返回的 `stats.pageFrom` / `stats.pageTo` 是实际覆盖页区间，`stats.nextPage` 是下一未读页
   （读完最后一页时不出现）——跨批续读按 `nextPage` 推进即可，不用自己算页码。

   **批间衔接（第四轮新增，跨批读取不再丢上下文）**：凡 `pages` 指定的分批读取，
   默认在 `stats.prevTail`（上一批末页的末行，约 100 字符）与 `stats.nextHead`（下一批首行的开头，
   约 100 字符）给出批间上下文，`notice` 里有一行摘要；`boundary: true` 时再以
   `> 〔dsh-office 批间回看｜…〕` / `> 〔dsh-office 批间预览｜…〕` 引用行**内联**进正文首尾
   （拼接时剔除这两类行即可，默认不内联——`content` 保持纯正文，分批拼接与整本读取逐字 diff 依然成立）。

   **批量盘点（第四轮新增）**：`paths` 传**数组**（条目可以是文件或**目录**，目录会展开为其中
   受支持的办公文件，`.ocr.md` 等插件缓存自动排除）→ 一次调用返回逐文件轻量 stats 清单：
   行字段与单文件 meta.stats 同名（`format` / `pages` / `characters` / `textLayerUsable` /
   `scannedPages` / `garbledPages`，非 PDF 另加 `blocks` / `sheets` / `slides`），另加
   `suggestedBatches`（建议分批数，与单文件 `meta.suggestion` **同口径**：乱码 PDF 按 20 页/批；
   **干净 PDF 只有超过 12 页才按 15 页/批**，≤12 页算 1 批；非 PDF 按内联上限估算）。
   **只聚合元数据、绝不返回正文**；单文件失败进该行 `error` 字段而不连累整批
   （盘点 20+ 个 PDF 从 20 次 meta 调用变成 1 次）。`path` 与 `paths` 二选一；
   `as="markdown"` 之类与批量形态冲突会显式报错。
   （"23 个真题 PDF ≈1.1 秒"是当轮单机实测、**未自动化**，见 [DEVELOPMENT.md](DEVELOPMENT.md)。）

3. **中文 PDF 正文乱码时，先分清是两种成因里的哪一种 —— 修法完全不同**：

   | 成因 | 判别 | 修法 |
   | --- | --- | --- |
   | **自家解析器错位**（`pdf.js` 的 ObjStm off-by-`First`，第三轮已修） | `stats.textLayerUsable === false` **且**页 `qualityGate.reasons` 里是"控制字符占 X%" 这类字节级噪声；正文里混着 ASCII 与 2 字节 CID 被拆开的单字节残渣 | **修解析器**（已修）。第三轮之后这类文档直接就是可读的，**不要**再 OCR |
   | **文档真的缺 `ToUnicode`** | `textLayerUsable === false` **且** `qualityGate.reasons` 明说「私用区码点占 X%（CID 字体缺 ToUnicode 的典型产物）」 | 本地没有映射表，只能 OCR |

   两者的共同入口都是 `as="meta"` 的 `stats.textLayerUsable === false`（旧版只能靠"看着像乱码"）；
   `qualityGate.reasons` 里的措辞是区分依据。
   确认是"真缺 ToUnicode"之后：照 `stats.suggestion.copy` 直接执行，或
   `office_read path=… ocr="always" pages="1-5"` 强制重识别（结果缓存为同名 `.ocr.md`，
   续读命中缓存、不重复消耗识别额度）。
   `ocr="auto"`（默认）**也会**把质量门判为乱码的页送去识别，不再只认"完全没文本层"。

   ⚠ **目录页不算乱码**（第三轮修）：整页由点前导（`第一章总论..................1`）与页码组成时，
   逐页 CJK 覆盖率天然极低（实测某 35 页样本的目录页 **12.2%** vs 全书 **73.8%**），会撞上
   "相对 CJK 覆盖率"判据 —— 那是**假阳性**。第三轮起**连续 ≥3 个 ASCII 句点视为排版装饰、
   不计入可见字符**，该页覆盖率回到 **79.7%**、整本 `garbledPages=[]`（目录页不再白跑一次 OCR，
   `convert` 也不再因此拒绝整本）。判据本身没有放宽：真缺 `ToUnicode` 的文档是
   **PUA / 替换字符成片**，"私用区 / 替换字符 / 控制符"三条仍各自独立命中。

   **第四轮新增——字符级/版面级启发（治"字符可提取但语义破碎"）**：旧判据只认
   "绝对不可信"的码位（替换字符/私用区/控制符）与相对 CJK 覆盖率，字符**能**提取但
   语义**碎**的页会漏检（实测某页读出"代表2的发言…"才发现）。现在 PDF 逐页质量门
   额外启用 5 条启发（全部带行数/占比门槛）。**门槛口径要分清**（第十七轮 R16 核对）：
   其中**版面级 2 条**（重复行率、单字/双字行率）只在**逐页质量门**（`structural` 模式）生效，
   整本口径不启用——合法的重复表格行不会把整本文档推进 sidecar 兜底；
   **字符级 3 条 + 相对判据 1 条**（Latin-1 高带、制表/几何符号、连续非词典字符、非词典区相对异常）
   挂在 `visible >= 40` 上，**页面级与整本口径都会跑**：

   | 启发 | 判定 | 典型病灶 |
   | --- | --- | --- |
   | 重复行率 | ≥10 行"像正文"的行里 ≥50% 逐字重复 | 解析破碎、同一段被反复贴 |
   | 单字/双字行率 | ≥12 非空行里 ≥75% 只有 1-2 字（纯数字行不算） | 逐字断行、每字一行的错位提取 |
   | Latin-1 高带占比 | 0xC0–0xFF 字符（×÷ 除外）占可见字符 ≥50% | UTF-8 被按 Latin-1 解的 mojibake（"ä¸­æ–‡"） |
   | 连续非词典字符 | 无元音字母串（≥8 连排不含 aeiou）占字母 ≥25% | 字符错映射出的"伪词" |
   | 制表/几何符号 | `U+2500–U+25FF` 占可见字符 ≥30% | CID 错映射（框线/几何图形替掉了正文） |
   | 非词典区相对异常 | 希腊/西里尔/制表符区占比 ≥12% 且全书基线 <3% | 干净书里混入错映射页（俄文**原文档**不误伤） |

   触发的页照旧计入 `stats.garbledPages` / `qualityGate.reasons`，`ocr="auto"` 会自动把它们
   送去重识别，notes 里给出可复制的 `ocr="always"` 参数串——**该页 OCR 一遍通常就修好**。

4. 扫描页无需特殊处理：`ocr="auto"`（默认）在无文本层时自动识别，但**未指定 `pages` 时最多 3 页**；
   指定 `pages` 时单次上限 20 页。要整本识别就分批传页码范围（本地引擎按 20 张一批**串行**跑完，
   第 21 页起不会被静默推给视觉：今天工具层已封顶 20 页，这条护栏是为将来放宽准备）。
   **两种"少做"都不再静默**（第八轮）：① 未指定 `pages` 只做前 3 页预览 →
   `stats.ocrPreview` 给 `{preview,total,skipped}`；② 要求的页数被 20 页上限砍掉 →
   `stats.ocrPagesCapped` 给 `{limit,requested,applied,skipped}`；
   两种都会在正文脚注附**可复制的续读命令**
   （`office_read path="…" ocr="always" ocrEngine="local" pages="<未做的页>"`）。
   **响应尾部会报覆盖情况**（如「第 4-115 页未识别（无文本层共 115 页，已覆盖 1-3 / 全书 115 页）」）；
   `stats.ocrCovered`（`1-3 / 115`）与 `stats.ocrUncovered`（`4-115`）是同一信息的机器可读形式；
   sidecar 头部还写着 `<!-- covered: 1-3 | total: 115 | parser: 2 | src: rapidocr=1-3 -->`——grep 一下就知道识别进度，
   不用整本读进来。分批识别时该行会**合并更新**，不会回退——`covered`、`src`（每页由哪个引擎产出）、
    `retry`（哪几页是换倍率才救回来的）**三段都跨批累积**：先跑 1-20 再跑 21-35，头部会得到
    `src: rapidocr=1-20;vision=21-35` 这样的并集，而不是只剩最后一批。
   紧随其后还有**身份行**（第十七轮 R16 新增）：`<!-- srcpath: <64hex> | srcsha256: <64hex> | srcsize: N -->`
   —— `srcpath` 是源文件**路径指纹**、`srcsha256` 是源**内容 SHA-256**；`parser:` 仍是**文字层解析器版本戳**。
   `parser:` 或身份**任一不符**，该 sidecar 就**整份作废**，原因写进 `stats.ocrCacheStale.reason`
   （`parser` / `parser-missing` / `identity-missing` / `path-mismatch` / `content-mismatch` / `unverifiable`），
   `stats.ocrCacheNote` 给人话解释。所以**同一路径的 PDF 被替换、或同名 PDF 换到另一个目录**，旧缓存都不会被误用；
   旧格式（没有身份行）的缓存同样被明确作废 —— 重新识别即可，旧文件可直接删。
   ⚠ 集中缓存目录（`DSH_OFFICE_CACHE_DIR`）与临时回退目录下的文件名多了 8 位路径指纹后缀
   （`<name>-<key8>.ocr.md`）：这是本轮唯一的兼容性变化，**同目录 sidecar 的命名未变**。

5. **识别引擎默认是本地 RapidOCR-json（离线、约 0.5 s/页、不消耗模型调用）**，只有本地引擎判为可疑的页
   才升级视觉模型复核。所以：批量扫描资料直接 `office_read … ocr="always" pages="a-b"` 即可；
   想**完全不花视觉额度**就加 `ocrEngine="local"`（可疑页会写明原因而不是猜）；
   怀疑本地引擎有偏（旋转页、花体字、繁简混排）时用 `ocrEngine="vision"` 强制走视觉。
   正文里的来源标注会写清"这一页是谁识别的、有没有切片"：
   `本地 OCR · RapidOCR（N 框，置信 x）` / `视觉模型识别 · 本地复核`（本地没过质量门后自动升级）/
   `视觉模型识别 · 用户指定 ocrEngine:"vision"`（你要求的，不是自动升级）/
   `视觉模型识别 · 本地引擎不可用` / `OCR 缓存`（auto 命中缓存）/
   `OCR 缓存 · local|vision（指定的 X 未执行）`（指定引擎但命中别家缓存）。
   一页被切 2/4/8 片时标注里还有 `· 切 4 片（7 次调用）`——**片数≠页数**，记账口径写明。
   `stats.ocrEngine`（可用的本地引擎）/ `stats.ocrRequested`（本次指定的引擎）/ `stats.ocrEscalated` /
   `stats.ocrVisionCalls` 给出同样的回显。
   **命中缓存不等于重跑**：若指定的 `ocrEngine` 与该页缓存来源不一致，插件**不会**自动重跑，
   而是回显「OCR 缓存 · local（指定的 vision 未执行）」并给出 sidecar 路径——想强制重识别就删那个 `.ocr.md`。
   旧格式 sidecar 没记录来源，会标"来源未记录"，同样给出删除路径。
   视觉复核回来的客套话在**写入 sidecar 之前**逐片剪掉：`Transcription:` 之类的标记行、紧跟其后的图说句、
   任意位置的 `Uncertain:`、``` 围栏；多片拼接时第二片起同样处理。命中缓存时也会对旧 sidecar 做同样的清洗并
   **自动回写**。清洗只认结构（标记行 + 图说句式），不认"所有可能的客套话"——没把握时保留原文，宁可留噪不删正文。

6. **质量门失败页会自动换渲染倍率重试**：同一页在原生分辨率下被判"置信度低 / 版面复杂"，换个倍率常能过
   （当轮实测：某些页原生不过、备选倍率能过；**这种好坏不单调**，所以重试表给的是多个候选而不是无脑放大）。本地批次里没过门的页按
   `DSH_OFFICE_OCR_RETRY_SCALES`（默认 `2,1.5,1`）**逐倍率串行**重试，取"过门且置信度最高"的一次；
   全失败才记 `failed`，原因里带"已试过 scale=…"。
   **短句占比高属"结构性失败"（表格/数字页天然碎），换倍率无效、直接短路不重试**
   （失败原因记成"（结构性失败，换倍率无效）"；该页**仍会照常交视觉桥**，只是不再烧渲染+引擎时间）；
   且这条判据**只在低置信时判硬** —— 置信 ≥ `DSH_OFFICE_OCR_SHORT_SCORE`（默认 `0.95`）即**放行**，
   不再误杀"认得很准"的表格页（旧实现无条件判硬，会把置信 0.989、文本完全可读的页丢成空字符串）。
   判据本体：在 ≥ `DSH_OFFICE_OCR_SHORT_MIN_BOXES`（默认 `30`）个识别框里，≤6 字片段占比
   **超过** `DSH_OFFICE_OCR_SHORT_RATIO`（默认 `0.6`）。三个阈值都可用环境变量调（见下方总表）。
   记账**不静默**：`stats.ocrRetried`（页区间）/ `stats.ocrRetryScale`（如 `{"17":2}`）/ 正文脚注 /
   sidecar manifest 的 `retry: 17=2`（`src` 仍记 `rapidocr`——引擎没变，只是换了倍率）。
   批内失败页同时进 `stats.ocrFailedPages`（页码 + 一句话原因），**不用去 grep sidecar 也能看到缺页**。

## 读不出正文时会怎样（自动降级链，无需手动绕行）

四级：**文本层提取 → 质量门 → 自动本地 OCR → sidecar 兜底**。

- 返回值带 `stats.fallback`：`none`（文本层干净）/ `ocr`（走了识别）/ `sidecar`（已转存文件）。
- 质量门命中且没能重识别时：正文整篇转存为**同名 `<文件>.read.md`**，返回值只给
  `path + stats + 首部摘录 + notice`，**且该返回值本身保证可无损序列化**。
  notice 里带可直接复制的下一步参数串；直接读那个 `.read.md` 即可（真实换行，不受内联上限影响）。
- `stats.quality` 记下判定依据（各比例 + reasons）。**绝不静默**——没命中时该字段不出现，干净文件逐字不变。
- `.read.md` 与 `.ocr.md` / `.ocr.json` 一样**受保护**：渲染缓存清理永不碰。
- 错误分支一律带四要素：**页码 / 格式 / 根因 / 可直接复制的下一步参数**；
  若这次失败**已经产生了 sidecar**（OCR 成果已落盘），错误文本还会给出 **sidecar 路径 + 当前
  `covered` / 缺页区间**，可直接照抄去读或去 grep（见「排查」最后一条）。

判据只看"绝对不可信"的信号（替换字符 / 私用区码点 / 控制符），**不拿语言当闸门**——
英文文档的 CJK 覆盖率天然为 0，用它当闸门会把正常英文 PDF 判成乱码；CJK 覆盖率只作**同文档内的相对**判据。

## `office_convert` 的出站质量门：宁可不产出，也不写垃圾

`convert` 在**写盘之前**跑一次质量门：

- **PDF**：逐页判文字层质量（本就无文字层的页不计入）。乱码页有完整 `.ocr.md` 覆盖 → 用 OCR 文本
  替换这些页后继续转，返回带 `stats.fallback="ocr"` / `stats.garbledPages` / `notice`（写明第几页被替代）。
- 乱码页**没有**可用缓存 → **拒绝写盘**，抛四要素错误（页码 / 格式 / 根因 / 下一步），
  **绝不产出目标文件**。旧版会把含 NUL 与 C0 控制符的二进制垃圾静默落盘且零告警
  （`read` 工具直接判 binary 拒读）—— 当轮的字节级实测（56,341 字节 / NUL 5135 个 / C0 13,666 个）
  属**当轮快照、未自动化**，见 [DEVELOPMENT.md](DEVELOPMENT.md)。
- **其他 kind**：对渲染出的正文本再跑一次模型级质量门（防 docx / xlsx 里塞乱码）。
- **干净源文件：返回值与产出文件逐字节不变。**

## 调优环境变量（全部 opt-in，不设时行为与旧版逐字一致）

| 变量 | 默认 | 作用 |
| --- | --- | --- |
| `DSH_OFFICE_VISION_CONCURRENCY` | `1` | 视觉复核**页与页之间**的并发度，只有 `2` / `3` 生效，其余值（含 `"4"`、`0`、垃圾值）一律回落 1。页内的 2/4/8 分带重试永远串行，本地引擎永远单进程单批 —— 不吃满多核这条红线不受影响。 |
| `DSH_OFFICE_VISION_MAX_CALLS` | 不设 = 无上限 | 单次 `office_read` 的视觉调用总预算。达限**不静默截断**：跳过的页连同页码写进 `stats.ocrVisionSkipped`（如 `3-6`）与脚注「第 3-6 页因视觉调用预算上限跳过（计划 6 / 完成 2 / 跳过 4）」，每页仍带失败原因。`0` = 一页都不许走视觉。 |
| `DSH_OFFICE_RENDER_SCALE` | 不设 = 系统原生分辨率 | 按 **96dpi × scale**（0.5–4，超出/非数字视为不设）栅格化，并**联动**引擎 `--maxSideLen`（`1123×scale` 向上对齐 256、最小 1024）。A4 像素（**已有自动化守门**，`P1-3b 各档 PNG 尺寸`，±2%）：`1` → 794×1123｜`1.5` → 1191×1685｜`2` → 1588×2246，跨机一致；不设时跟随系统 DPI（当轮本机 120dpi，未自动化）。 |
| `DSH_OFFICE_OCR_RETRY_SCALES` | `2,1.5,1` | **质量门失败页的备选渲染倍率表**（逗号分隔）。与当前生效倍率相同的候选自动跳过；空串 = 关闭重试（回到"失败即缺页"）。重试**串行**、渲染次数计入现有渲染记账；结果进 `stats.ocrRetried` / `stats.ocrRetryScale`、正文脚注与 sidecar manifest 的 `retry:`。**"短句占比高"（结构性失败）的页不进这条表**——见下三行。 |
| `DSH_OFFICE_OCR_SHORT_RATIO` | `0.6` | **短句占比阈值**：≤6 字片段占全部识别框的比例**超过**它才可能判硬（表格/数字页天然碎，一般别调）。`0` / 非法值回落 `0.6`。 |
| `DSH_OFFICE_OCR_SHORT_MIN_BOXES` | `30` | **短句判据的最小框数**：框数不足 30 时不套短句判据，避免小样本误判。`0` / 非法值回落 `30`。 |
| `DSH_OFFICE_OCR_SHORT_SCORE` | `0.95` | **短句判据联动的置信度下限**：平均置信度 `avg` **达到**它即"认得很准"，短句再多也**放行**（不判硬、不交视觉桥、不换倍率）。调高 = 更宽容。`0` / 非法值回落 `0.95`。 |
| `DSH_OFFICE_OCR_MIN_SCORE` | `0.88` | **质量门置信度下限**：`avg` 低于它直接判"置信度低"（可换倍率重试、可交视觉桥）。`0` / 非法值回落 `0.88`；调高更严（更多页升级视觉）。 |
| `DSH_OFFICE_RAPIDOCR_DIR` | 未设 | **指定本地引擎目录**（候选链**最高优先**；其余候选依次为 `vendor/`、`vendor/RapidOCR-json_v0.2.0`、`~/.dsh/ocr`、`~/.dsh/ocr/RapidOCR-json_v0.2.0`、`PATH` 下同名目录）。目录里必须有 `RapidOCR-json.exe` + `models/` 的四个默认模型，否则跳过该候选。 |
| `DSH_OFFICE_OCR_DISABLED` | 未设 | **设成任何非空值**（惯例 `=1`）即**关掉本地 RapidOCR**、退回纯视觉桥路径（来源标注会写"视觉模型识别 · 本地引擎不可用"）。只影响引擎查找，不删任何缓存。 |
| `DSH_OFFICE_OCR_MODELS` / `DSH_OFFICE_OCR_DET` / `DSH_OFFICE_OCR_CLS` / `DSH_OFFICE_OCR_REC` / `DSH_OFFICE_OCR_KEYS` | `models` / `ch_PP-OCRv4_det_infer.onnx` / `ch_ppocr_mobile_v2.0_cls_infer.onnx` / `rec_ch_PP-OCRv4_infer.onnx` / `dict_chinese.txt` | **本地引擎的模型目录 + 四个模型文件名**（`rapidocr.js` 的 `engineArgs()`）。`--models=` 取的是**相对引擎目录的目录名**（`process.env.DSH_OFFICE_OCR_MODELS \|\| 'models'`），另四项是**文件名**，逐项以 `--det=` / `--cls=` / `--rec=` / `--keys=` 传给 `RapidOCR-json.exe`（`process.env[k] \|\| DEFAULT_MODELS[k]`，`DEFAULT_MODELS` 见 `rapidocr.js`）。空串 / 未设回落上面五个默认值 —— 与引擎自带的 `models/` + 四个默认模型完全一致 ⇒ **不设 = 与旧版逐字一致**。换模型得连 ONNX 文件本身一起换；四项里少任何一个，`findEngine()` 会跳过该候选目录、**继续往下一个候选找**（那里的 `continue` 只跳过这一个目录；只有全部候选都不合格才退回纯视觉路径）。`_REC` 另决定 `engineLabel()` 里显示的模型名。⚠ **探测与运行不同源**：`findEngine()` 的可用性检查**写死**看 `<引擎目录>/models/` 里有没有 `DEFAULT_MODELS` 那四个文件，**不看** `DSH_OFFICE_OCR_MODELS` / `_DET` / `_CLS` / `_REC` / `_KEYS` —— 所以改了这五个变量时，引擎目录里**照样得留着 `models/` + 四个默认模型**；否则**该候选目录被跳过**，`findEngine()` 会**继续往下一个候选找** —— 本机因此常落到 `vendor/` 引擎，而运行侧仍传 `--models=<自定义>` 指向 **vendor 下不存在的目录**，引擎空跑、逐倍率重试后才升级视觉桥（禁用视觉桥时表现为整单报错）。见第九轮实测记录（`work\office-enhance2\stage\w4-item4-evidence.md`）（第九轮更正：原「安静回退视觉桥」在本机不成立）。 |
| `DSH_OFFICE_OCR_NO_ANGLE` | 未设 = 引擎默认（做方向分类） | **任何非空值**（truthy 判断，连 `"0"` 也算「设了」）即追加 `--doAngle=0 --mostAngle=0`，关掉方向分类 / 角度纠正（`rapidocr.js` 的 `engineArgs()`）。只对整页倒置、旋转的扫描件有意义；正常竖排文档开着更稳。 |
| `DSH_OFFICE_BLANKS` | 未设 | `1` / `true` / `on` / `yes`（不区分大小写）时，把 PDF 正文里的填空下划线 `_{3,}` 渲成 `<span class="blank">＿＿＿＿</span>`。**只改呈现**，不 gate 任何正确性修复；仅 `office_read` 的 PDF 分支 + `as="markdown"`/`as="json"` 生效，`as="meta"` 与 `as="text"` 不受影响，对 OCR 路线天然无效（识别产物里本就没有 `_`）。 |
| `DSH_OFFICE_PDF_EMBED_CJK` | 未设 = **开启** | 中文 PDF 写出端的**字体子集嵌入开关**。默认（未设）走 Identity-H + `FontFile2` + `ToUnicode` 的系统字体子集链；显式 `=0` 回退旧 STSong-Light 路径（**不推荐**：渲染空白/问号 + 提取丢字符的双风险，且会被产出质量门的字节级门槛拦下）。 |
| `DSH_OFFICE_PDF_SKIP_RENDER_CHECK` | 未设 | `=1` 跳过 PDF 产出质量门的**渲染级**抽检（"首页整页黑像素 < 40"判空白）；**字节级门槛照样拦**（`/FontFile2` 计数为 0 仍拒绝落盘）。非 Windows 平台自动等价于设置。 |
| `DSH_OFFICE_PDF_GATE_DIR` | 未设 = `%TEMP%\dsh-wen-pdfgate` | **渲染校验暂存目录**（第十八轮 R18 新增；`pdfGateStagingRoot()`）。渲染校验用的 PDF 副本一律落在这里、渲染完立即删除，校验通过后才原子发布到目标目录 —— 所以**目标目录在非 `%TEMP%` 位置也能正常产出 PDF**。⚠ 本机 WinRT 的 PDF 渲染器**只读得到 `%TEMP%` 下的文件**：把它指到 `%TEMP%` 之外会让渲染级抽检失败（报"第 1 页渲染失败"并附带位置提示），**要么不设、要么指到 `%TEMP%` 内的目录**。`DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` 或非 Windows 时本项无效。 |
| `DSH_OFFICE_CACHE_DIR` | 未设 | **sidecar 落盘目录**（`<name>.ocr.md` / `<file>.read.md`，文件名规则不变）。设了就把成果统一落到该目录——只读场景（附件目录、网络盘）也能缓存，且**不再污染用户目录**。未设 = 源文件同目录优先、不可写回退 `%TEMP%`（与旧版逐字一致）。该目录不可写则回退默认位置，并把**回退原因**写进 `stats.cacheDirNote` / `notice`（不静默）。 |
| `DSH_OFFICE_MAX_INLINE_CHARS` | `120000` | 内联体积护栏阈值（`0` = 关闭；**非 0 但 < 1000 的值会被忽略并回落 `120000`**）。超阈值时插件自己给"纯前缀 content + notice/stats.truncateNote"，不再靠 host 截断。显式传了 `limit` 就不拦。 |
| `DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` | `268435456`（256 MiB） | **zip 单条目解压后上限**。按中央目录的**声明尺寸**在解压前就挡（不"先解压再检查"），解压时再用 zlib `maxOutputLength` 早停。超限给可读中文错误 + 下一步；`0` = 关闭该上限。 |
| `DSH_OFFICE_ZIP_MAX_TOTAL_BYTES` | `536870912`（512 MiB） | **整归档累计解压上限**（同一次读取里所有已解压条目的和）。`0` = 关闭。 |
| `DSH_OFFICE_ZIP_CRC` | 未设 = 校验 | `=0` 关闭 CRC32 校验。默认只在中央目录 `crc !== 0` 时比对（旧式流式写包器不填 CRC 的包照常放行）；不符时**严格拒绝**、点名条目并给出这个逃生口（理由见 [DEVELOPMENT.md](DEVELOPMENT.md) 第十八轮）。 |
| `DSH_OFFICE_ATOMIC_FSYNC` | 未设 = 写盘时 fsync | `=0` 关闭原子写的 fsync（**网络盘 / OneDrive / 批量导入**这类 fsync 很慢的场景）。关掉只影响"断电后是否留下 0 字节/半截文件"这一条，不影响"临时件 → 成功后替换"的语义。 |
| `DSH_OFFICE_RECALC_MAX_CELLS` | 不设 = 无上限 | 公式重算（`office_read recalc=true`）的**单元格总数护栏**。超过就把整本原样返回 + `stats.recalc.skipped` 说明（**不静默截断、不半算、不动任何公式**）；`0` / 非法值 = 无上限（与旧行为逐字一致）。 |
| `DSH_OFFICE_TEST_AES256_PDF` | 不设 = 跳过 | **真实世界 AES-256（R5/R6）加密 PDF 的回归槽位**。设上它，`node test.mjs` 会额外用 `office_read` 打开该文件并断言能读出正文 —— 这是"夹具自证风险"的关闭开关（本机没有第三方 AES-256 实现可做交叉验证）。 |

**该不该调大**：当轮实测里**放大只增加耗时、不减少升级视觉的页数**（2× 时接近 2 倍时间）→ 默认保持不设。
"耗时 / 平均置信度 / 升级页数"那张逐档表属**当轮单机实测、未自动化**，见 [DEVELOPMENT.md](DEVELOPMENT.md)。
已有自动化守门的只有"A4 像素 = 96dpi × scale"。另注：引擎对**同一张 PNG 是确定性的**
（同一张图连跑 3 次结果逐字相同），但不同渲染分辨率之间的好坏**不单调** —— 自造小字夹具上 1.5 反而明显差于 1 和 2。
所以要调就用自己的资料对比，"调大一点"不能想当然。（正因为不单调，重试表才给多个候选而不是无脑放大。）

读取性能相关的两个内部优化（无需配置）：

- 同一 PDF 的重复读取命中解析 memo（key = 路径 + mtime + size，LRU 3）：`as="meta"`、分批续读不再整本重解析；文件一改自动失效。
- 渲染缓存会自动清理：`%TEMP%\dsh-wen-ocr` 下同 basename 的旧 mtime 目录 >7 天、`rapid-*` 引擎暂存 >1 天会被删；`.ocr.md` / `.ocr.json` / `.read.md`（识别与兜底成果）与任何内含它们的目录、以及**当前正在用的目录**永不碰。带倍率后缀的重试目录 `<safe>-<mtime>-s<倍率>` 同样纳入同一 basename 的清扫范围（否则一次重试就留一个孤儿目录）。

  **去哪儿找它（实测 F5，两个坑，都不是 bug）**：

  1. **`%TEMP%` 在 DSH 的 `pwsh` 工具里是"每次进程一份"的** `…\AppData\Local\Temp\dsh-XXXXXX\`，
     而**插件宿主进程**写的是不带这层前缀的 `…\AppData\Local\Temp\dsh-wen-ocr`。
     在 pwsh 里直接展开 `%TEMP%\dsh-wen-ocr`，看到的是"在 pwsh 里跑的 node"那份，**永远找不到宿主写的**。
     要查宿主的渲染产物请用绝对路径（`$env:LOCALAPPDATA\Temp\dsh-wen-ocr`）。
  2. **中文文件名会被 sanitize 成下划线**：`sample-A.pdf` → `____-____.pdf-<mtimeMs|0>`
     （mtime 只留 32 位）。按名字搜必然搜不到，按"目录内 PNG 张数 ≥ 页数"搜才靠谱。

  清理合同本身实测一致：跑完整本 35 页后渲染目录**原地保留**（不存在"批尾自清"），当日新建的目录一个都没被删
  —— 因为 7 天线还没到。

## 返回边界消毒（默认生效的 bugfix）

所有工具返回值出站前统一走一条消毒管线（`finalizeToolValue`，`office_read` / `office_create` /
`office_edit` / `office_convert` 的 execute 与 render 投影都挂这一条边界）：
游离代理 / `U+FFFE` / `U+FFFF` / C0+C1 控制符（`\t\n\r` 除外）→ `U+FFFD`，再 NFC 规范化；
结构压成纯 JSON 树（`undefined` 属性剔掉、非有限数→`null`、稀疏数组补洞、循环引用置 `null`）。
账记在 `stats.sanitized` 与 `stats.sanitizeNotes`，正文尾部加脚注"**N 个非法码点已替换为 U+FFFD**"。
原则是**宁留噪不删正文**：只替换、不删除、不断句；`U+2028/2029` 是合法行分隔符，保留不动。

## 创建与编辑

- 表格/精确排版优先传结构化对象（`workbook` / `document` / `slides` / `table`），而不是 markdown 字符串。
- **生成中文 PDF 默认内嵌字体子集（第五轮实测修正，旧表述已作废）**：`writePdf` 从系统字体取
  **实际用到的字形子集**——正文 `C:\Windows\Fonts\simsun.ttc`、emoji `seguiemj.ttf`、符号兜底
  `seguisym.ttf`——Identity-H 编码 + `FontFile2` + 完整 `ToUnicode` CMap（`BaseFont` 带 `XXXXXX+`
  子集前缀）。实测（2026-09-23）：WinRT `pdf-render.ps1` 渲染中文/`☆`/`①`/`🌍` 无问号无方块、
  Edge/Chrome 正常、`office_read` 提取逐字回读（☆① 不再变 `?`）。
  ⚠ **旧文档声称"生成中文 PDF 不内嵌字体、走 Adobe-GB1 预定义 CMap（STSong-Light）、Chrome/Edge/
  Adobe Reader 正常显示"——与实测矛盾**：不嵌字体的中文 PDF 在 WinRT 下首页空白+问号（Edge/Chrome
  同样乱码），且缺 `/ToUnicode` 导致提取丢字符；该表述已删除。回退旧行为设
  `DSH_OFFICE_PDF_EMBED_CJK=0`（**不推荐**，承担上述渲染/提取双风险）；系统字体缺失时自动回落
  STSong-Light 并把原因记进 `writePdf` 的 `opts.info.notes`（产出质量门会据此拦截，见 convert 一节）。
  体积：字形子集只收**用到的 GID** + FlateDecode，所以远小于"整字体嵌入"。
  当轮的体积实测（9000 字中文报告 0.39 MB vs 未子集化 4.35 MB）属**当轮快照、未自动化**，
  见 [DEVELOPMENT.md](DEVELOPMENT.md)；套件守的是"样本 <2 MB"这条护栏。
- **CSV/TSV 写出编码（第四轮收敛）**：`.csv` 默认 **UTF-8 with BOM**（Excel 双击直接正确识别中文），
  含逗号/引号/换行的字段按 RFC 4180 自动转义（内部引号成对双写、危险字段整体加引号）；
  `encoding: "utf-8-sig" | "utf-8"` 可对文本类目标（csv/tsv/md/txt）显式控制 BOM（默认 csv 带、其余不带）。
  gb18030 等旧代码页的**写出**不是零依赖可达（Node Buffer 只内建 utf8/latin1 系），不提供；
  **读取**侧本来就会自动识别 GBK/Big5，不受影响。`office_edit` 改 CSV/tsv/md/txt 会**保留源文件原有的 BOM**
  （旧版 decode→写回会静默剥掉 BOM，"编辑一次、Excel 再打开就乱码"）。
- `office_edit` 是 zip/XML 级原地修改：`docx` 用 `replace_text` / `append_markdown` / `set_meta` /
  **`insert_image` / `append_image`**；
  `xlsx` 用 `set_cell` / `append_rows` / `replace_value` / `add_sheet` / `rename_sheet` / `delete_sheet`；
  `pptx` 用 `replace_text` / `add_slide` / `update_slide` / `delete_slide`；
  `csv/tsv` 用 `replace_value` / `replace_text` / `append_rows`；
  `md/txt` **/`html`** 用 `replace_text` / `append_text` / `prepend_text`（`.html` 是文件级文本替换，不改结构）。
- **插图（第六轮新增，第七轮 R13 补齐 docx 写出端）**：
  - **输入端**：markdown 里**独立成行**的 `![alt](path)` 解析成 `image` 块（不再是"`!` + 链接"）。
  - **docx 插图**：`office_edit operations=[{op:"append_image",path:"D:\\图.png",alt:"示意",width:400}]`，
    或 `insert_image` 配 `after:"某段文字"`（锚点未命中会**明说**并退化为文末，不静默）。
    也支持 `base64:"data:image/png;base64,…"` 或 `base64:…`。格式 PNG/JPEG/GIF/BMP；
    zip 级新增 `word/media/imageN.*` + image 关系 + `w:drawing`（关系 id 接着既有 `rId` 往后编，
    绝不撞号；命名空间就地声明，不动文档根元素）。
    **尺寸规则（R13 显式化）**：省略 `width` → 原图像素 × 72/96；显式 `width`（磅）优先；两者都受
    A4 可用宽 `451.3pt` 限制等比缩小 —— 换算过程同时进 `stats.imageSizing` 与 `notice`。
    **内容去重（R13）**：同一张图（内容 SHA-256 相同）复用同一媒体部件与同一关系 id，复用次数进 `stats.imageReused`。
  - **md / HTML → docx 也内嵌（R13）**：目标是 `.docx` 时，image 块走 `writeDocx` 的 image 分支真正内嵌
    （`stats.imageMedia` 报部件名）；**读不到 / 格式不认识**的图片退化为字面 `![alt](path)` 文本 +
    `stats.imagesSkipped` + `notice`（**绝不静默丢图**）。
  - **PDF / HTML 写出端真内嵌**：PDF 里 JPEG 走 `/DCTDecode` 原样内嵌；PNG（含**索引图 colorType 3**）、
    **BMP（8/24/32-bit）**、**GIF 首帧**（透明索引 + 交错行序）解成原始采样走 `/FlateDecode`（带 alpha 挂 `/SMask`）；
    `/XObject` 按内容 SHA-256 去重；支持 `align: "center" | "right"`（缺省恒从左边距起排，**字节不变**）。
    HTML 写 `<figure><img src>`。
  - ⚠ **`.odt` / `.pptx` / `.md` 等写出端还没有 image 分支**：遇到 image 块会**显式降级**为字面文本
    `![alt](path)`，并写进 `stats.imageFallback` + `notice`（**绝不静默丢图**）。
  - ⚠ **历史 bug（第六轮引入，R13 修掉）**：`w:drawing` 的 `wp` 命名空间 URI 曾拼错
    （`…/drawingWordprocessingDrawing/2006/main`），含图 docx 会被 Word 16.0 判"文件已损坏"（`0x80070570`），
    而 zip 结构、`rId` 唯一性、包结构体检**全部查不出来**。手上若有更早版本产出的含图 docx，重新生成即可。
- **公式重算（第六轮起；R13 扩子集 + 护栏，opt-in）**：`office_read recalc=true`（表格）。命中的公式格返回计算值，
  `stats.recalc = {formulaCells, evaluated, unsupported, errors, details:[{sheet,cell,formula,value,error}], skipped?}`。
  支持 `SUM AVERAGE MIN MAX COUNT COUNTA COUNTIF COUNTIFS SUMIF AVERAGEIF IF VLOOKUP ROUND ROUNDUP ROUNDDOWN
  ABS INT MOD LEFT RIGHT MID LEN TRIM UPPER LOWER CONCAT CONCATENATE TEXT VALUE` 与四则/比较/`&`/`%`/区域引用（含跨表）。
  **Excel 口径**：区域里的文本/布尔一律忽略（`SUM(A1:A3)` 不把 `"3"`/`TRUE` 算进去），直接写在参数里的仍换算
  （`SUM("3",TRUE)` = 4）；数值条件（`">75"`）只命中数值单元格。**不支持的函数标 `unsupported` 并保留原值**（绝不猜），
  整列引用 `A:A`、外部工作簿 `[Book1]…`、`_xlfn.` 前缀、未加载工作表、循环引用（`#CIRC!`）都带单元格地址。
  默认 false —— 不设时返回值与旧版逐字一致；大表用 `DSH_OFFICE_RECALC_MAX_CELLS` 设护栏（超限记 `skipped`，不静默截断）。
- **申论稿纸（第六轮，opt-in）**：`office_create path="x.docx" markdown="…" grid="20x25"` → 写 Word 的
  `<w:docGrid w:type="linesAndChars">`，把每行字数/每页行数钉死（`"20"` 或数字 20 也可）。
  只对 `.docx` 目标生效，其它目标/非法取值显式报错。**可见的方格线框属于页面背景**，不在本参数范围内。

## 已知边界（别硬试）

- **未实现的能力（本插件不做）**：这五项的**接口草案 / 量化依据 / 红线约束 / 重开条件整段已迁到
  [DEVELOPMENT.md](DEVELOPMENT.md)「未实现能力：草案、结论与重开条件」** —— PDF 合并 / 拆分、
  OMML 公式写入、`.html` 结构化编辑、插图落点扩展（`before` / `in_table` / 页眉页脚）、
  公式写回 `office_edit recalc_formula`。**今天要达成目的就走现成的路**：
  PDF 要合并 → 先在阅读器里"打印成 PDF"，或用别家工具（插件**不会**用"拼字节"糊过去）；
  公式 → 用 Unicode 数学符号 + 文本分数（现有 `√ ∑ ∫ ≤ ≥ ≠ ≈ ± × ÷` + `^`/`_` 已覆盖目标语料）；
  `.html` 要改结构 → `office_convert` 成 `.md` / `.docx` 改完再转回。
- **PDF 不支持原地改文本**（文本行无稳定回写位置）：要改就 `office_convert` 重出一份。
- **HTML 读取（第四轮新增）的边界**：`office_read` / `office_convert` 对 `.html` / `.htm` 走
  插件自带的零依赖 DOM 解析器（`html.js`：分词 → 隐式闭标签树 → document 模型，不用正则剥标签）。
  `as="markdown"` 保留 h1-h6/嵌套列表层级/表格/引用/代码块/外链/加粗斜体，`style` / `script` /
  表单控件 / `display:none` 隐藏子树剥离，HTML 实体（named + 数字 + Windows-1252 数字别名）统一解码；
  `as="text"` 给真正的纯文本（表格→制表符分隔行，不残留 `#` / `|` / `**` 标记）。**不做的事**：
  CSS 伪元素（`content:attr(data-no)` 之类）生成的内容不在 DOM 里、转换不包含——那是排版层；
  `colspan`/`rowspan` 只取单元格文本不做网格展开；`<ol start="N">` 的起始编号在 markdown 渲染层
  无法表达。**第六轮起 `office_edit` 支持 `.html` / `.htm`**：走**文件级文本替换**
  （`replace_text` / `append_text` / `prepend_text`），不改结构、不重新排版；要改结构就 convert 成 md/docx 再改。
  **第五轮起写出端也有了**：`office_create` / `office_convert` 的目标可以写 `.html`（语义化 HTML5：
  `<h1>`-`<h6>`、`<p>`、嵌套 `<ul>`/`<ol>`、`<table>`（`header:true` → `<thead><th>`）、`<blockquote>`、
  `<pre><code>`、`<hr>`、行内 `<strong>`/`<em>`/`<u>`/`<s>`/`<code>`/`<a>`，UTF-8 + `<meta charset="utf-8">`，
  中文与 emoji 原样输出、不转数字实体）。往返实测：md → .html → md 与 md → md **逐字一致**；
  Word 16.0 能直接打开产出的 .html。已知不覆盖：HTML 没有分页概念（`pagebreak` 写成带 class 的 `<div>`），
  图片块写成真 `<figure><img src="…" alt="…">`（第六轮起；有路径才写 img，没有路径保留占位文本）。
- **过门 ≠ 可直接引用（自 WB 侧第七轮 §7.1 移植，重要）**：通过质量门只说明**质量门放过**，
  **不等于文本无噪声**。数字 / 正负号密集的表格页（4 列以上的"增速 / 同比(%)"列）**务必核对原 PDF**，
  典型症状是**负号被丢掉**（`-9.8 → 9.8`、`-1 → "[-"`），也存在把选项数字看错的可复现反例。
  当轮的抽样量化（510 个数据格 4.9% 出错、密集表 10–15%、具体反例页码）属**未自动化的单机实测**，
  见 [DEVELOPMENT.md](DEVELOPMENT.md)。
  **置信度筛不掉这类噪声**（0.991 的页也能错 12.5%）——本侧的**视觉复核升级同样防不住**：
  它治的是"引擎没认出字"，不治"字认对了、但负号丢了 / 分组错位 / 小数点漂移"这类**版面级**噪声，
  所以**不要**把"升级视觉"当成这条的解法。表格数字页的结论**必须回原 PDF 核对**，
  或明确标注"数字未核"。与「填空槽位」同属一类：**能读到 ≠ 读对了**。
- `pptx` 的 `update_slide` 会重写该页，原页上的图片/复杂图形会丢；其他页不受影响。
- `.wps` / `.et` / `.dps` 若是金山新版专有二进制，解析不了；请另存为 `.docx` / `.xlsx` / `.pptx`。OLE2 变体可直接读。
- 只设**权限密码**（禁复制/打印）的 PDF 会透明解密正常读取，不改写原文件；需要**打开密码**的不会去猜，
  会明确提示。加密覆盖 RC4-40 / RC4-128 / AES-128 / **AES-256（R5/R6）**（第六轮起）：
  R5 一次 SHA-256、R6 走 ISO 32000-2 算法 2.B 的迭代哈希；候选密钥必须通过 `/Perms` 已知明文校验
  才被采纳，校验不过就报"需要打开密码"（**口令候选只有调用方给的那组，默认空串，绝不试错口令**）。
  **第七轮 R13 的判别实验**：同骨架的自造 **AES-128** 夹具能被 **WinRT（微软实现）打开并渲染** ⇒ 骨架正确；
  因此本机 AES-256 夹具的 WinRT `OPEN_FAIL` 归因于"**WinRT 不支持 R5/R6**"，而不是 V5 字典写错。
  "与第三方逐字节一致"**仍未取得** —— 把真实 R5/R6 样本路径设进 `DSH_OFFICE_TEST_AES256_PDF`，一条命令即可复验。
- **填空槽位（`________`）本来就在文字层里 —— 第三轮之后能正常读到**（旧表述"下划线不带任何字符"是错的）：
  - **文字层路线**：填空**是真实的 `_` 字符**（WPS 产出器把下划线画成一串 U+005F 字形，CID 例：`0x0042`）。
    它"看起来消失"是**两个先存解析缺陷叠加**的结果，都已在第三轮修掉：
    ① `expandObjStms` 少加 `First` → ObjStm 里对象整体前错、页 `/Resources` 错位、CID 被拆成单字节 → 中文整本乱码；
    ② `mul()` 读越界的 `b[6]` → 坐标全 NaN、`Math.round(NaN/2.5)` 成为同一个 Map key → **每页塌成一行**。
    两者叠加时 `_` 既在乱码里、又被压成一行，所以"看不见"。
    实测：病灶样本修后 305 段填空、干净样本 700 段填空，**逐段不变**。
  - **OCR 路线**：识别引擎的产物里下划线**确实常常丢**（不输出装饰性 `_`），
    扫描件仍需人工或后处理补槽位。别把两条路的结论混用。
  - **呈现（opt-in）**：`DSH_OFFICE_BLANKS=1` 时 `office_read` 的 PDF 分支会把 `_{3,}` 渲成
    `<span class="blank">＿＿＿＿</span>`。它**只管呈现**、不 gate 任何正确性修复；`as="meta"` 不受影响，
    `as="text"` 按规格不生效（要原文），对 OCR 路线天然无效。
- **写 PDF 的成片私用区码点**：第三轮修掉过 `writePdf` 换行循环里的一个真实死循环（会把宿主事件循环占死），
  现在成片私用区 / 汉字 / 超长 ASCII 串都能正常写。历史诊断与回归面见 [DEVELOPMENT.md](DEVELOPMENT.md)。
- **缓存迁移（页码 → 哪张纸）**：`.ocr.md` 的成果**按页码**记，而页码有两个互不相干的来源 ——
  文字层走插件自己的 `pages()`（PDF 页面树的枚举顺序），OCR 渲染走 `pdf-render.ps1` 的 WinRT
  `GetPage(n-1)`（文件里声明的物理页序）。两者一旦不一致，`ocrMap.get(page)` 就会把**另一张纸**的
  识别结果贴到这一页上，而且是静默的。第三轮修掉 `expandObjStms` 的 off-by-`First` 正是这种情况：
  病灶样本①的 `/Pages` 主体原本被读成 ObjStm 的头部索引表 → `/Kids` 解析不出 → `pages()` 落到
  "扫描全部 `/Type/Page` 对象"的兜底，页序变成对象号升序（封面从第 1 页跑到第 34 页），
  第 1–33 页的 OCR 成果整体错一位。
  → 所以 sidecar manifest 现在带 `parser:` 版本戳；**缺失或不等于当前版本 = 整份视为未覆盖**
  （页正文一概不回收），`stats.ocrCacheStale` / `stats.ocrCacheNote`、`read` 的脚注与
  `convert` 的拒绝文本都会明说原因，**绝不静默复用错页文本**。重新识别需要的页即可，旧文件可直接删。
  任何再次改变"页码 → 页对象"映射的改动都必须把 `PDF_PARSER_VERSION`（`index.js`）+1。
- **`.pptx` 现在可以直接交付**：第七轮定位并修掉了"PowerPoint 16.0 拒开（`0x80070570`）"的根因 ——
  `notesMaster` 必须有自己的独立主题部件（不能与 `slideMaster` 共用 `theme1.xml`）。
  不必再先 `office_convert` 成 `.docx`/`.md`，也不必借 PowerPoint 模板另存。
  诊断链（zip 层 / OPC 不变量 / 逐个部件的单变量排除）与"改回共用主题必须仍 FAIL"的负向控制见 [DEVELOPMENT.md](DEVELOPMENT.md)。
- **缓存身份（第十七轮 R16）**：见上文第 4 步的「身份行」——同一路径内容变了、或同名文件换了目录，
  旧 `.ocr.md` 都整份作废（`stats.ocrCacheStale.reason` 说明原因），**绝不静默复用错页文本**。
  **已知限制（第十八轮 R18 记录）**：路径指纹是**对规范化后的源路径字符串**做 SHA-256（`pathKeyOf`，
  纯字符串运算、零 IO），**不做短名（8.3）/ 符号链接 / 网络盘别名的归一化**。所以同一个文件经不同路径形态
  访问（`C:\PROGRA~1\…` vs `C:\Program Files\…`、软链 vs 实体路径、`\\?\` 前缀）会被当成**不同来源** ——
  表现是"缓存没命中、白重识别一次"，**绝不会**误用别人的缓存（宁可重跑，不可错用）。
- **PDF 产出质量门与 `stats.pdfQuality`（第十八轮 R18 补全口径）**：写到 `.pdf` 时返回
  `stats.pdfQuality = { embedded, renderCheck: 'pass' | 'skipped', firstPageBytes }`；含插图时另加 `images`
  （真正嵌进去的张数）与**条件性**的 `imagesSkipped`（读不到 / 格式不认识时的 `{name, reason}` 数组，同时进顶层 `notice`）。
  渲染级抽检用的 PDF 副本落在 **`%TEMP%`**（`DSH_OFFICE_PDF_GATE_DIR` 可覆盖），渲染完立即删除，
  **校验通过后才**原子发布到目标 —— 所以**目标目录在非 `%TEMP%` 位置也能产出 PDF**（R18 修的正是这条：
  旧版把副本写在目标同目录，而本机 WinRT 只读得到 `%TEMP%` 下的文件 → 任何非 `%TEMP%` 目标的 PDF 产出 100% 失败）。
- **写盘原子性（第十七轮 R16）**：`office_create` / `office_edit` / `office_convert` 与全部 sidecar 都是
  "目标同目录唯一临时件 → 完整写入 → 成功后替换"。写入或发布失败时**原文件逐字节不变**，抛「写盘失败｜四要素」；
  不会留下半截文件（`.part` 临时件会被清理）。"编辑把原件写坏"这一类事故不再可能。
- **ZIP 解压护栏（第十七轮 R16）**：单条目 **256 MiB**、整归档累计 **512 MiB**
  （`DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` / `DSH_OFFICE_ZIP_MAX_TOTAL_BYTES` 可覆盖，`0` = 关闭该上限）；
  CRC32 默认校验（`crc !== 0` 才比对，`DSH_OFFICE_ZIP_CRC=0` 关闭）。
  高压缩比包会被**解压前**拦下并给可读错误 + 下一步；越界偏移、说谎的长度、损坏的中央目录也不再抛宿主级 `RangeError`。
  正常的大图 / 大表文档不受影响（实测 333 份真实文档 0 误拒）。
- **ZIP 写出端的 ZIP64（第十八轮 R18 新增 / 第十九轮 R19 定性）**：条目数 ≥ `0xffff`、尺寸或本地头偏移
  ≥ `0xffffffff` 时按规范写 ZIP64 扩展字段（ID `0x0001`）+ EOCD64 + 定位记录；**不需要 ZIP64 的包**
  （也就是所有真实的 OOXML / ODF 包）字节布局与旧版**逐字一致**。**已知限制**：">4 GiB **数据尺寸**"
  这条分支**没有端到端产物验证**（`makeZip` 要真的收 ≥4 GiB 输入缓冲、`concatBytes` 还要再复制一份
  ⇒ 峰值 ≈8.0 GB；本机 15.9 GB 总内存、探针运行时空闲 2.9–3.6 GB）。**但声明格式本身已被主流实现接受**：
  `.NET ZipArchive` 把我们的 ZIP64 声明读成 `Entry.Length = 4 GiB+`，真实 4 GiB deflate 条目的前 1 MiB
  也能正确解出。触发条件与影响面（以及"空闲内存 ≥10 GB 即可端到端复现"的探针）见
  [DEVELOPMENT.md](DEVELOPMENT.md) R19 §A。
- **大输入不会栈溢出（第十八轮 R18 修 / 第十九轮 R19 量化）**：`textQuality` 现在**零正则**
  （两条整串正则都改成手写扫描）。旧写法在"单个巨大匹配"上抛 `RangeError`，本机阈值落在
  **(5 MiB, 6 MiB]**（逐档表 + `--stack-size` 100–8000 KB 矩阵见 DEVELOPMENT.md R19 §D1）。
  这个阈值**不是可移植常数**（它是 V8 对单个超长匹配的内部限制，与 JS 调用栈大小无关）⇒
  **改动这条路径时不要引入整串正则**（`match(/…/g)` / `matchAll` / `split(正则)` 作用在整本文本 / 单个超长行上）。
- **工具参数（第十七轮 R16）**：`grid` 接受数字与字符串；`office_edit.operations` 的元素级参数错误带
  `operations[N]` 下标定位（缺 `op` / `op` 不认识 / 缺必需字段都在**参数层**报错，不会再落到执行期写 `undefined`）。
- **`as="json"` / `as="meta"` 的 `content` 是 JSON 文本字符串**（不是对象）：需要 `JSON.parse` 一次；
  这两种形态不写 sidecar、不带 `nextOffset`，超限时按页丢尾部并给 `stats.truncateNote`，但**永远是合法 JSON**。
- 交付文档若需署名，按 `dsh-badge` 技能在文末加官方徽章（121×20，不换色、不改 logo）。

## 排查

- **症状：工具报 `value is not lossless JSON`（含正文的 `office_read` 整条被拒收，`as=meta` 却正常）**
  → 这是**返回值结构**问题，不是正文内容问题，**不要**去改 `--ocr`。
  成因：host 侧在工具体返回值离开 `execute()` 之前做无损 JSON 判定（DSH 走
  `@deepseek-ai/dsh-util-values` 的 `snapshotJsonValue`），拒绝 `undefined` 属性 / 非有限数 /
  稀疏数组 / 非纯对象 / 循环引用。原 `stats.ocrCovered: undefined` 这类"有则给、无则 undefined"
  的写法会被判不可序列化 → 整条结果作废（连最窄的"单页、干净输出"路径也过不了）。
  已修：所有出站（含 render 投影、catch 分支）统一走 `finalizeToolValue()` / `sanitizeThrown()`。
  自检命令：`node repro.mjs`（🟢 GREEN = 返回值 host 一定收；退出码 0/1 即绿/红）。
- **症状：调用报错、但怀疑"识别其实已经跑了"** → 先看同名 `.ocr.md` sidecar 的 `covered:` 行；
  **OCR 成果其实已在盘上**。DSH 宿主里 sidecar 对"会话只读"的附件目录**写得进去**
  （WorkBuddy 侧"附件目录会写失败"的经验在 DSH 不成立），所以工具报错时 sidecar 往往已经落盘并在推进。
  直接 `read` 那个 `.ocr.md`、或 `Select-String -Pattern 'covered:'` 看进度，比反复重试省事。
  补丁后这条已是官方路径：错误文本自身就会给出 sidecar 路径 + `covered` / 缺页区间（无需再手动发明）。
- **症状：中文 PDF 正文乱码**
  → 先按「读取的正确顺序」第 3 步分清成因：**自家解析器错位**（第三轮已修，修完直接可读、别再 OCR）
  还是**文档真缺 `ToUnicode`**（本地没有映射表，只能 OCR）。这条**只治乱码，不治理拒收**（症状不同、修法不同）。
  确认是真缺映射后现已收进自动降级链：质量门判为乱码的页会被自动识别；识别不可用则落 `<文件>.read.md` 并用
  返回值里的 `notice` 给出可复制参数。想手动重跑就照 notice 里的串，或直接
  `office_read path=… ocr="always" ocrEngine="local" pages="1-5"`。
  ⚠ 若 `qualityGate.reasons` 只指控**目录 / 索引页**这类"页码密集、正文稀疏"的页，先排除**点前导假阳性**：
  第三轮起连续 ≥3 个 ASCII 句点算排版装饰、不计入可见字符（判据没有放宽 —— 私用区 / 替换字符 /
  控制符三条仍各自独立命中，纯点号页也不会被判乱码）。
- **症状：`.ocr.md` 明明在盘上，工具却说"covered: 无"**
  → 大概率是 `parser:` 版本戳缺失或不符（第三轮之前写下的 sidecar 一律如此）。
  这不是 bug 而是迁移协议：旧缓存的页号可能指向另一张纸（见「已知边界 / 缓存迁移」）。
  看 `stats.ocrCacheStale` / `stats.ocrCacheNote` 确认，然后重新识别需要的页。
- **症状：`office_convert` 报「转换拒绝｜四要素」且没产出目标文件**
  → 这是**故意的**：源 PDF 的文字层整本乱码、而本地 OCR 缓存又不完整时，继续转换只会把
  不可读正文写成目标文件（旧版会把含 NUL 的垃圾落盘且零告警）。照错误里的 `下一步` 先
  `office_read … ocr="always" ocrEngine="local" pages="1-20"` 生成 `.ocr.md`
  （分批跑到覆盖全部乱码页），再重跑转换。错误文本里的 sidecar 路径与 `covered` 状态可照抄执行。
- **症状：某页 OCR 显示"未完成"**
  → 先看 `stats.ocrRetried` / `stats.ocrRetryScale` / `stats.ocrFailedPages`：插件已按
  `DSH_OFFICE_OCR_RETRY_SCALES`（默认 `2,1.5,1`）自动换倍率重试过。全部倍率都失败才记 `failed`，
  原因里会写"已试过 scale=…"。想换别的倍率组合就设 `DSH_OFFICE_OCR_RETRY_SCALES=3,2`（空串 = 关闭重试）。
- **症状：附件目录 / 用户目录被 `.ocr.md` 污染** → 设 `DSH_OFFICE_CACHE_DIR=<目录>`，
  所有 sidecar 统一落到那里（只读场景一样能缓存）。
- **症状：`office_create` 产出的 docx 被 Word 拒开（0x800A1401「检查文档或驱动器的文件权限 / 用文本恢复转换器」）**
  → 这是**包级 Content-Type 写错**，跟文件权限、跟 styles.xml 里的 `pPr` 顺序都无关。
  `[Content_Types].xml` 里 `docProps/app.xml` 的 Override 必须是**包级**类型
  `application/vnd.openxmlformats-officedocument.extended-properties+xml`；写成
  `…wordprocessingml.extended-properties+xml`（pptx 的 `…presentationml.extended-properties+xml`）时
  Word 16.0 直接拒开，而同内容的 `.odt` 能开（odt 没有这层）。第五轮已修（`docx.js` / `pptx.js`）。
  定位手法（别再从 styles/numbering 猜起）：① 把坏包与 Word 自产对照包的同名部件逐字 diff；
  ② 每次只改一处重打包，用 Word COM 开一次。实测：只有把 app.xml 的 Override 类型改对（或整块删掉
  app.xml 部件）才能打开；把 app.xml / core.xml 的内容换成 Word 自产的**照样失败**——内容无辜，类型是根因。
  `test.mjs` 已加"Word 16.0 打开 + ExportAsFixedFormat"冒烟断言（Word 不可用按套件惯例 skip）。
- **症状：中文 PDF 在阅读器里空白 / 问号，或提取出的 `☆①` 变成 `?`** → 按顺序看三层：
  1. **字节级**：`/FontFile2` 计数为 0 就是没嵌字体（旧版 STSong-Light + Adobe-GB1 预定义 CMap 那条路）；
  2. **渲染级**：`pdfOutputGate` 已按"首页整页黑像素 < 40 个"判"渲染为空白"并**拒绝落盘**（四要素错误）；
  3. **人工复核**：`powershell -File pdf-render.ps1 <pdf> <outDir> "1"`（位置参数：pdf、输出目录、页范围）
     渲染首页，再用 modlens 读图看有没有 `?` / 空白方块。
  `DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` 只跳**渲染**级抽检，字节级照样拦；`DSH_OFFICE_PDF_EMBED_CJK=0`
  是"回退旧的不嵌字体行为"，落地前同样被字节级门槛拦下（等于显式复现缺陷用）。
  另注：彩色 emoji 的 `COLR/CPAL` 表对 PDF 无用却极大，第五轮起在子集化时剔除 ——
  中文/emoji 的渲染与提取结果均无变化（该表多大、省了多少体积属**当轮实测**，见 [DEVELOPMENT.md](DEVELOPMENT.md)）。
- **逃生通道（产出端坏了、但要马上交文件）**：md → `office_create(path="x.odt", markdown=…)` →
  PowerShell Word COM（`$w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.DisplayAlerts=0;
  $d=$w.Documents.Open($odt,$false,$true); $d.ExportAsFixedFormat($out,17)`）导出 PDF，
  系统会自己对字体做子集嵌入，实测可用。Word COM 姿势：先
  `Get-Process WINWORD | Stop-Process -Force` 并删 `%APPDATA%\Microsoft\Word\*.asd`（残留恢复文件会捣乱），
  **不要调 `$word.Quit()`**（首次运行/恢复对话框会阻塞），结束直接 Stop-Process；
  PS 5.1 脚本必须 ASCII-only（无 BOM 的 UTF-8 会被按 ANSI 解释）。
- 行为异常时跑 `node test.mjs`：端到端检查（数量随用例增减，看最后一行的 `N checks`）。
  在插件目录里直接跑也没关系——输出目录不可写时会自动回退到系统临时目录；
  `DSH_OFFICE_TEST_OUT=<目录>` 可指定产物目录，`DSH_OFFICE_TEST_SCAN_PDF=<某.pdf>` 可钉住 OCR 样本
  （不指定时自动挑 `~/.dsh/attachments` 里**最近**一个 >5 MiB 的 PDF，并**拷进产物目录**再冷启动，
  不会删你附件旁边真实的 `.ocr.md`）。
  `DSH_OFFICE_TEST_BROKEN_PDF=<某.pdf>` 现在只用于**①「修后文字层可用」+ 页序/换倍率重试**那一段
  （第三轮起乱码端到端改用**自造夹具**，零样本依赖），嫌慢可 `DSH_OFFICE_TEST_SKIP_BROKEN_OCR=1` 跳过；
  `DSH_OFFICE_REPRO_PDF=<路径>` 同理作用于 `repro.mjs`。`DSH_OFFICE_BLANKS=1` 打开填空呈现。
- **测试夹具与内部通道变量（仅测试用；不设时测试行为与旧版逐字一致）**：
  - `DSH_OFFICE_TEST_CID_PDF=<某.pdf>`：`test.mjs` 的 **CID 乱码 / 断行夹具**（用例名前缀 `样本 PDF：` 与 `断行：②`）。
    默认是一段**硬编码的附件绝对路径**（`…\attachments\v1\files\9c\…\sample-B.pdf`）。该文件在不在决定两段
    是**真跑**还是跳过，且两段口径不同：第一段（样本 PDF 正文完整性 / 不误触降级）夹具缺失时 **判红**
    （对应断言传 `false`），第二段（断行 / CJK 19879 基线）夹具缺失时 **跳过并算过**（对应断言传 `true`）。
    要复现这两段就把它指向手上的 43 页 CID 样本。
  - `DSH_OFFICE_TEST_EXPAND_ARCHIVE=1`：让 `test.mjs` 额外做一次**真·`Expand-Archive` 解包** ZIP64 包
    （65536 个条目）。默认**跳过** —— R19 已把它**跑到底**（`status=0`、65,536/65,536 个文件、
    2,180,154 ms ≈ 36.4 分钟 ⇒ **33.27 ms/文件**；同一包 `.NET ZipFile::ExtractToDirectory` 只要
    1.145 ms/文件，**慢 29×**），
    确认 `Expand-Archive` 的每文件开销比 `.NET` 高一个量级；所以套件默认改用等价的 `.NET ZipArchive`
    （`Expand-Archive` 内部就是它）自动验证"能打开并列出 65536 条"。
    **什么时候该手动跑这一条**：**发版前**、或改动了 `makeZip()` 的 ZIP64 分支 / EOCD64 布局时 ——
    它是"我们产出的 ZIP64 包能被**官方命令行工具**真解包"的唯一端到端证据（独占机器跑，约 20 分钟）。
  - `DSH_OFFICE_TEST_REAL_4GIB=1`：让 `test.mjs` 额外造一个**真实 >4 GiB 的合法 ZIP64 条目**
    （流式 deflate，内存 O(1)，约 15 秒 → 4.17 MB 压缩数据），断言 `.NET ZipArchive` 能把它读成
    `Entry.Length = 4 GiB+`、且前 1 MiB 解压内容正确。默认跳过（慢，且只对 ZIP64 尺寸分支有意义）。
  - `DSH_OFFICE_TEST_NET_DIR=<目录>`：在**网络盘 / OneDrive 目录**上跑"写临时件 → rename 覆盖"的原子写用例。
    默认跳过 —— 本机没有可用同步盘（`net use` 无映射网络驱动器；`~/OneDrive` 目录未登录 / 未同步），
    **不用本地目录冒充网络盘**；有环境的机器上一条命令即可验证。
  - `DSH_OFFICE_TEST_OTHER_VOL=<某卷上的目录>`：给"跨卷 `rename` → `EXDEV`"用例指定一个与 `%TEMP%`
    **不同卷**的可写目录（不设时按 `cwd` → 插件目录 → `D:\` → `E:\` 依次找，都不可写就打印跳过原因）。
  - `DSH_OFFICE_TEST_NONTEMP=<目录>`：可选的**非 `%TEMP%`** 暂存目录，供 `R18-质量门：非 TEMP 目标目录…`
    用例验证"目标目录不在 `%TEMP%` 时 PDF 也能产出"。不设时按 `~/.dsh/tmp/…` → `cwd/…` 依次找，
    且**刻意排除插件目录自身**（免得把测试产物写进目标目录的 `test-out/`）；都不可写就打印跳过原因。
  - `DSH_OFFICE_REPRO_KEEP=1`：**任何非空值**（truthy，`repro.mjs` 里读这个变量）即保留 `repro.mjs`
    拷贝 / 下载出来的临时样本目录（`repro.mjs` 的 `TMP` 常量，位于 `%TEMP%\dsh-wen-repro`）；
    未设 = 跑完 `rm -rf TMP`（旧行为逐字一致）。
  - `DSH_OFFICE_RENDER_QUIET=1`：**只认精确字符串 `'1'`**（`pdf-render.ps1` 里两处 `-ne '1'` 判断，
    与前面所有 truthy 变量语义不同）—— 设成 `1` 才让脚本不打印 `PDF pages = N` / `SAVED …` 两行 stdout；
    `=true` / `=yes` 都不生效。**插件自己从不设置它**（`index.js` 的 `runRenderScript()` 只往渲染子进程 env 塞
    `DSH_OFFICE_RENDER_SCALE`），也**不读**它的 stdout（成功判据是"PNG 在不在盘上"，失败详情走 stderr）
    ⇒ 它是**未接线的降噪开关**，只有手工跑 `pdf-render.ps1` 时有用，对工具产出零影响。
  - **`DSH_OFFICE_NATIVE_SRC` / `DSH_OFFICE_NATIVE_PDF` / `DSH_OFFICE_PPTX`：写入侧 IPC 通道，不是配置项，外部设了无效。**
    它们是 `test.mjs` **自己注入**给 PowerShell 子进程、供脚本用 `$env:…` 读的裸值（"路径只走环境变量，
    避免中文 / 引号在命令行上被切碎"）：每次 Word 冒烟都写 `DSH_OFFICE_NATIVE_SRC`，并按需写 / 删
    `DSH_OFFICE_NATIVE_PDF`；PowerPoint 冒烟写 `DSH_OFFICE_PPTX`。三者在**插件代码里没有任何读者**
    （唯一的读点在这些测试内嵌的 PS 脚本内部，`grep -n 'DSH_OFFICE_NATIVE_SRC' test.mjs` 即可看到），
    属**定义即消费**的内部钩子 —— 外部 export 它们不起作用。
- 本地 OCR 不工作时：`node -e "import('./rapidocr.js').then(m=>console.log(m.findEngine()))"` 看引擎是否被发现
  （期望 `vendor/RapidOCR-json_v0.2.0`）；`DSH_OFFICE_OCR_DISABLED=1` 可临时退回纯视觉路径，
  `DSH_OFFICE_RAPIDOCR_DIR` 指向别的引擎目录，`DSH_OFFICE_OCR_MIN_SCORE` 调质量门松紧。
- 引擎探测按**候选链逐个往下找**，顺序为 `DSH_OFFICE_RAPIDOCR_DIR` → `<插件目录>/vendor`（含 `vendor/RapidOCR-json_v0.2.0` 一层）
  → `~/.dsh/ocr`（含 `RapidOCR-json_v0.2.0` 一层）→ `PATH` 中的同名版本目录（`rapidocr.js` 的 `engineCandidates()`）。**某个候选缺
  `RapidOCR-json.exe` 或它 `models/` 里的四个默认模型，只会跳过该候选、探测继续往下找**（`findEngine()` 里两处 `continue`），
  插件不报错；**只有候选链全部不合格（或设了 `DSH_OFFICE_OCR_DISABLED`，`findEngine()` 的禁用分支）
  才退回视觉桥**，此时来源标注显示"视觉模型识别 · 本地引擎不可用"（`index.js` 的 `ocrSourceLabel()`），不会说"本地没把握"。
- 另注：`DSH_OFFICE_OCR_MODELS` / `_DET` / `_CLS` / `_REC` / `_KEYS` **只影响运行侧参数、不参与这份可用性检查**
  （可用性只看硬编码的 `DEFAULT_MODELS`（`rapidocr.js`）四个文件名与固定子目录 `models/`，判断在 `findEngine()`）。
- **改了插件源码要热生效**：`~/.dsh/cordis.patch.yml` 里 `tool-office` / `skill-office` 两行都用
  **绝对路径**挂载（`${DSH_HOME}/plugins/dsh-wen/index.js` 与 `…/skill.js`），
  **没有 `?v=N` 缓存爆破** —— 编辑任何插件源码（含 `html.js` 之类兄弟模块）后，都要**重启 profile**
  （关掉再起 dsh）才生效。裸包名 `dsh-wen` 在本机 desktop profile 下解析失败，patch 文件里已注明原因。
- 卸载：删掉 `~/.dsh/cordis.patch.yml` 里的 **`tool-office` 与 `skill-office` 两行**（或整个文件），
  再删 `${DSH_HOME}/plugins/dsh-wen/` 目录即可；也可只把这两行 `disabled: true` 临时停用。

## 延伸阅读（开发参考）

逐轮修复史、内部实现细节、诊断过程与实测附录都在 [DEVELOPMENT.md](DEVELOPMENT.md)：

| 想看什么 | 去哪 |
| --- | --- |
| 第十八轮（R18）改了什么、缺陷根因、兼容性影响、剩余风险 | DEVELOPMENT.md「第十八轮（R18）」 |
| 第十七轮（R16）改了什么、兼容性影响 | DEVELOPMENT.md「第十七轮（R16）」 |
| 更早的逐轮修复注记（第五/六/七轮 …） | DEVELOPMENT.md 同名小节 |
| **未实现能力的草案 / 量化依据 / 重开条件**（PDF 合并拆分、OMML、`.html` 结构化编辑、插图落点、公式写回） | DEVELOPMENT.md「未实现能力：草案、结论与重开条件」 |
| **从 README 迁出的实现细节 + "当轮实测、未自动化"的数字**（栅格化/识别两级、长页分带、逐档渲染倍率表…） | DEVELOPMENT.md「从 README.md 迁出的实现细节与"当轮实测、未自动化"数字」 |
| **从 SKILL 迁出的历史注记**（writePdf 死循环、notesMaster 独立主题…） | DEVELOPMENT.md「从 SKILL.md 迁出的历史注记」 |
| 渲染缓存位置 / 清理规则 | 本文件「调优环境变量」+ DEVELOPMENT.md |
| 挂载方式的历史（裸包名 → 绝对路径 → 绝对路径）与 locale 友好名机制 | DEVELOPMENT.md「挂载方式的历史」 |
| 测试夹具变量、内部通道变量 | 本文件「排查」最后一节 |
| 逐轮修复史、内部实现细节、诊断过程与全部实测附录 | DEVELOPMENT.md（全量） |

> 定位方式：本文档**不再用 `index.js:1234` / `test.mjs:1234` 这类行号定位**（行号每轮都会整体漂移，
> 第十七轮 R16 就漂过一次）—— 统一用**符号名**：`grep -n '<函数名>' index.js`、
> `grep -n '<用例名前缀>' test.mjs`、`grep -n '<变量名>' rapidocr.js`。符号名比行号稳。
> 当前行为一律以 `index.js` + `test.mjs` 为准；本文档与实现冲突时以实现为准。
> README.md 只装快速入门 / 格式能力矩阵 / 常见用法 —— 那是给"第一次用"的人看的，不是契约全集。
