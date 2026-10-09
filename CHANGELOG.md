<!-- 本项目原名 dsh-office，后更名为 dsh-wen；本文档保留当时的名称与实测记录。 -->

# CHANGELOG — dsh-office（DSH 版）

> 代码同步源：`<home>/.workbuddy\skills\dsh-office\index.js`（WorkBuddy 版，只读）。 ⚠ **本句是历史批次口径**（对应 L1764「以 WorkBuddy 版为源同步」、L1946 / L2002「自 WorkBuddy 版同步落地」）；**现行关系非单向**：两侧**双向分叉**（L1724），回移一律**逐文件比**、**禁止整目录覆盖**（L214 / L580 / L1131 / L1277 / L1710）。
> 下面第一条是本目录（DSH 宿主）自己的账；后三条是**随同步一并落地**的 WorkBuddy 侧历史，保留以备追溯。 ⚠ **本计数已过时**（原文保留）：现存「自 WorkBuddy 版同步落地」条目只有文末**两条**（L1946、L2002）；其余均含本侧自己的账，如无编号的「本轮前提更正」L1856。
> **修订轨迹不删**：已作废的判断保留原文，只在旁边标注"后来怎么变的"。
> 行号引用一律用全文件行号，或显式注明是条目内相对行号。

## 2026-10-09 — v1.1.0：手写稿能力边界 + 视觉桥隐私告知 + provider 命名冲突修复

> 版本从 `1.0.1` 升到 `1.1.0`（minor：有新增能力与新增可选参数，无破坏性改动）。
> **调用契约零变化**：工具名仍是 `office_read` / `office_create` / `office_edit` / `office_convert`，
> 环境变量仍全部是 `DSH_OFFICE_*`（本轮**只增不减**：新增 `DSH_OFFICE_VISION_MIN_SIDE`、
> `DSH_OFFICE_VISION_NO_NOTICE`）。既有缓存目录名与输出格式均未变。

### 1. 手写稿与结构化版面：能力边界（实测结论，写进 `SKILL.md`）

**新增一节"手写稿的能力边界"**，数据来自 2026-10 的实测（样本：鲁迅《壬子日记》1912-05-05 行草竖排手稿，
真值取维基文库排印本）：

| 引擎 | 单页耗时 | 手写稿准确率 |
| --- | --- | --- |
| 本地 RapidOCR | ~0.5 秒 | **15.0%** |
| 视觉桥（千问3.8-flash） | ~10–20 秒 | **54.4%** |
| 视觉桥（hy3） | 116 秒 | 38.7%（同输入两次结果不同） |
| 视觉桥（glm-5.3-flash） | 88 秒 | **空输出**（推理占满 token） |

**根因结论**：失败的根因不是"手写"，而是"**版面结构**"——工整楷书 + 横排的答题卡（14.5%）与
行草 + 竖排的手稿（12.2%）几乎无差别；检测模型把作文方格网格整片判为非文本结构跳过。
**分类维度应是"有没有结构性干扰"（方格/表格线/竖排），而非"手写还是印刷"。**

**视觉桥"自信编造"风险的实证**：低清图（318×550）下，原图「五日上午十一时舟抵天津」
被输出成「五月二十一日 晴 晨起，整理书案，阅《资治通鉴》数页…」——**后一段在图上完全不存在**，
是模型按"这是一页旧日记"的常识生成的，读起来完全通顺可信。同类：真人名被填成错误的具体名词。
分辨率是关键变量：同一张图同一模型，635×1100 得 54.4%，318×550 跌到 3.7% 且整段编造。

**方法论红线（一并写入）**：评估 OCR 时**真值必须来自被测系统之外**（权威排印本 / 人工校对稿），
**绝不允许拿被测模型自己的输出当真值去测它自己**——那是自证循环，必然得 100% 却毫无信息量。
本项目实测踩过：用同一视觉模型的转写当真值去测本地 OCR，得出"本地 3%"；换排印本后本地实际是 12%。

### 2. `DSH_OFFICE_VISION_MIN_SIDE`：送视觉桥前的分辨率门（新增，默认 1024）

低于下限时先按 2×/3× 重渲染，仍不达标则**明确报错、拒绝送图**（而不是把低清图交给模型诱发编造）。
升倍率会记账进 `stats.upscaled` 与脚注。设 `0` = 关闭该护栏（回到旧行为，自担编造风险）。
配套新增导出函数 `visionMinSide()`、`layoutGapWarning()`（版面结构性干扰预警）。

### 3. 视觉桥隐私告知（新增，默认开启）

**视觉桥会把页面图上传到所配置的第三方 API 服务，原图离开本机。** 本轮把这条从"文档里写一句"
提升为**运行时告知**：只在真正发生过上传时，本次输出的脚注里会出现一行

> ⚠ 隐私：以上识别使用了**第三方视觉 API**，页面图已上传离开本机。涉密或不宜外传的文档请改用
> `ocrEngine:"local"` 禁用回落。

只在首次真正上传时记一条（多页/多分带不刷屏）。设 `DSH_OFFICE_VISION_NO_NOTICE=1` 可隐藏。

### 4. 修复：skill provider 命名冲突（`skill.js`）

`PROVIDER_NAME` 由 `dsh-wen` 改为 `dsh-wen-tools` —— 原名与桌面版 `@deepseek-ai/dsh-skill-office`
注册的 `dsh-office` provider 在**同一全局层**撞名，`SkillsRegistry.registerProvider` 会因重名抛错
（"a skill provider named ... is already registered"），导致该行在桌面端整行挂载失败。
**技能名（`SKILL.md` 的 `name:`）未动**，`skill("...")` 的调用方式不变。

### 5. 仓库卫生（不影响功能）

- `test.mjs` / `repro.mjs`：**移除全部硬编码的个人绝对路径**（12 + 1 处），改为
  环境变量驱动（`DSH_OFFICE_TEST_REAL_DIR` / `DSH_OFFICE_TEST_CID_PDF` /
  `DSH_OFFICE_TEST_ENC_PDF` / `DSH_OFFICE_TEST_REAL_XLSX`）＋ `samples/` 通用占位。
  样本缺失时按套件既有惯例记"未执行"，**不算失败**。
- `index.js`：注释里的真实目录名改为占位（`D:\<中文目录>\...`），技术信息不变。
- `REMOVED-MODELS.md`：实测例子里的具体资料名改为泛化表述。

### 6. 本轮验证

- **语法**：`node --check` 通过（`index.js` / `skill.js` / `test.mjs` / `repro.mjs`）
- **模块加载**：`apply` / `inject` / `name` 三个 Cordis 契约导出齐备
- **测试套件**：`node test.mjs` → **714 项检查通过，2 项失败**。2 项失败均为
  `F1矩阵：b/c`，其夹具 `test-out/scan-sample.pdf` 是 2026-09-14 的旧缓存产物，
  与本地引擎的当前预期不符；**本轮改动的 6 个行段均不涉及该段代码**，属既存夹具状态问题。
- **契约复核**：工具名与 HEAD 逐字一致；环境变量与导出函数**只增不减**；
  `dependencies` / `devDependencies` / `peerDependencies` 均为空（零第三方依赖成立）。
- **版本门槛**：扫描全部有版本门槛的语法特性，**实测所需最低 Node ≥ 16**（`engines.node`
  声明的 `>=18` 为保守值）；全部文件**无顶层 await**。

> ⚠ **如实限定**：上述"版本门槛"为**语法特性静态扫描**结论，非多版本实机运行验证
> （本机仅有 Node v26.7.0，无 nvm/fnm/volta）；实测运行环境为 Node 26 / DSH 0.1.7-alpha.2 / Windows x64。

## 2026-09-26 — 第十九轮（R19）：把 R18 的「剩余风险」钉成硬证据或明确结论（零功能改动）

> 事由：收尾清单（A–E）。**本轮不新增功能、不改任何既有判据**。原计划"零代码改动"，但任务 D2 的
> **全仓正则审计**查出 **4 类真实可复现的 `RangeError`**（与 R18 修的那两条同病），因此范围收在
> "修掉这 4 类 + 补回归用例"（**正常输入逐字不变**，每条都有等价性哨兵；只有"输入大到旧写法会崩"
> 时行为不同 —— 旧：宿主级 `RangeError`；新：正常返回）。另外 `zip.js` 有一处 JSDoc 改动
> （`opts.baseOffset` 定性为测试接缝）。逐项证据、量化表与交付报告在 `DEVELOPMENT.md`「第十九轮（R19）」。

### 1. ZIP64「>4 GiB 数据尺寸」：**声明级证据已拿到**，代码路径端到端落成**已知限制**

- **声明格式被主流实现接受（本轮新证据）**：手工拼"声明 ≥4 GiB、实际数据 5 字节"的 ZIP64 包
  （extra 直接调本仓 `zip64ExtraField()`，字段序与 `makeZip` 一致），`.NET ZipArchive`（`Expand-Archive`
  的引擎）把 `Entry.Length` 读成 **4294967396**；`Expand-Archive` 本体对形态①也能解出 1 个文件。
- **本仓读侧比 .NET 更严（有意为之）**：形态①本仓按**声明**拦在 256 MiB 单条目上限（`.NET` 反而
  只按 `compSize` 读 5 字节就返回成功）；形态②（压缩 + 未压缩尺寸都 ≥4 GiB）本仓直接判"声明压缩长度
  超过文件长度"，`.NET` 读到流时报"本地文件头已损坏"。
- **真实 >4 GiB 合法条目（本轮新证据）**：流式 deflate（1028.9:1）造出 4,294,971,392 字节的真实条目
  （12.9 s → 4.17 MB 压缩数据），`.NET` 正确解出前 1 MiB（全为 `A`，`nonA=0`）⇒ 声明与真实数据都被接受。
- **代码路径端到端不可行 → 已知限制（不是"风险"）**：`makeZip` 走 `unc64` 需要 ≥4 GiB **真实可遍历**输入
  （`crc32` 逐字节）+ `concatBytes` 等量输出，**峰值 ≈8.0 GB**；本机 15.9 GB 总内存、探针运行时空闲
  2.9–3.6 GB。触发条件（4 GiB 级单条目，现实里到不了）、影响面（写侧按规范、读侧先被 256 MiB 上限拦下，
  **不存在静默废包路径**）、重开条件（空闲 ≥10 GB 的机器跑 `work/r19-probe/probe-zip64-makezip-4gib.mjs`）
  全部写进文档。
- `opts.baseOffset` 定性为**测试接缝、非稳定 API（`@internal`）**：本仓零生产调用点。

### 2. `Expand-Archive` 真解包：**跑到底**（不再只跑到超时）

- `status=0`、**65,536/65,536 个文件**、**2,180,154 ms（36.4 分钟）⇒ 33.27 ms/文件**；
  同一包的 `.NET ZipFile::ExtractToDirectory` 只要 75,046 ms（**1.145 ms/文件**）、逐条 `ExtractToFile`
  2.272 ms/文件、`OpenRead + Count` 1,590 ms。
- ⇒ "套件默认只用 `.NET ZipArchive` 读取"是**有量化依据的取舍**（`Expand-Archive` 慢 **29×**，
  开销全在 PowerShell cmdlet 的每文件参数/路径处理上），真解包列为 opt-in（**发版前**手跑一次）。

### 3. 瞬时锁 / 跨卷 / 网络盘

- **瞬时锁（新用例 `R19-原子写：`）**：子进程用 `FileShare.None` 拿住目标文件，**主进程开始写之后 70 ms
  释放**（反向握手保证首次 `rename` 落在锁持有期内），`writeFileAtomic` 在退避窗口内**重试后成功** ——
  终态是新内容、原文件未被破坏、无临时件残留，**实测 203 ms**（走完 60 + 120 两次退避后成功）。
  R18 只证明了"锁一直不放 → 退避窗口被走到"，这一半是新的。
- **跨卷 `rename` → `EXDEV`（新用例）**：把"临时件必须与目标同目录"从注释变成实测断言
  （网络盘 = 另一个卷，临时件放 `tmpdir()` 必然 EXDEV）。
- **网络盘 / OneDrive：本机无法验证（明确结论）** —— 本机无映射网络驱动器（`net use` 空），
  `<home>/OneDrive` 目录存在但只有一个 `desktop.ini`（未登录 / 未同步），**不能代表同步盘语义**。
  已留 opt-in 用例 `DSH_OFFICE_TEST_NET_DIR=<网络盘目录>`：有环境的机器上一条命令即可验证，
  **没有拿本地目录冒充**。

### 4. 大输入：阈值量化 + 全仓正则审计 + 端到端守门

- **阈值量化**：用**真实旧实现**（`work/r18-run/backup-before/index.js` 的两个热点）逐档定位到
  **(5 MiB, 6 MiB]**（5 MiB 通过、6 MiB 抛 `RangeError: Maximum call stack size exceeded`）；
  `--stack-size` **100 / 200 / 400 / 984 / 2000 / 8000 KB** 六档 + `--no-opt` **一次都不移动阈值**
  ⇒ 不是 JS 调用栈深度、也不是 JIT 优化状态，而是 **V8 对单个超长正则匹配的内部限制** ⇒
  **不是可移植常数**（修正 R18 的"与机器栈大小相关"表述，并升级为"只能在实现层消除整串正则"）。
  点号路径旧实现 = O(匹配长度) 次 `Set.add`（8 MiB → 1,626 ms），新实现 75 ms（≈21×）。
- **全仓正则审计**：危险 **14** 处 / 需注意 **27** 处 / 已判定安全 **25** 处（报告
  `work/r19-probe/regex-audit.md`）。本轮**修掉 4 类真实可复现的 `RangeError`**（全部补了回归用例）：
  `model.js::markdownToDocument`（单行 2 MiB `-`）、`index.js::pdfOutputGate`（整篇 8 MiB 单行 `-`）、
  `xml.js::parseXML`（单标签内 2e6 个属性 —— 所有 OOXML / ODF 部件的入口）、
  以及 `odf.js::decodeNumeric`（**可达**）+ `docx.js::decodeEntitiesRaw` / `pptx.js::decodeRaw`
  （后两处当前零调用点）的 `String.fromCodePoint` 缺守卫（`&#` + 400 位数字 ⇒ `Infinity` ⇒ 崩溃）。
  **判据**（本轮实测）：只有 ① `{n,}` 且 n ≥ 4、② 量词循环落在**复合体**上 这两类会抛；
  简单原子上的 `+` / `*` / `{1,3,}` / `\.{3,}` 在 8–64 MiB 都安全。修复前后同一份探针：
  4 THROW → **10/10 ALL OK**。其余（6 处"计数即物化匹配数组"、4 处同量级数组/Set）已登记
  判定与重开条件，**本轮不动**。
- **端到端守门（新用例 `R19-大输入：`）**：`office_read` 读 **8 MiB 单行 `.md`** 与
  "**8 MiB 单行 XML 部件**的 `.docx`"都不抛 `RangeError`、都在耗时上界内。

### 5. 测试与文档

- 新增 **20 项检查**：工作副本 `node test.mjs` → **739 checks / ALL PASS**（基线 719 ⇒ +20）；
  **同步到目标目录后用 `danger-full-access` 再跑一次 → 744 checks / ALL PASS**（+5 = 受限沙箱下被跳过、
  提权后真跑的两类：非 `%TEMP%` 目标目录质量门 4 条 + 原生 Office COM 冒烟 6 条中净增的 5 条）；
  `node repro.mjs` → 🟢 GREEN。跑完再取 SHA-256 清单：**43 个文件逐字不变**（测试没污染目标目录）。
- `test-out/`（**133 个历史文件、98.45 MB**，跨 2025-08-16 → 2026-09-26）确认是历次"在插件目录里直接跑
  `node test.mjs`（没设 `DSH_OFFICE_TEST_OUT`）"留下的产物，与本轮无关 ⇒ **清理**（同步时删除）；
  文档写明"测试产物一律落 `%TEMP%`"。
- `DEVELOPMENT.md` **顶部新增"⚠ 验收硬规则"**：工作副本跑绿 ≠ 验收 —— 必须同步到目标目录并用
  `danger-full-access` **再跑一次**（受限沙箱会**静默跳过**原生 Office COM 冒烟与非 `%TEMP%` 质量门用例；
  R13 / R16 的教训是"套件全绿盖不住 Office 拒开"）。
- `dsh-badge` 技能目录**仍不存在**（R19 复核）⇒ 继续跳过、不自造徽章；文末署名说明按本轮复核更新。
- 工作区归档策略：`rNN-src` / `rNN-probe` / `rNN-run` **保留原名**（不改名，避免破坏既有文档引用）。

### 6. 兼容性影响与未处理项

- **兼容性影响：无（对外契约）**。四个工具的参数、返回值、错误文本**一个字没变**；本轮代码改动
  全在解析层（`model.js` / `index.js` / `xml.js` / `odf.js` / `docx.js` / `pptx.js`），
  且**只在"输入大到旧写法会崩"时行为不同** —— 旧行为是宿主级 `RangeError`（整条调用作废），
  不是可依赖的契约；正常输入逐字不变（每条修法都配了等价性哨兵）。
  （新增的 `test.mjs` 用例全部是内部回归，不改变任何对外行为。）
- **未处理项（明确结论）**：网络盘 / OneDrive 的实测（本机无环境，已留 opt-in 用例）；
  `makeZip` 的 ≥4 GiB **数据尺寸**端到端（内存门槛，已定性为已知限制 + 可复现探针）。
- **明确不做**（维持 R18 的"只记录结论 + 重开条件"）：PDF 合并 / 拆分、OMML 公式写入、`.html` 结构化编辑、
  插图落点扩展、`formula.js` 的公式写回。

## 2026-09-26 — 第十八轮（R18）：真实功能缺陷（PDF 产出质量门落点）+ ZIP64 写出端 + 真实锁回归 + 输出边界 + 文档治理

> 事由：收尾清单（A–G），一次做完。**本轮只有一处行为不兼容点**（见 §7）。
> 最高优先项是一条**真实功能缺陷**：在任何非 `%TEMP%` 目录 create/convert PDF **100% 失败**（§1）。
> 另修掉一条"上一轮号称已修、其实没修好"的栈溢出（§5）。逐轮修复史与交付报告在
> `DEVELOPMENT.md`「第十八轮（R18）」。

### 1. PDF 产出质量门的**渲染暂存件落点**（最高优先，真缺陷）

- **病灶**：`pdfOutputGate()` 把渲染校验用的 PDF 副本写在**目标同目录**，而本机 WinRT 的
  `StorageFile.GetFileFromPathAsync` **只允许 `%TEMP%` 下的文件**（探针实测：`%TEMP%` 下 `png=1/status=0`，
  工作区与 `~/.dsh` 下一律 `png=0/status=1`，错误是"拒绝访问，该项目没有位于应用程序可以访问的位置"）
  ⇒ **在任何工作区 / 用户目录产出 PDF 都报"第 1 页渲染失败…逃生通道"**。R16 只是把测试产物目录挪到
  `%TEMP%` 绕开，代码没修。
- 渲染副本改写到 `tmpdir()`（默认 `%TEMP%\dsh-office-pdfgate\<tag>\check.pdf`，新增
  `DSH_OFFICE_PDF_GATE_DIR` 覆盖根目录），渲染完**立即删除**；校验**通过后**才复用公共两段式
  （目标同目录唯一 `.part` 临时件 + `rename`）发布 —— "未经校验不发布"不变，且**不再为了渲染往目标目录写任何东西**。
- 失败路径不变：清理暂存件与临时件、不产出目标、抛「PDF 产出拒绝｜四要素」；渲染失败且暂存目录不在
  `%TEMP%` 内时，错误文本追加一句位置提示（否则 `DSH_OFFICE_PDF_GATE_DIR` 指错只会看到 `exit=1`）。
- `stats.pdfQuality` 字段名与语义**完全未变**；新增导出 `pdfGateStagingRoot()` 与 `runRenderScript()`。
- 连带修掉两条长期 FAIL：`R13-2 判别实验（WinRT 打开同骨架 AES-128 夹具）` 与 `探针本身有效`。
  **真实根因是两个**：① 用例自己 `spawnSync` + 管道 stdio 在受限沙箱里被拒成 `EPERM`（实测
  `spawnSync powershell.exe EPERM`，`stdio:'ignore'` 即可）；② WinRT 只读 `%TEMP%`。修法是**与实现同源**
  （探针 PDF 落 `pdfGateStagingRoot()` + 用 `runRenderScript()`），**断言一个字没放宽**。
  修后 `plainPng=1 aes128Png=1`，结论仍是"骨架正确 ⇒ AES-256 打不开指向 WinRT 不支持 R5/R6"。

### 2. ZIP 写出端支持真实 ZIP64（`makeZip`）

- 条目数 ≥ `0xffff`、尺寸 / 压缩尺寸 / 本地头偏移 ≥ `0xffffffff` ⇒ 写 **ZIP64 扩展字段（ID 0x0001）+
  EOCD64 + 定位记录**；阈值取 `≥ 0xffff`（`0xffff` 本身就是"见 ZIP64"的哨兵，`.NET ZipArchive` 会据此去找
  ZIP64 记录，留一个只有普通 EOCD 的 65535 条目包会踩歧义）。
- **不需要 ZIP64 的包字节布局与旧版逐字一致**（旧 `zip.js` 快照 vs 新版逐字节对照，4 组条目 SHA-256 全等）；
  `store` 语义与 ODF 的"mimetype 首个且不压缩"未动。
- **条目名 > 65535 字节仍然显式报错**：`nameLen` 在 local header / 中央目录都是 16 位，ZIP64 只扩尺寸与偏移，
  **规范里没有扩展条目名长度的记录**（APPNOTE 4.5.3）；报错文案已改成说清这一点。
- 新增 `opts.baseOffset`（默认 0，把整包当成从该偏移开始的片段），同时是"覆盖 >4 GiB 偏移分支"的测试接缝。
- **CRC 误拒面：保持严格拒绝**（不做"静默自愈重试"）—— 中央目录声明了非 0 CRC 而数据不符意味着流真的不是
  同一份内容，自动关闭校验重试等于把不可信数据当可信；两个逃生口（`crc === 0` 放行、`DSH_OFFICE_ZIP_CRC=0`）
  与点名条目的文案已足够，本轮补齐三态断言。

### 3. 原子写：真实文件系统错误回归（不再只有注入钩子）

- 只读目标（`chmodSync(0o444)`）→ `rename` 报 **EPERM**、抛四要素、**原文件逐字节不变**、不留临时件、
  **没有 unlink 只读目标**；`writeFileAtomicSync`（sidecar 路径）同等断言。
- 同进程持有句柄（`openSync(target,'r+')`）→ 同样 **EPERM**（Windows 语义下 rename 被挡）。
- **跨进程 `FileShare.None` 真锁**（子 PowerShell 持有 + 就绪/释放标志文件）→ EPERM + 四要素 + 原文件不变，
  且**退避重试窗口真的被走到**（实测 395 ms ≥ 60+120+180 ms 上界）。
- 不用 `chmod` 目录（Windows 上无效），不用注入钩子顶替（证明不了真实锁下的安全性）。

### 4. 输出边界与 schema 加固

- `defineToolLite` 的 `output.schema` 增加 `properties: { content: { type: 'string' } }`（**不加 `required`**）：
  将来把 `content` 写成对象会在 host 侧立刻变 `ToolOutputError`。动手前核实全部 `content` 赋值都是字符串，
  并用 `as=json/meta/markdown/text + paths` 五形态断言未误伤。
- `itemViolations`：`null` / `undefined`（含**稀疏数组空洞**，校验循环由 `forEach` 改为下标循环）与
  "实际收到字符串/数字/数组"三类文案分开口径。
- **新收窄**：`find` / `replace` / `text` / `markdown` / `cell` / `name` / `newName` 非字符串 → 带 `operations[N]`
  下标的参数错误（此前会被 `String(op.x ?? '')` **静默字符串化**）。**刻意不收窄** `value`（数字是常态）与
  `regex` / `whole`（truthy 习惯）。
- `.ocr.json`（遗留保护名，本仓无写入点）复核结论：**保留**（删掉只有"清扫器删用户成果"的单向风险），
  结论写进 `isProtectedCacheName` 注释。

### 5. `textQuality` 大输入：修掉上一轮没修好的栈溢出

- R16 把 `s.match(/[A-Za-z]{6,}/g)` 换成 `matchAll`，但**单个巨大匹配时 `RegExpStringIterator.next` 自己就炸**
  （8 MiB 连续 `a` 仍 `RangeError: Maximum call stack size exceeded`；8 MiB 连续 `.` 走旧写法不崩但要 2564 ms）。
- 现在 `textQuality` 里两条整串正则**全部改成手写扫描**（点前导就地识别、字母连排一遍数完），零正则、
  零中间字符串、O(n) 零分配。判据口径用小样本与旧"正则 + `split`"参考实现**逐样本对照一致**。
- 实测：8 MiB 连续字母 128 ms（原为崩溃）、8 MiB 连续句点 86.6 ms（原 2564 ms，约 **30×**）、
  8 MiB 正常中文 38.6 ms、8 MiB 混合 67.6 ms。

### 6. 测试与文档

- 新增 **63 项检查**（656 → 719），工作副本 `node test.mjs` → **719 checks / ALL PASS**（基线 656 / 2 FAILED）；
  **同步回目标目录后再跑一次 → 724 checks / ALL PASS**（多出的 5 项是受限沙箱下被跳过、提权后真跑的两类：
  非 `%TEMP%` 目标目录的 PDF 产出 4 项，以及**真的开成功的原生 Office COM 冒烟**——
  Word 16.0 打开 `含图 .docx`（`inlineShapes=2`）并导出 PDF、PowerPoint 16.0 打开 `.pptx`（`slides=3`），
  负向控制 `notesMaster` 改回共用 `theme1` 仍 `0x80070570`）；`node repro.mjs` 🟢 GREEN。新增：大 stored 图片端到端（2000×2000 → 11.45 MiB PNG，`store` 落地）、
  大输入基准、zip 早停 RSS 自校准断言（0.9 MiB vs 干净子进程真解压 131.5 MiB）。
- **文档治理**：`README.md` 瘦身为"快速入门 / 能力矩阵 / 常见用法 / 指向"（删掉与 SKILL 重复维护的长段）；
  五类"未实现能力的草案"从 SKILL 迁到 `DEVELOPMENT.md`；SKILL 里 8 处已漂移的
  `index.js:xxxx` / `test.mjs:xxxx` 行号**全部符号化**（改为 `grep -n '<符号名>'`）；
  "23 个 PDF 1.1 秒""510 格 4.9%""56,341 字节 / NUL 5135 / C0 13666""渲染倍率实测表"等
  **不可证实的数字**迁到 `DEVELOPMENT.md` 并标注"当轮实测、未自动化"。
  补全 `stats.pdfQuality` 的 `images` / 条件性 `imagesSkipped` / `notice`、批量盘点的 `blocks`/`sheets`/`slides`、
  `suggestedBatches` 的 ">12 页才 15 页/批" 口径、`DSH_OFFICE_MAX_INLINE_CHARS` 的 <1000 回落 120000、
  缓存身份"路径指纹不做短名/符号链接归一化"这条已知限制，以及环境变量总表里缺的
  `DSH_OFFICE_PDF_GATE_DIR` / `_ZIP_MAX_*` / `_ZIP_CRC` / `_ATOMIC_FSYNC`。
- 署名徽章：`~/.dsh/skills` 下没有 `dsh-badge` 技能目录、全盘搜不到规范图片 ⇒ **按约定跳过，不自造徽章**，
  已在 `DEVELOPMENT.md` 文末留一行说明。

### 7. 兼容性影响（本轮唯一的不兼容点）与未处理项

- **不兼容**：`office_edit` 的 `find` / `replace` / `text` / `markdown` / `cell` / `name` / `newName`
  传**非字符串**时由"静默字符串化"变成**参数错误**。原因：运行时与早就声明 `string` 的 schema 对齐。
  迁移方式：传字符串（`find: 2024` → `find: "2024"`）。影响面：只有传数字/对象的调用方可见；
  字符串调用逐字兼容。
- **未处理项 / 剩余风险**（详见 `DEVELOPMENT.md`「剩余风险」）：>4 GiB **数据尺寸**分支未能端到端
  （不可分配 4 GiB 以上缓冲，只由布局单测 + `baseOffset` 接缝间接覆盖）；`Expand-Archive` 真解包 65536 条目
  实测 >10 分钟（900 s 时已解 54,357/65,536，无错误）故列为 opt-in；网络盘 / OneDrive 的 fsync 与 rename
  未验证；杀毒 / 同步盘类的**瞬时锁**未模拟（只用确定性锁）；CRC 严格拒绝策略下"可用但 CRC 不符"的包
  需显式设 `DSH_OFFICE_ZIP_CRC=0`；`textQuality` 栈溢出的精确阈值未定（只在 (1 MiB, 8 MiB] 内确认）。

## 2026-09-26 — 第十七轮（R16）：审计驱动的加固（缓存身份 / 原子落盘 / ZIP 限额 / schema 一致性）

> 事由：外部审计列出五组发现，按优先级全部处置（未处理项见 §6）。
> **本轮唯一的行为不兼容点**：集中缓存目录 / 临时回退目录下的缓存文件名新增 8 位路径指纹后缀，
> 旧文件不再被当作命中 —— 但会被**发现并明确作废**（`stats.ocrCacheStale.reason = "identity-missing"`）。
> 同目录 sidecar（`<file>.ocr.md`）命名**未变**。
> 文档侧：逐轮修复史从 `README.md` / `SKILL.md` 迁到新建的 `DEVELOPMENT.md`（本轮全量记录也在那里）。

### 1. 缓存身份（`.ocr.md` / `.read.md` / 渲染目录）

- 身份 = **规范化源路径指纹**（`sha256(lowercase(realpath))`，零 IO）+ **源内容 SHA-256**
  （优先复用解析时已读进内存的 `buf`；无 buf 时 1 MiB 分块同步哈希 + `路径|size|mtime` memo）。
- manifest 新增**独立一行** `<!-- srcpath: 64hex | srcsha256: 64hex | srcsize: N -->`
  （covered 行逐字不变，旧调用方 `split('\n')[1]` 的断言不受影响）。
- 读取时 `parser` 与身份**逐项核对**，任一不符整份作废；`staleReason` 分六类：
  `parser` / `parser-missing` / `identity-missing` / `path-mismatch` / `content-mismatch` / `unverifiable`。
- 集中目录 / 临时回退：`<name>-<pathKey8>.ocr.md`、`.read.md` 同理、渲染目录 `<safe>-<pathKey8>-<mtime>`；
  旧命名（无指纹）会被**发现并明确作废**（报真实路径 + 原因），绝不静默复用。
- `covered` / `src` / `retry` 的跨批累积语义原样保留（有回归用例守着）。

### 2. 原子落盘（create / edit / convert / sidecar / 渲染切片）

- `saveBuffer` 从 `writeFile`（open 即截断）改为「目标同目录唯一临时件（`wx`/O_EXCL）→ 完整写入
  （fsync best-effort）→ `rename` 发布」；失败时清理临时件并**保留旧目标**，抛「写盘失败｜四要素」。
- Windows 三条实测约束：rename **能**覆盖已存在目标；目标被占用 → EPERM/EBUSY（退避重试 60/120/180ms，
  仍失败保留原目标，**绝不 unlink** —— 那会绕过用户的只读保护，而真正的文件锁下 unlink 同样 EBUSY）；
  跨卷 = EXDEV → 临时件必须与目标同目录（禁用 `tmpdir()`）。
- `pdfOutputGate` 复用同一套两段式（不再自带 stamp 命名）；"未经校验不发布"的语义不变。
- sidecar 用同步版 `writeFileAtomicSync`（保持"同步返回、失败 undefined"契约）；`bandPng` 切片缓存
  同样原子写（半截 PNG 只要 >500 字节就会被当成"已渲染好"复用，视觉模型于是读到坏图）。

### 3. ZIP 资源限制与结构校验（`zip.js`）

- 新增 `ZipView` 有界读取器：中央目录 / local header / 偏移 / 长度 / ZIP64 尺寸全部边界校验，
  越界一律**可读中文错误**（不再 `RangeError: Offset is outside the bounds of the DataView`）。
- 修掉两个真 bug：① 旧第 90 行 `if (localOff === 0xffffffff) { /* localOff = */ … }` 的赋值被注释掉
  （完全合法的 ZIP64 包 100% 读不了）；② ZIP64 分支多写 `&& cdCount === 0`（真实 ZIP64 永不进入）。
  另修：stored 条目静默返回「负载+CD+EOCD」拼接物、nameLen 说谎导致名字变垃圾、重名静默覆盖。
- 解压前按声明尺寸挡两道上限，解压时用 `maxOutputLength` 早停，解压后核对实际长度；CRC32 默认校验
  （`crc !== 0` 才比对）。默认单条目 **256 MiB** / 整归档累计 **512 MiB**
  （`DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` / `DSH_OFFICE_ZIP_MAX_TOTAL_BYTES` 可覆盖，`0` = 关闭该上限），
  `DSH_OFFICE_ZIP_CRC=0` 关闭 CRC。阈值依据：实测最大合法条目 17.31 MiB（20000 行 xlsx）、
  最大 stored 图片条目 11.45 MiB；**刻意不做压缩比判据**（图片条目压缩比恒为 1.00×，按压缩比拒会误杀）。
- `makeZip` 写侧补守卫（条目数 > 65535 / 名字 > 65535 字节 / 数据·偏移 > 4 GiB 一律报错），
  不再 `Math.min(...)` 静默丢条目、或写 0 哨兵产出废包。
- 连带修复：`textQuality()` 的 `s.match(/[A-Za-z]{6,}/g)` 在单个 8 MiB 连续字母串上会栈溢出（合法输入）
  → 改 `matchAll` 惰性迭代。
- 验证：17 个攻击用例全部转为可读错误或**正常读取**；
  **333/333 份真实 Office 文档逐条目 SHA-1 一致、0 误拒、0 条目被上限命中**。

### 4. schema 与运行时行为一致

- `grid`：schema 从 `type:'string'` 改为 `oneOf:[{type:'string'},{type:'number'}]`（host 的 JSON Schema
  子集支持 `oneOf`，且明确拒绝 `type:[...]` 数组），validator 同步放宽 → **数字形式从"完全不可用"变成可用**。
- 同源修复：`jsonSchemaOf` 不再把 `integer` 降级成 `number`
  （`office_read` 的 `offset` / `limit` / `pageFrom` / `pageTo` 的 schema 与 validator 现在一致）。
- `office_edit.operations.items`：补 `op` 枚举（16 个实现真正支持的取值）+ `required:['op']` +
  常用字段描述；运行时新增 `itemViolations`，非法元素报 **`operations[N]` 下标定位**
  （旧行为会把 `{}`、缺 `newName` 的 `rename_sheet` 放到执行期，甚至写出 `Sheet1 → undefined`）。
  容器相关的必需性仍由执行期精确报错。
- `width` 描述改为实现真相（省略 / 0 / 负数 / NaN 时按 `px × 72/96`，上限 451.3pt；**没有"默认 450"**）；
  `as="json"` 描述写明 `content` 是 **JSON 文本字符串**（返回结构未改）。

### 5. 测试与验证

见 `DEVELOPMENT.md`「第十七轮（R16）· 交付报告」：实际命令、通过 / 失败 / 跳过数与原因。
（攻击用例 17/17 与真实语料 333/333 两项的实测见 §3。）

### 6. 未处理项（明确不做 / 留待后续）

- `zip.js` 写侧仍**不生成 ZIP64**（>4 GiB / >65535 条目改为显式报错，不再静默丢数据）。
- PDF 合并·拆分、OMML 公式写入、`.html` 结构化编辑、插图落点扩展：维持第五/六/七轮结论（不做）。
- 本机限制（非代码问题）：沙箱内 WinRT PDF 栅格化**只能读 `%TEMP%` 下的文件**，测试产物必须落在
  `%TEMP%`，否则渲染级用例整片失败。

## 2026-09-26 — 待裁定记录：GPU OCR 引擎栈是否接入 DSH 侧（复核结论，**零代码改动**）

> 事由：主人问「DSH 桌面端现在也有 Python 环境了，dsh-office 是否能接入 GPU OCR」。
> 本条目**只记录复核证据与结论**：**不实施、不改任何引擎代码**——`rapidocr.js` / `index.js` /
> `SKILL.md` / `test.mjs` / `vendor/` **本轮全部零改动**，对侧 WB 全程只读。
> **裁定（本次，主人）**：① **暂不动，保持 CPU-only**；② 「不引入 Python」红线**先记入待裁定、不实施**。
> ⚠ **CHANGELOG 自身行号基准**：本条目插在全文**最顶部**，下文引用的 CHANGELOG 行号一律为
> **插入前**基准（插入后整体后移），并同时给**文字锚点**——偏移后按锚点定位。

### 1. 结论

**技术上已具备条件，但尚未接入，且不宜现在自动接入。** 挡路的是 §3 三道闸门（引擎栈未移植 / 沙箱 /
宿主私有资产），**不是 GPU 能力本身**。

### 2. 前提变化：DSH 桌面端确已提供 Python 运行时

| 事实 | 实测值 |
| --- | --- |
| 解释器 | `~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/python.exe`，**3.12.14** |
| 包管理 | pip **26.2.1** 在位 |
| 预装包（`runtime.json` 的 `pythonPackages`，**固定载荷**） | numpy 2.3.5 / pandas 3.0.1 / python-docx 1.2.0 / python-pptx 1.0.2 / openpyxl 3.1.5 / Pillow 12.3.0 / lxml 6.1.3 / XlsxWriter 3.2.9 + 传递依赖 |
| **不含** | onnxruntime / opencv / rapidocr / torch / paddle —— **给的是解释器，不是推理栈** |

⇒ 当年「**不引入 Python**」红线的**前提已消失**（该红线的三处锚点：「不引入 Python / LibreOffice / pdftotext / npx」，插入前行号 L1706 / L1819 / L2062），
但**并没有随附 GPU 推理能力**——栈仍得自己装，而沙箱下装不了（§3 闸门二）。

### 3. 三道闸门（本次实测）

**闸门一 · 引擎栈未移植。** 本侧 `rapidocr.js` 497 行 / 21673 字节、单一 C++ 引擎
（`LOCAL_ENGINE='rapidocr'` `:37`、`LOCAL_MAX_IMAGES=20` `:40`、`DEFAULT_MODELS` `:42-47`）；
对侧 WB `rapidocr.js` 55950 字节，含 `engine:` 三态、`vendor/pyocr/service.py`（482 行，常驻 ONNX Runtime +
**DmlExecutionProvider**）、候选解释器链、`DmlExecutionProvider` **门禁**（"没报错 ≠ 跑在 GPU 上"）、
`ocrEngineRole` 记账与回退。本文件已三处明载该栈为对侧独有（**插入前**行号 + 锚点）：
L789「**WB 侧独有（本侧红线不做）：GPU OCR 引擎栈**」、L1143-1144「反向分叉（WB 有、DSH 没有，
本批**不移植**）：… GPU OCR 链路」、L1287「GPU OCR 链路（Python/DirectML）… **是否回移由主人决定**」。

**闸门二 · 沙箱（本侧特有，对侧设计不能照抄）。**

- **禁 named pipes**：Node `spawnSync(python, …, {encoding:'utf8'})` ⇒ `error.code='EPERM'`（`status=null`）；
  同命令改 `stdio:'inherit'` ⇒ `status=0`。⇒ GPU 常驻服务的 IPC **必须走文件轮询**；对侧 `--mode auto`
  首选**管道**，在 DSH 侧会**每次白试一次 EPERM**。本侧现有 C++ 路径早已出于同一原因改用 `fs.openSync`
  文件 fd（`rapidocr.js:24-25`「no named pipes are needed (sandboxed hosts block pipes; fds to files
  always work)」）⇒ **移植时 file 轮询应设为本侧默认**、pipe 降为可选。
- **禁装包**：`pip install --target … onnxruntime-directml`（清华镜像）⇒ `[Errno 13] Permission denied`（临时目录）；
  把 `TEMP`/`TMP` **重定向进工作区**后**仍失败**（同一错误 + "Failed to remove contents"）。⇒ 当前沙箱档位
  **无法就地安装**，GPU 栈只能**预置**。
- **插件目录不可写**：`~/.dsh/plugins/dsh-office` 下写探针文件 ⇒ `UnauthorizedAccessException`
  （workspace-write 只覆盖会话工作区）。⇒ 预置 venv 需**升华权限**或沙箱外操作——与第十一轮 §1 表
  「全量回归必须升华权限在沙箱外跑」**同源**。

**闸门三 · 不该依赖宿主 Python。** 宿主运行时在 `~/.dsh/dsh-runtimes/dsh-primary-runtime/…`
（`runtime.json`：`desktopVersion 0.1.7-rc.2` + `payloadDigest`，即**随 DSH 版本/载荷变**），属**宿主私有资产**；
AGENTS.md 明令**不得把宿主私有资产塞进 office 核心路径**（"DSH 一升级就碎，而 web / headless 上没有"）。
另有**版本绑定**：onnxruntime wheel 是 **cp312** 绑定，宿主升 3.13 即碎。
⇒ 正解是**自足**：仿对侧候选链置于插件内 / 用户级（`DSH_OFFICE_OCR_PYTHON` → `vendor/pyrt/` →
`vendor/pyocr/.venv/` → `~/.dsh/ocr/pyocr/.venv/`），全链拿不到则**回退 C++ 并显式记账**
（`ocrEngineRole=fallback` + 正文脚注），**绝不静默**。

### 4. 若日后实施：两条路线与已知代价

| 路线 | 取舍 |
| --- | --- |
| **A. 插件自带 venv**（`vendor/pyocr/.venv/`，约数百 MB） | **倾向**：自足、不惧 DSH 升级、不触红线纪律；代价是仓库体积 + 首次预置需升华权限 |
| **B. 复用宿主 `python.exe` + 插件依赖走 `PYTHONPATH`** | 省一套 Python；但**绑定 cp312**（宿主升级即碎）且**违纪**（闸门三），不推荐 |

其余照抄对侧即得的**代价**（引自 WB `SKILL.md:349-375` 对侧实测，**未在本侧复现**）：
GPU ≈ **3.6×** 快于 C++；DirectML **首帧 1.5–2 s** shader 编译 ⇒ **必须常驻**；
GPU 与 C++ 字符产出差 **−1% ~ +6%**（要逐字复现历史结果时不可开）；换倍率重试时显存峰值可达 **2.7 GB**、
**可能比 C++ 更慢**；空闲显存 **<1.5 GB** 不宜用。

### 5. 本次实测清单（可复现）

1. 对侧 `~/.workbuddy/binaries/python/envs/ocr-dml/Scripts/python.exe`（Python **3.13.14**）：
   `providers=['DmlExecutionProvider','CPUExecutionProvider']`，onnxruntime **1.24.4** / cv2 **5.0.0** /
   numpy 2.5.3 / **rapidocr_onnxruntime 1.4.4** / pyclipper 1.4.0 / shapely 2.1.2 ⇒ **DirectML 在本机真实可用**。
2. 对侧 `vendor/pyocr/service.py --selftest`，**在当前 DSH 沙箱内**返回
   `{"ok":true,"providers":["DmlExecutionProvider","CPUExecutionProvider"],…}` ⇒ 引擎执行本身在沙箱下**没问题**，
   **堵点是 IPC 与安装，不是执行**。
3. 本机硬件：**GTX 1660 Ti 6 GB** / 驱动 **610.88**（DX12，满足 DirectML）——来自 `nvidia-smi`。
4. 本侧 `vendor/RapidOCR-json_v0.2.0/cmd.txt`（54 行）参数表**无任何 GPU / EP 开关**（仅
   `--models/--det/--cls/--rec/--keys/--padding/--maxSideLen/--boxScoreThresh/--boxThresh/--unClipRatio/--doAngle/--mostAngle`）
   ⇒ 现役引擎确为 **CPU-only 构建**；**不换引擎的前提下给它加 GPU 开关是做不到的**。

### 6. 本次未做

- **未改任何代码**：`rapidocr.js` / `index.js` / `pdf.js` / `formula.js` / `SKILL.md` / `test.mjs` / `vendor/`
  零改动；未新增环境变量；未新增 `engine:` 语义。
- **未装任何包**（沙箱不允许，见闸门二）；工作区留下的探针目录 `work\gpu-ocr-probe\`（venv 半成品 + `pip-full.log`）
  **与插件无关**，可随时删。
- **未开新轮次施工**：本条目**不是**第十二轮，只是待裁定留档；将来实施时**另开条目**，并按本文件
  「修订轨迹不删」体例在此条目旁注"后来怎么变的"。

## 2026-09-25 — 第十一轮（DSH 侧）：RapidOCR 失实描述修订 / R9-3 注释同步 / 三条环境规则持久化 / CRLF 例外注记 / CHANGELOG 行号纪律 / 环境变量审计

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。上一批基线 **588 checks**
> （第十轮收口态）。本批 = 第十轮 §7 八项挂账的收口：**6 项本轮完成**（§7 #1/#2/#3/#4/#5/#6，其中 #5 为"确认在案"）、
> **#7 `findEngine` 与 #8 沙箱归因此前已封案**（本轮零改动）。
> **不动 OCR 引擎栈**（本侧 = CPU RapidOCR + 视觉桥）、**不动** `findEngine` / `PDF_PARSER_VERSION` /
> sidecar 协议 / 续读协议 / 返回值消毒管线；**不开新功能**（五项草案不碰）；
> **对侧 WB 全程只读，一个字节未动**（181 个文件，最新 mtime 14:00:55，早于本会话起点 15:50）。
> **行号基准**：本条目对其它文件的行号一律为**改后**状态；对 CHANGELOG 自身的行号一律显式标注"改前/改后基准"。

### 1. 基线与权限口径（改前实测，非沿用最近基线）

| # | 命令 | 结果 |
| --- | --- | --- |
| 1 | `node test.mjs`（**workspace-write 沙箱内**，首次） | **❌ 2 FAILED (586 checks)**，exit 1 |
| 2 | `node test.mjs`（**升华权限、沙箱外**，第十轮 §7 #8 既裁正解） | **✅ ALL PASS (588 checks)**，exit 0 |
| 3 | `node repro.mjs`（升华权限、沙箱外） | **🟢 GREEN**，exit 0（43 页 / 31200 字符，坏码点 0 / 非 NFC 0） |

**沙箱内 586 / 2 FAIL 的归因（本轮实证，非代码回归）**：`test.mjs:38-50` 的 `ensureOut()` 自述"从插件目录直跑在
workspace-write 沙箱下会 EPERM"，本次产物确实被迫回退 `%TEMP%\dsh-*\dsh-office-test-out`（沙箱外口径为插件目录 `test-out`）；
两条 FAIL 均为 `R13-2 判别实验`（`plainPng=0 aes128Png=0`），其 `pdf-render.ps1` 子进程输出走**管道捕获**，
受限沙箱禁 named pipes ⇒ 探针拿不到 stdout ⇒ 判 FAIL；同一份代码的历史 `test-run.log` 尾行是 `✅ ALL PASS (588 checks)`
且 `plainPng=1 aes128Png=1`。⇒ 与 §7 #8 裁定一致：**全量回归必须升华权限在沙箱外跑**。
证据：`stage\dsh-office-r11\baseline-test.txt`（沙箱内原样输出）、`baseline-test-sandboxoff.txt`、`baseline-repro-sandboxoff.txt`。

### 2. `SKILL.md:510-516` 的 RapidOCR 失实描述 → 修（先核实代码再改写）

**代码事实**（`rapidocr.js` 497 行；主代理逐行实读 + 子代理只读复核，证据 `stage\dsh-office-r11\audit-rapidocr.md`）：

- 候选链 `engineCandidates()`（`rapidocr.js:75-85`）：`DSH_OFFICE_RAPIDOCR_DIR`（`:78`）→ `<插件目录>/vendor` 与
  `vendor/RapidOCR-json_v0.2.0`（`:79`）→ `~/.dsh/ocr` 与 `~/.dsh/ocr/RapidOCR-json_v0.2.0`（`:80`）→ `PATH` 每个条目两子形态（`:81-83`）；
- 两处 `continue`：`:110`（目录/exe 不存在）、`:113`（有 exe 但 `models/` 缺 `DEFAULT_MODELS` 任一项）——**都只跳过该候选、继续探测下一个**；
  只有整链走完 `found` 仍为 `null`（`:117-118`）或 `DSH_OFFICE_OCR_DISABLED`（`:105`）才 `findEngine()` 返回 `null`；
- 可用性检查**硬编码** `DEFAULT_MODELS`（`:42-47`）、模型子目录固定 `'models'`（`:111`），`:112` 该行**无任何 `process.env`**
  ⇒ `DSH_OFFICE_OCR_MODELS` / `_DET` / `_CLS` / `_REC` / `_KEYS`（运行侧，`engineArgs()` `:134`/`:138-141`）**不参与可用性检查**；
- 退回视觉桥后 `via='no-local'`（`index.js:3098`）⇒ 标注"视觉模型识别 · 本地引擎不可用"（`index.js:3143`）。

**病灶**：原句"引擎缺 `RapidOCR-json.exe` 或四个默认模型时…只是安静地改用视觉桥"主语未限定，读起来是"任一候选缺件 ⇒ 改用视觉桥"，
与上述 `continue` 语义矛盾（`SKILL.md:191`/`:193` 本身已是正确口径）。

**改法**：`SKILL.md` 改前 L510 一行 → 改后 **L510-516**（7 行：候选链 +「只跳过该候选」+「全链不合格才退回」+ 五变量不参与可用性检查），
逐点对齐 `README.md:129-132` 基准。**规模**：`SKILL.md` 59448 B / 521 行 → **60220 B / 527 行**（+772 B / +6 行）。
候选文本与核实行号：`stage\dsh-office-r11\r11-patch-candidates.md` P1。

### 3. `test.mjs:806-809` 的过时说明 → 修（与 `:1834-1838` 互指）

R9-3 之后"降级 notice/content 会覆盖这条脚注——那是产品现状"已不是当前行为。修订只落在 `//` 注释内：

- 区块 A（**`test.mjs:806-809`**，4 行 → 4 行）："那是产品现状"→"那是旧行为"+「R9-3 起从 stats 重建同一批脚注并追加」，
  保留夹具结论"用有文本层夹具验文案通道"，末尾加 `见 test.mjs:1834`；
- 区块 B（**`test.mjs:1834` / `:1836`**）：只换两个 token（`⚠ 实测产品现状` → `⚠ 历史判断`；`…单独覆盖。` → `…单独覆盖（见 test.mjs:806）。`），补 A⇄B 互指。

**断言数未增未减（改后实测）**：区块 A `ok(` = **3**、区块 B `ok(` = **6**、全文 `ok(` = **585**（与改前同值；
`checks` 由 `results.length`（`test.mjs:33` push）决定，注释改动不产生/不删 `ok()`）。
**规模**：`test.mjs` 249200 B / 3719 行 → **249372 B / 3719 行**（+172 B / **0 行**）。`node --check test.mjs` exit 0。

### 4. 三条环境规则持久化 → 方案 A（`~/.dsh/AGENTS.md`）

**备份名冲突与负责人裁决**：阶段 0 要求备份"不存在或与源文件完全一致"，实测 `AGENTS.md.bak-20260925` **已存在且不一致**
（1236 B / SHA256 `5B178146…574132`，第九轮**改前**版：第 3 行多 ` / EPUB`、第 8 行旧 OCR 口径；`byteIdentical=False`）。
经负责人裁决：**保留旧备份不覆盖，本轮改前快照另存新名** `AGENTS.md.bak-20260925-r11`。

**两处任务给定措辞与盘上事实不符，经负责人裁决后按实测改写**：

- 规则 3 原拟"`repro-run.log`（UTF-16LE）"——实测 `repro-run.log` = **1055 B / UTF-8 无 BOM / 纯 LF**（CR=0、NUL=0），
  `test-run.log` 同（63254 B / 651 LF / 0 CR）；UTF-16LE 只在 **PowerShell 重定向捕获**时出现 ⇒ 规则 3 改写为
  "全 CRLF 行尾例外 + 运行产物当前为 UTF-8/纯 LF（重定向会变 UTF-16LE）+ 跨侧按字节复制"；
- `REMOVED-MODELS.md` 注记原拟"（32 行 CRLF）"——实测"行数=32"与"CRLF 序列总数=32"两读同时成立，头部加一行必成 33 行/33 CRLF
  ⇒ 按裁决改为**不含会过时数字**的准确注记（见 §5）。

**diff（改前 → 改后，位于「办公文档」节末、`## 思维链语言` 之前）**：

```diff
 完整规范：`skill("dsh-office")`。
+
+- 行尾敏感文件禁用裸 `git apply`（本机 `core.autocrlf=true` 可能把整份 CRLF 文件改写）；用 `git -c core.autocrlf=false apply`，或复制已校验的纯 LF 产物。
+- workspace 内禁止放置名为 `AGENTS.md` 的候选副本（会被宿主当指令注入），候选文件统一命名 `AGENTS.md.candidate`。
+- `REMOVED-MODELS.md` 是全 CRLF 行尾例外；运行产物 `repro-run.log` / `test-run.log` 为 UTF-8 无 BOM、纯 LF（用 PowerShell 重定向捕获会变 UTF-16LE）—— 跨侧搬运一律按字节复制。
 
 ## 思维链语言
```

**规模**：1328 B / 20 行 → **1882 B / 24 行**（+554 B / +4 行），UTF-8 无 BOM、**纯 LF**（CR=0）、末字节 `0x0A`；三条规则位于 **`AGENTS.md:11-13`**。
生效面：注入每会话系统提示 ⇒ **新会话**生效（本轮写入后宿主已即时重载，实测三条规则可见）。
备份：`<home>/.dsh\AGENTS.md.bak-20260925-r11`（1328 B，SHA256 `84D0FDD0…F475413`，与原文件逐字节相同）。

### 5. `REMOVED-MODELS.md` 头部注记 → 加（CRLF 例外）

**改前逐字节**：1743 B / 32 行 / **CR=32 / LF=32 / 裸 LF=0** / 无 BOM / 末字节 `0x0A`。
**新增注记**（**改后 `REMOVED-MODELS.md:1`**，全文首行）：`本文件为 CRLF 例外（全 CRLF 行尾），跨侧搬运按字节复制`
**改后逐字节**：1820 B / 33 行 / **CR=33 / LF=33 / 裸 LF=0** / 无 BOM / 末字节 `0x0A`；注记之后的 1743 B 与改前**逐字节相同**（脚本实测 `true`）。
除新增注记外零改动、零行尾变更。**冲突判定记录**：任务给定"CRLF×32"在本文件上"行数=32"与"CRLF 序列总数=32"两读同时成立；
按"每行统一 CRLF 属形态描述、非总数硬约束"解释并**经负责人裁决**采用不含数字的注记，改后行数按实际（33/33）验收。
证据：`stage\dsh-office-r11\audit-agents-encoding.md` §3.1/§3.4，`r11-patch-candidates.md` P5。

### 6. CHANGELOG 纪律行 + 行号引用影响（本轮唯一获准的顶部局部修改）

新增 1 行，插在顶部纪律声明块末行「修订轨迹不删」之后（**本行即本轮唯一获准的顶部局部修改**，旧行只位移、不改字节）：

> 行号引用一律用全文件行号，或显式注明是条目内相对行号。

**影响（只读审计 `stage\dsh-office-r11\audit-changelog.md` §3）**：

- 全文自指引用 31 处 = **显式全文件 2 处**（改前 L96 的 `CHANGELOG.md` L7-L265、改前 L148 的 `全文件对应 L385-L393`）
  + **条目内相对 29 处**；`CHANGELOG.md:NNN` / `#LNNN` / `第 N 行` 三形态 **0 命中**；
- 顶部新增（1 纪律行 + 本条目 197 行）⇒ 旧 L6 起全体位移 **+199**；那两处显式自指改后位于
  **L295** 与 **L347**（改前基准 L96 / L148）；
- **29 处条目内相对行号不受影响**（引用与被引目标同步位移）；`§3a L1…L5` 8 处是条目标签、约 41 行指向他文件，均不受影响；
- ⚠ 那 2 处显式自指**在本次插入前就已失准**（均为第九轮时期快照）：改前 L96 的 `L7-L265` 今日语义自相矛盾
  （改前 L7 是第十轮标题），第九轮条目真值为改前 L374–L632；改前 L148 的 `L385-L393` 对应"第八轮条目 L119-127"今日真值 = **L751–L759**。
  按「修订轨迹不删」**原文保留、不回改**，本轮以"新条目注记 + 显式基准"取代；
- **不改任何旧轮记录、不删任何修订轨迹**（append-only：旧内容零字节改动）。

### 7. 逐项局部检查（改后实测命令与结果）

| # | 命令/动作 | 实际结果 |
| --- | --- | --- |
| 1 | `node --check test.mjs` / `node --check repro.mjs` | 均 **exit 0** |
| 2 | 断言计数（锚点切片正则，两种口径） | A=**3** / B=**6** / 全文=**585** ⇒ **未增未减** |
| 3 | 改后字节核验（`ReadAllBytes` + 正则统计） | `SKILL.md` 60220 B / CR=0 / 无 BOM / `0x0A`；`test.mjs` 249372 B / CR=0 / 无 BOM / `0x0A`；`AGENTS.md` 1882 B / CR=0 / 无 BOM / `0x0A`；`REMOVED-MODELS.md` 1820 B / CR=33 / LF=33 / 无 BOM / `0x0A` |
| 4 | `grep SKILL.md` 残留检查（`安静地改用视觉桥` / `安静回退视觉桥` / `候选目录不合格` / `立即退回`） | **仅 1 命中**：`SKILL.md:193` 的"（第九轮更正：原「安静回退视觉桥」在本机不成立）"——属**修订轨迹**（引用并否定旧表述），非当前表述 |
| 5 | 对侧 WB 只读核验 | `<home>/.workbuddy\skills\dsh-office\` **零写入**（181 文件，最新 mtime 14:00:55 < 本会话起点 15:50；全程仅只读 `read`/`grep`，无写命令、无备份、无产物） |

### 8. 环境变量文档对账（第 6 项）→ 只读审计完成，**差值 0**，文档一字不改

扫描方法复用对侧 `.r15-artifacts\env-scan.mjs`（**仅复用方法**，未在对侧执行任何命令），本侧化脚本
`stage\dsh-office-r11\env-scan-dsh.mjs`，输出 `stage\dsh-office-r11\env-scan-out.txt`：

- 三类间接引用全覆盖：字符串形式（`envOr('DSH_OFFICE_TEST_CSV',…)`）、`pyEnv()` 类封装的**对象裸键**（`OMP_NUM_THREADS` 等）、`os.environ` / `$env:` 读取形式；
- 范围：本侧 22 个活代码文件（`*.js`/`*.mjs`/`*.ps1`/`*.py`，跳过 `test-out`）+ 4 个文档；
- **结果**：插件自有变量 **34 个**，`SKILL.md` 完全未收录 **0 个**；只在文档出现、代码里没有的 **0 个**；
  宿主/系统变量 6 个（`APPDATA`/`HOME`/`PATH`/`SystemRoot`/`USERPROFILE`/`WINDIR`）不计入插件差值；
- 中间伪影已修正并留档：首版用 `` `VAR` `` 精确包裹判定，而本侧文档一律写作 `` `VAR=<值>` ``（如 `` `DSH_OFFICE_TEST_OUT=<目录>` ``、`` `DSH_OFFICE_RENDER_QUIET=1` ``），
  误报 11 个缺失；改判"反引号内**包含**变量名"后归零。
⇒ **差值 0 ⇒ 文档一字不改**；本项以"审计完成、差值 0"登记（非"文档已补齐"）。

### 9. U1 确认在案（本轮零改动）

第十轮 §7 #5（改前全文件 L369）原文仍为「**已登记待考**：不复述旧字节数，**不判 K=2**」；口径正文在改前 L149-151。
**本轮对该项零改动、未复述任何旧字节数、未判 K=2**（确认位置：改后 L568）。

### 10. 未做项与理由

| 项 | 状态 | 理由 |
| --- | --- | --- |
| §7 #6 WB 侧 `index.js:3689-3694` 注释过时 | **本轮不动** | 跨侧留档：WB 由自己的轮次更新；**不得判为本侧偏差** |
| §7 #7 `findEngine` 探测与运行不同源 | **本轮不动** | 第九轮 §4 已封案（含「下轮不再翻案」），除非出现真实"换模型"需求 |
| §7 #8 沙箱归因 | **本轮不动** | 第九轮 §6 已给正解（升华权限在沙箱外跑），本轮以之取得 588 基线 |
| §7 #5 U1 字节数记账口径 | **零改动，仅确认在案** | 见 §9；不复述旧字节数、不判 K=2 |
| 五项草案（PDF 合并拆分 / HTML 结构化编辑 / 插图落点 / `recalc_formula` / OMML） | **不碰** | 本轮不开新功能 |

### 11. 挂账清零对账（第十轮 §7 八项）

| §7 # | 条目 | 本轮处置 |
| --- | --- | --- |
| 1 | `test.mjs:806-809` O1 措辞未同步 | **完成**（§3，与 `:1834-1838` 互指，断言数不变） |
| 2 | 三条环境规则无持久落点 | **完成**（§4，方案 A + 负责人授权的新备份名） |
| 3 | `REMOVED-MODELS.md` 自身零注记 | **完成**（§5，改后 33 行 / 33 CRLF） |
| 4 | X3 引用歧义（相对 vs 全文件） | **完成**（§6，顶部纪律行 + 本条目显式基准） |
| 5 | U1 字节数记账口径 | **确认在案**（§9，零改动） |
| 6 | WB 侧注释过时 | **登记为跨侧留档**（§10，本轮不动，不判偏差） |
| 7 | `findEngine` 探测与运行不同源 | **此前已封案**（第九轮 §4，本轮零改动） |
| 8 | 沙箱归因 | **此前已封案**（第九轮 §6，本轮以之取基线） |

### 12. 子代理盘上产物复核（主代理已读盘，不据返回消息下结论）

四路只读子代理（仅写 `stage\dsh-office-r11\`，未写任一 live 目录、未跑全量测试、未碰对侧）：

| 产物 | 复核结论 |
| --- | --- |
| `stage\dsh-office-r11\audit-rapidocr.md` | 候选链 / `continue` / 可用性检查逐条行号证据；主代理独立复读 `rapidocr.js:75-119`、`:134-143` 核对一致 |
| `stage\dsh-office-r11\audit-test-comments.md` | 两区块原文、唯一性、断言计数（A=3/B=6/585）均与主代理改后实测一致 |
| `stage\dsh-office-r11\audit-agents-encoding.md` | AGENTS / 备份 / REMOVED-MODELS 全部字节数与哈希经主代理 `ReadAllBytes` 复核一致；**`repro-run.log` 非 UTF-16LE 的反证成立** |
| `stage\dsh-office-r11\audit-changelog.md` | 顶部纪律区 / §7 八项 / 行号引用影响分析，经主代理 `read` 抽验一致 |

### 13. 收口测试（改后，升华权限、沙箱外、独占）

| # | 命令 | 结果 |
| --- | --- | --- |
| 1 | `node test.mjs` | **✅ ALL PASS (588 checks)**，exit 0，**0 FAIL**（产物 `${DSH_HOME}/plugins/dsh-office\test-out`） |
| 2 | `node repro.mjs` | **🟢 GREEN**，exit 0 |
| 3 | `node --check test.mjs` | exit 0 |
证据：`stage\dsh-office-r11\final-test.txt`、`final-repro.txt`。

### 14. 生效方式

- `test.mjs`：**新进程运行即生效**（`node test.mjs` / `node repro.mjs` 直接读盘）；
- `SKILL.md`：由 `skill.js` 每次读盘 ⇒ 修改后**即时生效**，无需重启；
- `~/.dsh/AGENTS.md`：注入每会话系统提示 ⇒ **新会话**生效（本轮写入后宿主已即时重载）；
- `REMOVED-MODELS.md` / `CHANGELOG.md`：纯记录文件，无运行期加载；
- **本轮未修改任何 ESM 常驻模块**（`index.js` / `rapidocr.js` / `pdf.js` / `formula.js` 等**零改动**）
  ⇒ **不存在"重启 profile 才生效"的文件，也未发生重新部署**。

> **第十一轮收口：登记挂账清零。** 第十轮 §7 八项处置完毕（6 项本轮完成、#7 findEngine 与
> #8 沙箱归因此前已封案、#5 确认在案）。此后改为**触发式维护**，仅三种情况开轮：
> ①某侧改功能/修 bug 需回移对侧；②`node test.mjs` 出现 FAIL；③出现真实功能需求（五项草案之一被实际需要）。
> 纯文档打磨、重复对账、两侧 OCR 对比**不构成开轮理由**。

## 2026-09-25 — 第十轮（DSH 侧）：PDF 分支补回 `__fullBody` / `formula.js` 头注释对齐 29 函数 / 第九轮留档清点

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。上一批基线 **586 checks**
> （第九轮收口态，实证 `test-run.log` 尾行 `✅ ALL PASS (586 checks)` / `repro-run.log` 🟢 GREEN）。
> 本批 **2 项主修 + 1 项只记账的防守核对**（第 3 项**不改代码**，只把漏登记补进本条目）。
> **不动 OCR 引擎栈**（本侧 = CPU RapidOCR + 视觉桥）、**不动** `PDF_PARSER_VERSION` /
> `DSH_OFFICE_CACHE_DIR` / sidecar 协议 / 续读协议不变式（`offset + content.length === nextOffset`）；
> **`findEngine` 不修**（第九轮 §4 已裁定，含 L133「下轮不再翻案」）；**本轮不动 `SKILL.md` / `README.md`**
> （无与之直接矛盾的改动）；**禁止整目录覆盖 WB 侧**（`<home>/.workbuddy\skills\dsh-office\`）——
> WB 全程**只读**，一个字节未动。

### 1. `index.js` PDF 分支补回 `out.__fullBody`（第九轮 §9 唯一真实偏差 K = 1 的收口）

**病灶**（第九轮 §9 L255-264 已登记，本轮实测复核）：`readPdf` 的返回对象（`index.js:3619-3627`）
**没有** `out.__fullBody`；而**非 PDF 分支**（`index.js:3900-3901`）有
`nonPdfOut.__fullBody = content`。消费端 `finishRead`（`index.js:836-839`）一直在读它：

```js
const fullBody = typeof out.__fullBody === 'string' ? out.__fullBody : ''
if (fullBody) delete out.__fullBody          // 绝不出现在返回值里
…
const sidecarBody = fullBody.length > out.content.length ? fullBody : out.content   // L866
… stats.sidecarChars = sidecarBody.length                                           // L871
```

⇒ **PDF 正文被内联护栏截断 + 质量门判乱码**同时命中时，`sidecarBody` 退化成
`out.content`（= `capped.content` = `body.slice(0, limit)`），sidecar 落的是**截断版**，
而 notice（`index.js:881-882`）承诺的是「已把**整篇正文**转存到 sidecar」—— **成了假话**。

**补丁**（1 行代码 + 4 行注释，插在 `index.js` L3637 的 `  }` 与 L3638 的 `  return out` 之间）：

```js
  // 同 non-PDF 分支（下方 `nonPdfOut.__fullBody = content`）：截断过时必须把**截断前的完整正文**
  // 交给 finishRead，sidecar 才真落得下"整篇"（见该函数头注释）。第十轮补回 —— 第三轮误删了本行，
  // 而 finishRead 一直在消费 `out.__fullBody`：PDF 正文被内联护栏截断且质量门判乱码时，sidecar
  // 会落**截断版**，notice 承诺的"整篇正文"就成了假话。
  if (capped.truncated) out.__fullBody = content
```

- **消费端契约一字未动**（`finishRead` L836-839 照旧读 + `delete`）；
- 补丁后**返回值里绝不出现 `__fullBody`**（断言 R10-2 显式锁住，见 §4）；
- WB 侧同语义实现在 **`index.js:3695`**（注释 3689-3694），本侧注释按 **DSH 口径**重写，
  **不含** WB 的「⚠ 反向待确认项 / 建议 DSH 侧下一轮补回」表述（该表述在本侧补回后已过时，见 §5）。

**规模**：`index.js` 223638 B / 4173 行 → **224142 B / 4178 行**（+504 B / +5 行，逐 hunk 实测仅此一处）。

### 2. `formula.js` 头注释对齐 29 函数（**纯注释，代码体一行未动**）

**病灶**：`formula.js` L6 自述「支持（子集，**写死在这份清单里，README/SKILL 同步**）」，
但 L7 的函数清单只列到 `… ABS INT MOD`（18 个名字），**漏 `COUNTIFS / SUMIF / AVERAGEIF` 与 11 个文本函数**；
实现侧的 `SUPPORTED_FUNCTIONS`（L15-20）与文档侧 `README.md:208-209` / `SKILL.md:291-292` **都是全的**
⇒ 头注释的两句自述**同时失实**（三方不一致）。此前全 CHANGELOG `头注释` **0 命中**（第九轮范围外，本轮新建）。

**补丁**：只替换 L6-12 头注释段为 **29 函数版**，并按 DSH 侧实际补三行说明：

- 函数 29 个（`SUM AVERAGE MIN MAX COUNT COUNTA COUNTIF COUNTIFS SUMIF AVERAGEIF` /
  `IF VLOOKUP ROUND ROUNDUP ROUNDDOWN ABS INT MOD` /
  `LEFT RIGHT MID LEN TRIM UPPER LOWER CONCAT CONCATENATE TEXT VALUE`）；
- **不认识的函数/引用一律显式 `unsupported`，绝不猜值**：整列引用 `A:A` → 标签带原文
  `A:A（整列引用）`；外部工作簿 `[Book1]…` → 词法把 `[` 判坏字符，标签 `不支持的引用/字符 "["`；
  `_xlfn.` 前缀按函数名比对，不在 29 个里 → 标签 `_XLFN.SUM`；
- **大表护栏 `DSH_OFFICE_RECALC_MAX_CELLS`**（opt-in，见 `recalcMaxCells`）：未设 / `0` / 非法值 = 无上限；
  整本单元格总数超限 → **整本原样返回**（`evaluated: 0`、`formulaCells: 0`）+ `report.skipped`；
  **绝不静默截断、不半算、不动任何公式与缓存值**。

> **措辞据实测而定**（`stage\w10\s2-probe-unsupported.mjs`，只 import live 只读）：
> `SUPPORTED_FUNCTIONS.length = 29`；`SUM(A:A)`/`[Book1]Sheet1!A1`/`_xlfn.SUM(1,2)` 三例全部
> `unsupported`，`unsupportedTokens = ["A:A（整列引用）","不支持的引用/字符 \"[\"","_XLFN.SUM"]`，
> 对照组 `SUM(1,2)` 正常 `evaluated`；护栏：`recalcMaxCells(unset/0/abc/100) = 0 0 0 100`，
> 超限时 `evaluated = 0 / formulaCells = 0 / skipped = "表格单元格总数 25 超过 DSH_OFFICE_RECALC_MAX_CELLS=10…"`，
> 且 `model.sheets[0].rows[0][0]` 原公式与缓存值**逐字未动**。

**「代码体一行未动」的字节级证明**（`stage\w10\s2-verify.mjs`，剥离头注释块后逐行比对）：

| 检查项 | 结果 |
| --- | --- |
| `L1-5` 逐字节相同 | ✅ true（sha `b1d92fde…`） |
| `L13..EOF` 逐字节相同 | ✅ true（sha `2ce5b2b8…`，**707 行 = 707 行**） |
| 剥离头注释后 codeBody 逐行一致 | ✅ true（两侧 sha256 **同为 `376f28dd19f62dfe687b495d1bfcda01e82503bf89d535245944a4794ad174f1`**，706 行） |
| `const ERROR_KINDS .. EOF` 逐字节相同 | ✅ true（**28652 B = 28652 B**，同 sha） |
| 全文件 LCS diff | −1 行（L7）/ +10 行，**新增与删除行全部是注释或空行** |
| 全文件 sha256 | live `5f49c43b…`（29589 B）→ 产物 `c2262128…`（**30520 B / 728 行，+931 B / +9 行**） |

**前提校验（改动前必做）**：diff 两侧 `formula.js` 证明**差异只在头注释块、代码体逐行一致** ⇒ 成立。
（WB 侧 L20-25 已是 29 函数版本，但 WB L17 写「本侧**无**护栏」是 WB 第十二轮口径；
本侧**有**护栏，措辞按 DSH 侧实际重写，**未照抄**。）

### 3. 第九轮「三类留档」清点（**只记账，不改代码**）

按第九轮 L46 的三类口径逐条清点第九轮条目（`CHANGELOG.md` L7-L265），共 **28 行**
（A = 已改矛盾处 **13**｜B = 确认无需改 **1**（EPUB，唯一具名）｜C = 发现但本批不动 **11**｜— = 口径/回归 **3 组**）。

**C 类 11 条的「去向强度」分层**（这是本轮清点的核心结论）：

| 去向强度 | 条数 | 条目 |
| --- | --- | --- |
| **去向齐备**（已登记待办 / 不动理由完整） | **6** | §4 `findEngine`（L133 已封案）、§4b 附带发现（并入 spec）、§6 沙箱归因（L179 已给正解）、§6 诊断跑留档、§7 L192-194 落盘纪律、**§9 `__fullBody`（唯一「指定第十轮修」→ 本轮已修，见 §1）** |
| **有规则表述、但无持久落点**（弱） | **3** | §7 L195-199 `core.autocrlf` 坑、§7 L200-201 `AGENTS.md` stage 注入坑、§8 L246-249 `REMOVED-MODELS.md` CRLF 例外 |
| **部分落实**（两处同款只改一处） | **1** | §8 L251-253 附带观察 O1 → `test.mjs:806-809` 未同步 |
| **完全无登记** | **1** | 本轮新增事实：WB 侧 `index.js:3689-3694` 注释已过时（见 §5） |

**落点核查**（证明「无持久落点」不是措辞问题）：`~/.dsh/AGENTS.md`（20 行，仅 3 节）
**无** CRLF / git apply / 候选副本 / 留档规则；`graph_memory_recall("CRLF autocrlf AGENTS 候选 留档")` → **total=0**；
`.dsh-graph` grep → **No matches**（HANDOFF 长期记忆 =「（无）」）；第九轮条目 `grep 第十轮` 仅 L255/L264（均为 `__fullBody`）。
另：第九轮**草稿** `_r9-changelog-draft.md` 止于 L145，`autocrlf|CRLF|候选副本|REMOVED|附带观察|第十轮`
**六词 0 命中** ⇒ §7 两坑 / §8 O1+CRLF 例外 / §9 是**收口期才写进 live**，从未进过草稿。

#### 3a. 漏登记补记（本轮**只补进本条目**，不修代码、不改其它文件）

**L1｜行尾敏感文件禁用裸 `git apply`**（第九轮 §7 L195-199 原文）：
本机 `core.autocrlf=true`，裸 `git apply` 会把整份文件改写成 CRLF。
**实证（第十轮逐字节重测）**：`stage\SUSPECT-work-index-before-reapply.js` = **227811 B / CRLF×4173 / 0 裸 LF**，
与正确产物 `work-index.js` = **223638 B** 恰差 **4173 = 行数** ——「行数相同、字节差恰等于行数」由此坐实。
正解 = `git -c core.autocrlf=false apply -p1`，**或直接复制已校验的纯 LF 产物**（第九/十轮均采用后者；
第十轮 three files 全部走复制并逐个 `Get-FileHash` 确认）。
> **引用注**：第九轮 L195 的「221890 → 227811」是**补丁前 live → CRLF 错产物**（差 5921），
> **不是**纯行尾改写；纯行尾改写的对照应是 **223638 → 227811（差 4173）**。二者勿混用。

**L2｜`AGENTS.md` 的候选副本不得以原名留在 workspace**（第九轮 §7 L200-201 原文）：
`stage\AGENTS.md` 会被宿主当作 **workspace 指令注入会话**（第九轮已实测发生）。
正解现役形态已在位：全 workspace **无任何名为 `AGENTS.md` 的文件**，只有 `stage\AGENTS.md.candidate`（1328 B）。

**L3｜`REMOVED-MODELS.md` 是 CRLF 例外，跨侧搬运必须按字节复制**（第九轮 §8 L246-249 原文）：
第十轮逐字节实测 DSH live = **1743 B / LF 32 / CRLF 32**（无 BOM、末字节 `0x0A`）；WB = LF 44 行。
`repro-run.log` 同类例外（仅 live，UTF-16LE + BOM，loneCR×24）。**这两件禁止用文本读写 / 行尾归一化工具搬运。**
（该文件**自身 32 行内零注记** —— 待办：日后在其头部补一行声明。本轮不动。）

**L4｜O1 只改一处**（第九轮 §8 L251-253）：第九轮记 O1「已就地补写修订注记」，
实证只在 `test.mjs` **乱码夹具段**（现 L1834-1838，已写「第八轮记录；第九轮 R9-3 已修」+ 指向紧邻的 `R9-3:` 两条断言）落实；
**同款措辞第二处 `test.mjs:806-809` 未同步**，仍写「降级 notice/content **会覆盖**这条脚注（记账只剩 stats）——**那是产品现状**」。
R9-3 之后该判断已假（降级路径会从 `stats` 重建同一批脚注并追加进 `notice`/`content`）。
**改法**：只补修订注记、**不改夹具选择**（该段「用有文本层夹具验文案通道」的**结论**在 R9-3 后仍正确）。**本轮不动，登记待办。**

**L5｜WB 侧注释已过时**：见 §5 事实更正 ③。

#### 3b. 清点查出的次级缺陷（登记，本轮只处理 X2）

- **X1（登记错位）**：第九轮 L46 声明「三类留档」，但 §1c 正文（L38-47）**只具名 (1) 类 9 条 + (2) 类 1 条**，
  (3) 类实条目全散在 §4/§4b/§6/§7/§8/§9。⇒ 本轮**已按 (3) 类逐条具名并各带去向**（上表即补齐形态）。
- **X2（`formula.js` 头注释陈旧）**：**本轮已修**（见 §2），不再挂待办。
- **X3（引用歧义）**：第九轮 L75 写「第八轮条目 L119-127」，实为**条目内相对行号**，
  全文件对应 **L385-L393**。⇒ 以后的引用统一用**全文件行号**，或显式注明「（该轮条目内行号）」。
- **U1（记账口径待考）**：第九轮 §8 表（L220）记 `test.mjs 245474 B / 3673 行`，
  而第十轮批写前实测 live = **245725 B / 3675 行**（+251 B / +2 行）。**不排除是纯记账时点差异**
  （对账时点 vs 收口批写），**不宜据此判「真实偏差 K=2」**；本轮**不复述该字节数**，只登记观察。

### 4. 回归与证据（第十轮）

| # | 项 | 结果 |
| --- | --- | --- |
| 1 | `node test.mjs`（升华权限、沙箱外、独占） | **✅ ALL PASS (588 checks)**，**exit code 0**，**0 FAIL**；尾行 `✅ ALL PASS  (588 checks)  产物：${DSH_HOME}/plugins/dsh-office\test-out`（新断言现场见 §4b） |
| 2 | `node repro.mjs` | **🟢 GREEN**，exit=0（\"返回值可无损序列化，harness 一定收\"；43 页 / 31200 字符，坏码点 0 / 非 NFC 0） |
| 3 | **反向自证**（批写**前**，未修复 live） | R10-1 **FAIL** / R10-2 **FAIL** ⇒ 非同义反复（原始输出见 §4c） |
| 4 | 工作副本回写 | 逐文件 `Get-FileHash` **same=True**（见 §4d） |
| 5 | 静态计数 | `ok(` 调用点 **+2**（583 → 585 静态；运行期 **586 → 588**）；test.mjs 内**无**硬编码 checks 常量（`'586'` 两侧均 `false`） |

#### 4a. 新增断言 R10-1 / R10-2（`test.mjs`，插在原 big-junk 断言块 L2238 之后）

**为什么必须新建**：WB/DSH 两侧**同款既有断言** `sidecar提示：内联截断时 sidecar 仍拿到全文`
（WB `test.mjs:1977-1992` / DSH `test.mjs:2223-2237`）用的夹具是 **`big-junk.md`（非 PDF）**，
走的是**已有 `__fullBody` 的非 PDF 分支** ⇒ 它**不可能**锁住 PDF 缺口，**照抄 = 同义反复**。
新断言同时满足三条硬条件：**(a) 走 PDF 分支**、**(b) `capped.truncated === true`**、**(c) `stats.fallback === 'sidecar'`**。

**夹具与参数**（`stage\w10\` 实测定稿）：35 页 PDF（`buildGarbledFixture` 同构：heading +
英文句 + `'\uE0A1\uE0A2\uE0A3'.repeat(30)`，私用区码点是质量门判乱码的触发物）+ **末页独有标记 `R10TAILMARK`**；
`limit = 600`（PDF 分支下界 100，`index.js:3574`）；**`ocr: 'never'`** ⇒ 只读文本层，
**零 OCR 引擎、零 PDF 渲染**（纯 JS 造夹具 + 纯 JS 文本层解析）。

#### 4b. 降级路径三数实测对比（核心证据）

| 状态（夹具 35 页 / `limit:600` / `ocr:'never'`） | `stats.sidecarChars` | 截断后 `content.length` | sidecar 文件字符数 | 正文真实长度 | R10-1 | R10-2 |
| --- | --- | --- | --- | --- | --- | --- |
| **未修复**（`stage\w10\v1b-unfixed\` 完整副本） | **600**（= limit） | **973** | **772** | 8547 | **FAIL** | **FAIL** |
| **修复后**（`stage\w10\v1b-fixed\` 完整副本） | **8545** | **969** | **8713** | 8545 | **PASS** | **PASS** |

- 未修复态：`sidecarChars === limit`（`capWithOffset` 截断时 `content.length` 恰为 `limit`），
  sidecar 文件 **772 < content 973**，且**不含**末页标记 / 末页句（页数只数到 **3/35**）⇒ 落的是**截断版**。
- 修复后：`sidecarChars = 8545 > 600`，sidecar 文件 **8713 > content 969**，含末页标记与末页句（页数 **35/35**）。
- **可复现**：本表与 §4c 的两态都是 `stage\w10\` 下的**独立完整副本**（`v1b-unfixed\` = 用 `baseline\index.js`
  覆盖的完整副本；`v1b-fixed\` = 修复版完整副本）⇒ `node tc-bidir.mjs` **可随时重跑复现**。
  > **取证修正（V1a 观察项 A）**：首次自证跑在**当时的 live**（未修复）上，两态数字为未修复 `983 / 782 / 8552`；
  > 批写后 live 已是修复版，原 `LIVE` 常量会让「未修复」一侧失真 ⇒ 已把未修复基线**固化为 `v1b-unfixed\` 副本**
  > 并重跑（现表即重跑结果）。另：live `index.js` 的 mtime 因 `Copy-Item` 保留源时间戳（11:52:52）而
  > **不等于批写时刻**（真实落地在 11:58:52 创建 `.bak-r10` 之后）—— **mtime 不能用来推断批写顺序**，以 SHA256 为准。
- 「正文真实长度」两态差 2 字符（8547 vs 8545）：来自降级 notice 里**内嵌的文件名长度差**
  （`tc-r10-unfixed.pdf` 18 vs `tc-r10-fixed.pdf` 16），**非正文差异**
  （`test.mjs` 里两侧夹具文件名同为 `r10-pdf-sidecar.pdf`，不存在该偏差）。
  首跑（`…-unfixed-live.pdf` 23 vs `…-fixed.pdf` 16）差 7 字符，**同一机制**。
  **三者数字自洽：整篇 > 截断。**
- 附带确认（未修复态）：`'__fullBody' in out = false`、`JSON.stringify(out).includes('__fullBody') = false`
  —— 该子判据在**未修复态即为 true**，是**防泄漏/防退化**判据，**不是**区分项（区分项是 `sidecarChars` 与末页标记）。

#### 4c. 反向自证原始输出（**可重跑复现**：两态均为独立完整副本）

> 首跑（**批写前**，未修复 live）原件留档 `stage\w10\out\tc-bidir.batch-time.txt`；
> 下面是**固化副本后**的重跑输出（`node tc-bidir.mjs`，exit=0）—— 任何人可复现同一结论。

```
=== [unfixed] .../stage/w10/v1b-unfixed/index.js ===
  夹具页数=35 limit=600 ocr=never
  truncated=true fallback=sidecar sidecarChars=600
  content.length=973 sidecarFileChars=772
  R10-1 子判据: truncated=true fallback=true sidecarChars>limit=false sidecar文件>content=false  ⇒ FAIL
  R10-2 子判据: 尾标记=false 末页句=false 页数=3/35(false) 无__fullBody键=true 无__fullBody串=true  ⇒ FAIL
=== [fixed] .../stage/w10/v1b-fixed/index.js ===
  truncated=true fallback=sidecar sidecarChars=8545
  content.length=969 sidecarFileChars=8713
  R10-1 子判据: truncated=true fallback=true sidecarChars>limit=true sidecar文件>content=true  ⇒ PASS
  R10-2 子判据: 尾标记=true 末页句=true 页数=35/35(true) 无__fullBody键=true 无__fullBody串=true  ⇒ PASS
BIDIRECTIONAL_OK=true  (未修复必 FAIL、修复后必 PASS)
```

**独立复核（V1b，第三方子代理**自建副本 + 自写探针，只读 live）：在 `stage\w10\v1b-unfixed\`
（完整副本，`index.js` 用 `baseline\index.js` 覆盖 ⇒ sha `7C2E4FA6…`）与 `stage\w10\v1b-fixed\`
（完整副本 ⇒ `EA68DD0A…`）上跑同一探针（35 页 PDF / `limit:600` / `ocr:'never'`），独立得出：

| 判据 | `v1b-unfixed` | `v1b-fixed` |
| --- | --- | --- |
| `1-trigger`：`truncated===true && stats.fallback==='sidecar'`（**证明确实走 PDF 分支**） | PASS | PASS |
| `23-sidecar-numbers`：`stats.sidecarChars` | **600**（= limit） | **8568** |
| sidecar `hasTailMark` / `distinctFixturePages` | **false / 3** | **true / 35** |
| `r.content.length` / 未截断对照 `sidecar.fileChars` | 975 / 774 | 975 / 8742 |
| `'__fullBody' in r` / `JSON.stringify(r).includes('__fullBody')` | false / false | false / false |
| `4b-noisy-txt`（**非 PDF** 分支）无泄漏 | PASS | PASS |
| `5-clean`（干净文件）无泄漏 | PASS | PASS |
| 干净文件返回值 `v1b-clean-v1b-{unfixed,fixed}.json` | **`identical = true`（逐字节相同）** | 同 |

⇒ **独立复现「未修复 FAIL、修复后 PASS」**，且确认：夹具**确实走 PDF 分支**（`stats.quality.reasons` =
「私用区码点占 45.2%…」）、`__fullBody` **三态（PDF / 非 PDF / 干净）均无泄漏**、**干净文件零影响**。
（V1b 子代理在**返回环节**中断，但盘上产物 `out\v1b-*.json` / `out\v1b-*-stdout.txt` 完整，
由总控逐项验收 —— 与 §6 的教训一致。）

**独立复核（V1a，第三方子代理，纯静态）**：判据 ①–⑦ **全部 PASS**（产物 `stage\w10\v1a-static.md`，27678 B）：

| # | 判据 | 结论 | 关键证据 |
| --- | --- | --- | --- |
| ① | `index.js` 恰好 +5 行、位置在 `return out` 之前 | **PASS** | `git diff --numstat` = `5 0`，唯一 hunk `@@ -3630,16 +3630,21 @@`；新增 5 行字节 121+137+121+76+49 = **504** = 224142−223638 |
| ② | 新增注释无过时表述 | **PASS** | 「反向待确认项」「建议 DSH 侧下一轮补回」等 4 个关键词在三份改动文件**全 0 命中** |
| ③ | `formula.js` 代码体零改动 | **PASS** | 剥离后两侧 `const ERROR_KINDS..EOF` = **28652 B / sha `376f28dd…`** 完全相同；`SUPPORTED_FUNCTIONS.length = 29` |
| ④ | `__fullBody` 无外泄路径 | **PASS** | 全文恰 **7 处**（2 赋值 + 1 delete + 1 读取 + 3 注释）；无 `...out` / `Object.assign(out` / `JSON.stringify(out`；4 个 `return` 全在 `delete` 之后 |
| ⑤ | 新断言**确实走 PDF 分支**（非同义反复） | **PASS** | 夹具 `writePdf` 产出 `.pdf`（魔数 `%PDF-1.4`）；对照块用 `.md`；**`baseline/index.js` 的 `__fullBody` 只有 4 处且 PDF 分支无赋值** ⇒ 既有 big-junk 断言批写前就 PASS，原理上锁不住 |
| ⑥ | 三份改动文件全 LF / 无 BOM / 末字节 `0x0A` / 语法 | **PASS** | CRLF=0、BOM=no、`node --check` exit=0 |
| ⑦ | `.bak-r10` = 批写前原件 | **PASS** | 三对 SHA256 **IDENTICAL** |

> **V1a 观察项（已处置）**
> **B｜`stage\w10\fixed\` 不是 live 镜像**（其 `formula.js` 仍是旧版 29589 B、且**无** `test.mjs`）——
> 已在 §4d 显式声明「勿以 `fixed\` 为已修复全集」（本轮三份产物分别取用，未受影响）。
> **C｜`live\index.js` 与 `fixed\index.js` SHA256 相同**（`EA68DD0A…`）⇒ 二者**同源**，
> **不能**当作「两次独立落地」的交叉验证；独立验证由 V1b 的**自建副本**（`v1b-unfixed` / `v1b-fixed`）承担。
> **方法学提醒（新增留档）**：PowerShell `Get-Content` 在 **UTF-8 无 BOM** 文件上默认按 **CP936** 解码，
> 会把中文行**合并**从而造成**行号偏移**（实测把 L3638-L3642 报成 3352-3354）。
> ⇒ **取行号一律用 `[System.IO.File]::ReadAllLines(path, UTF8)` 或 node 的 `split('\n')`，
> 不得采信 `Get-Content` 默认编码下的行号。**（本轮总控生成 `tc-test.mjs` 时曾撞上该假象，已由 node 统计澄清。）

#### 4d. 落盘账（live ← 工作副本）

| 文件 | live 批写前（= `.bak-r10`） | live 批写后（= 工作副本） | 字节账 | 行尾 / BOM |
| --- | --- | --- | --- | --- |
| `index.js` | 223638 B / 4173 行 / sha `7C2E4FA6…` | **224142 B / 4178 行 / sha `EA68DD0A…`** | +504 B / +5 行 | LF 4178 / CRLF 0 / 无 BOM / 末字节 `0x0A` |
| `formula.js` | 29589 B / 719 行 / sha `5F49C43B…` | **30520 B / 728 行 / sha `C2262128…`** | +931 B / +9 行 | LF 728 / CRLF 0 / 无 BOM / 末字节 `0x0A` |
| `test.mjs` | 245725 B / 3675 行 / sha `23A80200…` | **249200 B / 3719 行 / sha `32E59E17…`** | +3475 B / +44 行 | LF 3719 / CRLF 0 / 无 BOM / 末字节 `0x0A` |

三侧（live / 工作副本 / WB）**改动文件全 LF + 无 BOM + 末字节 `0x0A`** ⇒ 批写无行尾风险；
唯二例外（`REMOVED-MODELS.md` CRLF、`repro-run.log` UTF-16LE）**均在本轮写集之外**（见 §3a L3）。
`node --check` 对三份产物均 **exit=0**；live 批写为**一次升华权限、逐文件**复制（**禁整目录覆盖**），
`index.js.bak-r10` / `formula.js.bak-r10` / `test.mjs.bak-r10` 三份备份**同目录落地**
（V2 独立复核：三份备份的 SHA256 与 `stage\w10\baseline\` 三份**分别相等** ⇒ 备份确为批写前原件）。

**V2 独立复核（只读 live + 写工作副本）**：工作副本 `work\office-enhance2\plugin\` 三对
`Get-FileHash` 全部 **same=True**（`EA68DD0A…` / `C2262128…` / `32E59E17…`，与 live 逐位一致）；
三侧（live / 工作副本 / WB）**9 个观测点全部 LF + 无 BOM + 末字节 `0x0A`**；
工作副本 33 → 33 文件、仅 3 个被改（无 `*.bak/*.tmp/*.orig/*.rej/*~` 污染）。
**⚠ 下游注意（V2 提醒）**：`stage\w10\fixed\` **不是** live 镜像 —— 它只有 `index.js`（=live）是本轮修复版，
其 `formula.js` 仍是旧版（29589 B = baseline）、且**没有** `test.mjs`；本轮 formula/test 的新件分别是
`stage\w10\s2-formula-header.js` 与 `stage\w10\tc-test.mjs`。**勿以 `fixed\` 为「已修复全集」**。

#### 4e. 宿主内工具重载实验 —— **新发现：`plugin_manager` 的 disable/enable 不能替代 profile 重启**

`~/.dsh/cordis.patch.yml` 的注释早已写明：「a bare specifier cannot carry `?v=N`, so after editing
plugin source a **profile restart** is needed instead of a cache-buster bump」。第十轮**实测坐实**这条，
并补上一条**负面结论**（此前无记录）：

| 步骤 | 操作 | 观测 |
| --- | --- | --- |
| 1 | 批写后、**重载前**，用**宿主内** `office_read` 读 35 页乱码夹具（`stage\w10\out\s1-garbled-35p.pdf`，`limit:600` / `ocr:"never"`） | sidecar = **772 字符**、无第 35 页、页数 **3** ⇒ 落的是**截断版**（旧代码） |
| 2 | `plugin_manager set_plugin` 对 `include:tool-office`（`dsh-office`）**disable** | `{"changed":true,"application":"overridden","warnings":[]}` |
| 3 | 再 **enable** | `{"changed":true,"application":"applied","warnings":[]}` |
| 4 | **重载后**同一探针再跑 | sidecar **仍 772 字符**、仍无第 35 页 ⇒ **未生效** |

**结论（新增留档）**：`plugin_manager` 的 disable/enable 只重跑 `apply(ctx)`，**不会重新求值 ESM 模块顶层**
（Node 的 ESM 模块缓存按 URL 命中），而 `readPdf` / `finishRead` / `capWithOffset` 都是**模块顶层函数**
⇒ **改 `index.js` 后，宿主内常驻工具只有「重启 profile」才生效；`plugin_manager` 重载不构成等效手段。**
（`SKILL.md` 由 `skill.js` **每次读盘**，不受此限 —— 与第七轮留档一致。）

**生效验证因此走「新进程」路线**（与第九轮 §6 同源）：`node test.mjs` / `node repro.mjs` / 反向自证
**全部是新进程直接 import 新代码**，不受 ESM 缓存影响 —— 这正是本轮 **588 ALL PASS** 的效力来源，
也是「反向自证能在未修复态 FAIL」的前提。

**⑤ profile 重启后复验（负责人于本轮内重启 profile，随即实测）**：

| 探针（35 页乱码夹具 / `limit:600` / `ocr:'never'`，**宿主内** `office_read`） | 重载前（旧代码） | **重启后（新代码）** |
| --- | --- | --- |
| sidecar 字符数 | **772** | **8652** |
| sidecar 含 `Fixture page 35` | false | **true** |
| sidecar 内 `Fixture page N` 页数 | **3** | **35** |

⇒ 与 S1 基线「**不传 limit**（不截断）」测得的整篇长度 **8652 完全一致** —— **宿主内 `office_*` 的缺口同样关闭**。
重启后复跑全量回归：**✅ ALL PASS (588 checks) / exit 0**（日志 `stage\w10\out\r10-test-postrestart.txt`）。
⇒ 至此「**盘上代码 / 新进程回归 / 宿主内实测**」三层一致，**无遗留**。

### 5. 事实更正（本轮实测，**修订轨迹不删**）

**① WB 侧落点行号漂移**：第九轮 §9（L260）记「WB 侧在 `index.js:3670-3676` …… `L3676`」——
**实测 WB 注释在 `index.js:3689-3694`、代码行在 `L3695`**；原记的 3670-3676 是 WB `out` 对象
（`const out = {…}`）的位置。⇒ **以本次实测 `3689-3695` 为准**（WB 文件 4224 行，行号本就与 DSH 不同源）。

**② 同款既有断言锁不住 PDF 缺口（必须走 PDF 分支 + 反向自证）**：
两侧 `sidecar提示：内联截断时 sidecar 仍拿到全文` 用的是 **`big-junk.md`（非 PDF）** 夹具，
走**已有 `__fullBody` 的非 PDF 分支** ⇒ **照抄它等于同义反复**。
第十轮据此新建 R10-1/R10-2 并**在批写前做反向自证**（未修复态必 FAIL，见 §4c）；
`test.mjs:2223-2237` 的原文**一字未改**（修订轨迹保留）。

**③ WB 侧「反向待确认项」注释已过时（本轮只改 DSH 侧，WB 一个字节不动）**：
WB `index.js:3689-3694` 那段注释的三句断言在本侧补回后**全部失效**：
「DSH 第三轮误删了这一行」→ 已补回；「sidecar 会落**截断版**……那成了假话」→ 已不成立；
「建议 DSH 侧下一轮补回」→ 已完成。
**处置：只改写 DSH 侧并在本条目登记该 WB 注释已过时；WB 侧不动**（其注释由 WB 自己的轮次更新）。
⇒ **第十轮及以后不得把「WB 注释未改」判为偏差**，也不得据此再开一次「反向待确认项」。

**④ 消费端行号精确化**：第九轮 §9（L258）记「消费端 `index.js:833-839`」——
`finishRead` 里 `__fullBody` 的**读取与删除在 L836-839**（L833-835 是函数头注释）。
本轮引用一律用 **L836-839**。

### 6. 并发拓扑与落盘方式（留档）

原单设计「W0 **4 路并发**（S1 `__fullBody` patch 侦察 + 反向自证夹具探路 / S2 `formula.js` 头注释对账 /
S3 第九轮三类留档清点 / S4 新断言设计）+ W0.5 串行定稿 + W1 串行批写 + W2 串行独占回归 +
W3 **2 路只读复核**（V1 证据核验 / V2 工作副本回写）+ W4 串行收口」。

实测执行要点：

- **W0 的 4 路是真并行**，互不写同一文件（S1→`stage\w10\s1-*`、S2→`s2-*`、S3 只出表、S4→`s4-*`）；
  子代理**一律只 stage 到 workspace、不写 live**（沙箱实测 `Access to the path … is denied`），
  live 由主管用**一次**升华权限（`danger-full-access`）**逐文件**批写 + 同目录 `.bak-r10` 备份。
- **子代理禁止跑 `node test.mjs` / `node repro.mjs`**（test-out 共享 + 渲染/OCR 竞态 + 全量必须独占）；
  子代理只跑自己任务内的**最小复现脚本**（产物落 `stage\w10\out\`），且**一律 `ocr:'never'`、零渲染**。
- **一次子代理服务抖动**：S1/S2/S3 首轮的**返回环节**中断（无 closing message），但**产物均已落盘**，
  由主管逐件复核（S1：`s1-baseline.txt` 三数 + patch；S2：`s2-verify.mjs` 自证 + `s2-probe-unsupported.mjs` 探针；
  S3：清点表 + 补记草案）后采用；S3 已重派一路补完。⇒ **教训：子代理交付以「盘上产物 + 可复核证据」为准，
  不以「返回消息」为准**（与 §3.10「无证据的结论视为未完成」同源）。
- **禁用裸 `git apply`**（本机 `core.autocrlf=true`）：第十轮全程**未调用 `git` 写盘**，
  三份产物均以**直接复制**落地（见 §3a L1）。
- **`AGENTS.md` 候选副本**：`stage\AGENTS.md.candidate` 保持改名形态，workspace 内无同名文件（见 §3a L2）。

### 7. 发现但本批不动（登记给第十一轮）

| # | 条目 | 类别 | 明确去向 |
| --- | --- | --- | --- |
| 1 | `test.mjs:806-809` 的 O1 同款措辞未同步（§3a **L4**） | 记账/注释 | **已登记**：只补修订注记、不改夹具选择；措辞与 L1834-1838 对齐并互相指针 |
| 2 | 三条环境规则**无持久落点**（§3a **L1/L2/L3**） | 操作纪律 | **已登记**：`~/.dsh/AGENTS.md`（≤3 行短句）或 standing memory（≤200 字）二选一，由负责人拍板；本轮**只补进本条目**（按本轮口径「不修代码」） |
| 3 | `REMOVED-MODELS.md` 自身零注记（§3a L3） | 记账 | **已登记**：日后在其头部补一行「本文件为 CRLF 例外，跨侧搬运按字节复制」 |
| 4 | X3 引用歧义（条目内相对行号 vs 全文件行号） | 文档纪律 | **已登记**：以后统一用全文件行号或显式注明 |
| 5 | U1 `test.mjs` 字节数记账口径（§3b） | 记账口径 | **已登记待考**：不复述旧字节数，**不判 K=2** |
| 6 | WB 侧 `3689-3694` 注释已过时 | 跨侧留档 | **已登记**（§5 ③）：WB 由自己的轮次更新；**不得判为第十轮偏差** |
| 7 | `findEngine` 探测与运行不同源 | 裁定类 | **不动**（第九轮 §4 L99-133 已裁定，含 L133「下轮不再翻案」；除非出现真实「换模型」需求） |
| 8 | 沙箱归因（受限沙箱下全量不可复现） | 环境纪律 | **不动**（第九轮 §6 L173-186 已给正解：**升华权限在沙箱外跑**；本轮 588 即由升华权限跑出） |

## 2026-09-25 — 第九轮（DSH 侧）：README/SKILL 对账 / AGENTS.md 口径修正 / sidecar 降级脚注 / findEngine 裁定 / 环境变量文档更正

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。上一批基线 **584 checks**
> （第八轮收口态，实证 `test-run.log` 尾行 `✅ ALL PASS (584 checks)` / `repro-run.log` 🟢 GREEN）。
> 本批原定只做 4 项、逐项回归；**经负责人授权扩为 5 项**（第 5 项 = W4 实测证伪后的环境变量文档更正，
> 见 §5）。**不动 OCR 引擎栈**（本侧 = CPU RapidOCR + 视觉桥）、**不动** `PDF_PARSER_VERSION`、
> `DSH_OFFICE_CACHE_DIR`、sidecar 协议与续读协议不变式（`offset + content.length === nextOffset`）。
> **禁止整目录覆盖 WB 侧**（`<home>/.workbuddy\skills\dsh-office\`）。

### 1. README.md 与 SKILL.md 对账（纯文档）

**1a. pptx 条目：把第五轮的「尚未解决」改为「✅ 已修」。** README 原 L379-383 仍写着
「`office_create` 产出的 `.pptx` 仍被 PowerPoint 16.0 拒开（`0x80070570`）……需要 pptx 时先转
`.docx` / `.md`，或用 PowerPoint 自己的模板另存」，并给「详见 SKILL.md「已知边界」」。该条**已过时**，
且与 SKILL.md L391-411「✅ 已修（第七轮 notesMaster 独立主题部件）」**直接矛盾**；README 自身 L458
早已写「产出的 `.pptx` 能被 PowerPoint 16.0 打开（并带负向控制）」—— 三处互相打架。改写要点：
- 保留**历史诊断链一句话注记**（zip 层用 `makeZip` 重打包仍可开 / OPC 不变量 / 非必需部件逐个删 /
  叶子部件互换全部排除；与 `presentation.xml` 元素顺序、空白、`sldId`、`notesSz`、占位符 `idx` 全无关）；
- **删除**「先转 `.docx` / `.md`」「用 PowerPoint 模板另存」两条绕行指引；
- 正/负向控制与 SKILL.md 对齐：正向 = PowerPoint COM `Presentations.Open` 真开成功（`slides=3`）；
  负向控制 = 把 notesMaster 主题改回共用 `theme1` **必须仍 FAIL**（实测 `0x80070570`）——「改对才开、改错必不开」；
- 根因一句话：`notesMaster1.xml.rels` 的 theme 与 `slideMaster` 共用，PowerPoint 要求 notesMaster
  有独立主题部件（`theme2.xml`）；补丁 = `pptx.js` 的 `THEME_NOTES` + `ppt/theme/theme2.xml` +
  Content_Types Override + `NOTES_MASTER_RELS` 改指 `theme2`。

**1b. 写死的自检基线改为「看最后一行的 `N checks`」口径。** 点名的三处历史数字
（`ALL PASS / 467 checks`、「既有基线 467 checks 全绿」、「既有 467 条用例全部保持通过」）会被误读成
**现行**基线（现 584）。处置：**保留各轮历史数字（修订轨迹不删）**，每处加括注说明「当轮基线」，
统一采用「准确数量看最后一行的 `N checks`，别再往文档里写死」的写法；同类 2 处（第三/六轮的自检句）
一并加同样括注；「验证」节注明本轮实测 **584 checks**、新增断言集成后 **586**。

**1c. 全文对账**（逐节比对 SKILL.md 的第八轮新增/变更，只修**矛盾句**）：短句判据三阈值
（`DSH_OFFICE_OCR_SHORT_RATIO` 0.6 / `_SHORT_MIN_BOXES` 30 / `_SHORT_SCORE` 0.95，「只在低置信才判硬」）
+ 结构性失败 `retryable=false`（换倍率无效，修掉 README「所有未过门页都换倍率」的原矛盾）；
`ocrPreview` / `ocrPagesCapped` 两种「静默少做」不再静默（脚注 + 可复制续读命令 + stats 字段，
**顺序 capped 先 preview 后**）；新增「过门 ≠ 可直接引用」（数字/正负号密集表格务必核对原 PDF，
典型症状**负号被丢** `-9.8 → 9.8`，视觉复核升级**同样防不住**版面级噪声）；sidecar 示例补 `parser:` 版本戳；
候选链补 `RapidOCR-json_v0.2.0` 两层与 `PATH`；环境变量清单补 3 个 `SHORT_*` + 指向 SKILL.md 总表指针；
`pdfOutputGate` 辅助判据补「仅单页文档生效」（`pageCount === 1`，第八轮 P1 修复当时未落文档）。
对账结论按三类留档：**已改矛盾处 / 确认无需改（附理由）/ 发现但本批不动（附理由）**。
EPUB 表述：README grep = **0**（第八轮删的是 `skill.js` 的虚假元数据，README 本就没有）。

**规模**：README.md 53042 字节 / 466 行 → **58303 字节 / 506 行**；`SKILL.md` 见 §5（58868 → 59448，
**仅 L193 一行**不同）。

### 2. `~/.dsh/AGENTS.md`「办公文档」一节的 OCR 口径与能力面（纯文档）

只动「## 办公文档」这一节，逐行核对后**只改 2 行**（另 18 行逐字节相同，`## 思维链语言` 与
`## 交付署名` 两节零改动，L11-L20 拼接段 583 字节 = 583 字节）：
- **删 EPUB**：标题由「Word / Excel / PPT / PDF / ODF / CSV / RTF / EPUB」→ 去掉 ` / EPUB`
  （本侧 `index.js` / `legacy.js` / `skill.js` grep `EPUB` = **0 处**；第八轮已删 `skill.js` 的 EPUB 虚假元数据）。
- **OCR 口径改本地优先**：原「中文 PDF 正文若出乱码（CID 字体缺 ToUnicode 映射），改用 `ocr="always"`
  并指定 `pages` 走视觉识别」→ 改为「先按 `ocr="always"` 指定 `pages` 让**本地 RapidOCR** 离线重识别
  （默认路径、不耗视觉额度）；可疑或不可用才升级视觉桥」。依据 SKILL.md「自动降级链」：
  质量门判乱码的页**自动**送本地识别，识别不可用才落 sidecar / 升级视觉。
  （原文 72 字符 → 新文 115 字符，仍是**单行**；「不扩篇幅」按**行数不扩**执行。）
- **规模**：1236 字节 / 20 行 → **1328 字节 / 20 行**（−7（标题）+99（L8）= +92），UTF-8 无 BOM、纯 LF、
  末尾带 `0x0A`，均与原文件一致。
- **备份**：批写前 `AGENTS.md.bak-20260925`（同目录）。生效面：注入每会话系统提示 ⇒ 对**新会话**生效。

### 3. 裁定：sidecar 降级路径的 OCR 上限脚注 → **修**

**病灶机理（已定位到行）**：`index.js` `finishRead()` L836-885 的 sidecar 兜底分支，L841 对**已含脚注的**
`out.content` 重跑质量门 → 判乱码后 L883 `out.content = excerpt + '\n\n> ' + out.notice`，而
`excerpt = cutAtBoundary(out.content, 800)`（L868）只取**头部 800 字符**；两条上限脚注位于 content
**尾部**（L3475-3484 push 进 `notes` → L3515 `content += '\n\n> ' + notes.join('；')`）→ **随截断丢失**。
记账仍在 `stats`（L3558-3561）：`ocrPagesCapped = {limit,requested,applied,skipped:number[]}`、
`ocrPreview = {preview,total,skipped:<区间串>}`。第八轮 P1 第一版就是被它咬到，才把断言降级为
「降级不丢记账」（第八轮条目 L119-127 已留档，其「已知缺口」即本项）。

**修法**：抽出**唯一文案源** helper `ocrCapNotes(capped, preview, file)`（skipped 形状归一：数组 → `pageRanges`，
字符串 → 原样；顺序 **capped 先 preview 后**），正常路径的 `notes` 组装（改为
`notes.push(...ocrCapNotes(ocrPagesCapped, ocrPreviewLimited, file))`，文案逐字等价）与降级分支**共用**同一份文案；
降级分支从 `stats` 重建两条脚注，追加在**原有 sidecar 说明之后**、同样的 `；` 分隔；并对
「短正文时 `excerpt` 已含脚注」做 `excerpt.includes(note)` **去重**。**零新造措辞**。

**断言**：新增 2 条（`R9-3:` 前缀，实测均 PASS）——① 降级后 **content 通道**出现上限脚注完整文案
（`fallback=sidecar · content 1416 字符 · 脚注出现 1 次`）；② 降级后 **notice 通道**里
「正文质量门判定为不可读」（原有说明）与上限脚注**共存且顺序正确**
（`notice 655 字符 · sidecar说明@156 < 上限脚注@480`）。反向自证：同一段断言原文跑**未修复**的 baseline
→ 2 条**全 FAIL**（非同义反复）。静态 `ok()` 调用点 +2（573→575，静态计数因循环里的 `ok()` 偏低；
运行期增量就是 **584 → 586**）。

> **对原单的两处修正（主管裁定，理由留档）**：原单派 W3 用「26 页有文本层夹具 + 断言 `fallback==='sidecar'`」
> ——该夹具**有文本层、不触发降级**，断言不可能成立；真正的降级夹具是 `buildGarbledFixture()`
> （`pages="1-30"`，21-30 页保留乱码文本层 → 质量门判不可读 → `fallback='sidecar'`，且命中已铺好的 OCR 缓存）。
> 原单第 ② 条「stats 通道」**已被既有断言完整覆盖**（`test.mjs` L1839-1845 逐字段 + L1846-1850 降级记账），
> 复读会让「+2」虚增；故 ② 改落在**从未被覆盖的 notice 通道共存性**上。

**不变式（全部未动）**：sidecar 协议；`offset + content.length === nextOffset`（`capWithOffset` 一行未改）；
**干净文件逐字不变**（`!q.garbled` 早返回分支零行为差异；正常路径 content/notice/stats 与 baseline 逐字节相同）。

### 4. 裁定：`findEngine` 探测与运行不同源 → **不修**（记理由 + 预留改法）

**病灶**：`rapidocr.js` `findEngine`（L104-119）写死 `join(hit,'models')` + `DEFAULT_MODELS` 四文件，
**不读** `DSH_OFFICE_OCR_MODELS`；而 `engineArgs`（L133-152）读 `_MODELS` 与 `_DET/_CLS/_REC/_KEYS`。

**裁定理由（三条，已按 W4 实测更新）**：
1. **触发场景为 0（可判定事实）**：全 live 配置面（插件目录、`AGENTS.md`、`cordis.patch.yml`、
   web/headless 各级 yml、`skills\`）**零 setter**；进程 `DSH_OFFICE_*` = 0、持久 User+Machine env = 0、
   4 个 PowerShell profile 均不存在；workspace 内命中全是往轮产物/diff/快照。⇒ **从未有人换过模型**。
2. **修它是行为变更、不是纯重构**：见 §4b —— 现状在本机不是「安静回退视觉桥」，而是穿透到 `vendor`
   后用**不存在的模型目录**空跑；修好后同机会**真用本地引擎出词** ⇒ 正文 / 脚注 / stats / 耗时全不同，
   **必须有真实四文件自定义模型目录才能验收**（本侧不得动模型栈）。
3. **收益 0**：不设这些变量时新旧逐字一致，没有任何使用者受益。

**§4b. W4 探针的实测更正（本批最重要的新发现，**替换**原单与本目录第八轮 L106 的支撑理由）**：
第八轮 L106 与 SKILL.md L193 都写「候选目录被跳过 → **安静回退视觉桥**」——**该后果描述在本机不成立**。
三组实测（`DSH_OFFICE_RAPIDOCR_DIR=<夹具>`）：
- ① 默认 → 命中 `fixture\models`，`engineArgs` 给 `--models=models` ✅
- ② `_MODELS=custom` 且 `models\` 仍在 → 探测仍认 `fixture\models`，运行却给 `--models=custom`
  ⇒ **不同源铁证** ✅
- ③ `_MODELS=custom` 且**删掉 `models\`** → ❗ **不返回 null**，而是**继续落到下一个候选**（本机 `vendor`）；
  运行侧仍传 `--models=custom`，而 `<vendor>\custom` **不存在** → 引擎空跑 → 每页 `gateResult(undefined)`
  → 按默认 `2,1.5,1` 逐倍率重渲重跑 → **之后**才升级视觉桥（禁用视觉桥时表现为整单抛错）。
  **字节同源对照副本**（屏蔽 vendor 候选）才返回 null ⇒ 纯语义确为 null，偏差来自**候选集合**
  （`continue` 只跳过那一个目录），不是别的分支。真引擎初始化报错文本**未实测**（本侧不得动模型栈），
  该段由代码结构判定。
- 附带发现：`_DET/_CLS/_REC/_KEYS` **同样**是「探测不看、运行照传」；`grep engine\.models` = **0 命中**
  ⇒ `findEngine` 返回的 `models` 字段**全仓无人消费**，探测与运行之间**零交叉校验**（潜伏的根本原因）。

**处置**：裁定不变（**不修代码**），但理由换为「该组变量在本机**不可达**，修它反而引入一次真实行为变更，
须先拿到真模型目录再做 `fix` 级验收」；**预留改法 spec**（附 W4 证据）：抽出唯一 `modelSpec()`
（dir + 四个 files）供 `findEngine` 与 `engineArgs` 共用，缺失判据改 `join(hit, spec.dir)` 且逐项用
`spec.files`，保持 `--models=` **相对 cwd** 语义不变；断言 ③ 期望改为**命中 `fixture\custom`**
（若要保留「null」断言**必须显式屏蔽 vendor 候选**，否则会命中 vendor 而假失败——这正是 ③ 的教训）；
建议按 `fix` 而非 `patch`，验收必须有真实模型目录。**下轮不再翻案**，除非出现真实「换模型」需求。

### 5. 环境变量文档更正（**负责人授权的第 5 项**，纯文档）

§4b 的证据同时证伪了两处文档：
- `SKILL.md` L193（环境变量总表 `OCR_MODELS/_DET/_CLS/_REC/_KEYS` 行）**行内两处**同义失实句都改
  （只改末句会把同一行改成自相矛盾）：
  - 行中：「四项里少任何一个，`findEngine` 会跳过**整个**候选目录 → **安静回退视觉桥（不报错）**」
    → 改为「会跳过**该**候选目录、**继续往下一个候选找**（`continue` 只跳过这一个目录；只有全部候选
    都不合格才退回纯视觉路径）」；
  - 末句：「否则整个候选目录被跳过、安静回退视觉桥」→ 改为实测口径（**该候选目录被跳过** → 继续
    往下一个候选找 → 本机常落到 `vendor/` → 运行侧仍传 `--models=<自定义>` 指向 vendor 下不存在的
    目录 → 引擎空跑、逐倍率重试后才升级视觉桥；禁用视觉桥时表现为整单报错），并注明
    **（第九轮更正：原「安静回退视觉桥」在本机不成立）**（保留修订轨迹）。
  该行其余内容（`rapidocr.js:111-113` / `:134-141` / `:493-495`、`DEFAULT_MODELS` 四文件名、
  空串未设回落、`rapidocr.js:42-47`）一字未动。
- `README.md` L118-119 同款口径「找到即用，找不到自动退回纯视觉路径」→ 改为「一个候选目录不合格只会被
  **跳过**、探测继续往下走，只有**全部候选都不合格**才退回纯视觉路径」，并补一句指针：
  `_MODELS`/`_DET/_CLS/_REC/_KEYS` 只影响**运行侧参数**，不参与这份可用性检查（机理详见 SKILL.md）。
  README 保持速查层，未搬 W4 机理。

**规模与验证**：`SKILL.md` 58868 → **59448 字节**，521 → 521 行，**逐行比对唯一差异行 = L193**
（其余 520 行字节合计 57418 = 57418；UTF-8 无 BOM、纯 LF、末字节 `0x0A` 与 live 一致）。

**范围说明**：原单 §4 把 `SKILL.md` 标为「本批无人改」（依据是裁定 2=b、⚠ 注记保持原样）；
该依据被 §4b 实测证伪后，经**负责人明确授权**（裁定：修正），本项随批落地。
`SKILL.md` 由 `skill.js` 每次读盘（`skill.js:47-48`，注释明写 "without remounting"）⇒ **改完即时生效**
（但读的是 **live** 副本）。

### 6. 回归（逐条实测）

| 命令 / 探针 | 结果 |
| --- | --- |
| `node test.mjs`（工作副本集成后，升华权限） | **`✅ ALL PASS  (586 checks)`** / EXIT=0 / 0 FAIL |
| `node repro.mjs` | **`🟢 GREEN`** / EXIT=0（返回值可无损序列化） |
| `R9-3` 两条新断言 | 均 PASS —— content 通道 `fallback=sidecar · 1416 字符 · 脚注 1 次`；notice 通道 `655 字符 · sidecar说明@156 < 上限脚注@480` |
| 新增断言反向自证（跑未修复 baseline） | 2 条**全 FAIL**（非同义反复） |
| 工作副本 vs live 逐文件 `Get-FileHash`（批写后） | 批写前差异**仅限本轮 5 个文件**（其余 **22** 个共有文件逐字节相同）；批写后 **5/5 全 `same=True`**；`AGENTS.md` 1328 B / SHA256 `84D0FDD0…5413` 与候选副本逐位一致，`AGENTS.md.bak-20260925` 备份同目录落地 |
| W6 两侧逐 hunk 对账 | 4 文件共 **91** 差异 hunk ⇒ **有意差异 90 / 真实偏差 1**；行级逐字一致 **6243** 行（DSH 独有 2630 / WB 独有 1829）；`R9-3:` 两条断言 WB 侧 **0 命中** ⇒ 本侧新增、不计偏差；SKILL.md 字节账确认**仅 L193 一行**不同；三侧全 **LF 无 BOM** ⇒ 批写无行尾风险；唯一 K 见 §8、已登记 §9 |

**⚠ 沙箱归因（必读，避免下轮误判）**：**受限沙箱（workspace-write）下 584 / 586 都不可复现** ——
`pdf-render.ps1` 走 WinRT，沙箱拒其路径授权（实测
`System.UnauthorizedAccessException：拒绝访问…该项目没有位于应用程序可以访问的位置`），
`pdfOutputGate` 因此在**第 23 条 PASS 后**以 **`EXIT=2`** 硬中断、**没有 `N checks` 尾行**；
换成 ASCII 产物目录、手工直调脚本、5 个不同 PDF（952 B–6.7 MB）均同样失败 ⇒ **不是 PDF 问题、
不是中文路径问题、不是 node 管道 EPERM 问题**（后者现象真实存在，但 `index.js:2625` 已有 `stdio:'ignore'`
兜底，失败在其后）。**黄金回归必须用升华权限在沙箱外跑**：本批 586 checks 即由升华权限跑出，
且同一命令升华后 R13-2 的 WinRT 探针由 `plainPng=0` 转为 `plainPng=1`（`aes128Png=1`）。
反证：live 那轮 `test-out\r13-winrt-plain\page-1.png`(12213 B) 等与 `test-run.log`(08:07:09) **同轮**落地
⇒ 基准日 WinRT 与 live 写入都可用，**那轮 584 不产生于受限沙箱**。
诊断跑（`DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` + ASCII 产物目录，**已改动环境，不得当验收证据**）：
**582 checks / 4 FAILED**，除 2 条由 SKIP_RENDER 造成的诊断扭曲与 2 条纯沙箱假失败（`plainPng=0`：连明文
PDF 都渲染不出）外 **578 条全 PASS** ⇒ 沙箱假失败**只集中在 WinRT 渲染**。全量耗时 ≈7 分钟且
**必须独占**（`index.js:235-236` 注释「4 路并发会炸」的瞬态竞态）；`repro.mjs` ≈1 秒，可作每次改动的快速哨兵。

### 7. 并发拓扑与落盘方式（留档）

原单设计「Wave A **5 路并发**（W1 README 对账 / W2 AGENTS.md / W3 item3 / W4 item4 证据 / W5 回归场地）
+ Wave B **2 路**（B1 主管集成 / B2 W6 两侧对账）+ Wave C 主管串行收口」。实测执行：
- 子代理**一律只 stage 到 workspace，不写 live**（live 写入被沙箱拒，实测 `file access denied`）；
  live 由主管用**一次**升华权限**逐文件**批写（含 `AGENTS.md.bak-20260925` 备份），**禁止整目录覆盖**；
  工作副本回写后逐对 `Get-FileHash` 确认 `same=True`。
- **踩到并留档的一个坑**：本机 `git core.autocrlf=true`，**裸 `git apply` 会把整份 `index.js` 改写成 CRLF**
  （221890 → 227811 字节、4173 行全部带 `\r`）。识别方式：与正确产物**行数相同、字节差恰等于行数**，
  且 `git diff` 归一化后**判为 0 差异**。正解 = `git -c core.autocrlf=false apply -p1`，
  或**直接复制**已校验的纯 LF 产物；本批最终采用后者（工作副本 `index.js` = W3 的 `work-index.js`，
  223638 字节 / 4173 行 / CRLF=**0**）。⇒ **以后凡涉及行尾敏感文件，禁用裸 `git apply` 写盘。**
- `stage\AGENTS.md` 会被宿主当作 **workspace 指令注入本会话**（已实测发生），已重命名为
  `AGENTS.md.candidate` 以免污染；**凡把 AGENTS.md 的候选副本放在 workspace 内，都必须改名或移出。**

### 8. W6 对账结论

### 8. W6 对账结论（DSH 工作副本 vs WB 侧，只读、字节级）

> 方法：`ReadAllBytes` / `ReadAllLines` / SHA256 + 自研逐行 diff（recursive patience + DP-LCS），
> **未用 git、无行尾归一化**；每份 diff 均通过「逐行重建校验」（`reconstructA/B = true`）；
> 4 个改动文件另做「未变行字节 + 差异行字节 + 行数 × LF = 文件字节」**零残差账**，
> 证明差异**只**落在列出的行号上（无隐藏行尾/编码噪声）。

**三类计数（hunk 级，4 个本轮改动文件 vs WB，共 91 个差异 hunk）**：有意差异 **90** / 真实偏差 **1**
（逐字一致的 hunk 为 0 —— hunk 即差异块）。行级（LCS）：逐字一致 **6243** 行｜DSH 独有 2630｜WB 独有 1829。

| 文件 | DSH vs WB（字节/行） | 差异 hunk | 逐字一致行 | DSH 独有 | WB 独有 | 有意差异 | 真实偏差 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| README.md | 58303/506 ↔ 37041/379 | 9 | 261 | 245 | 118 | 9 | 0 |
| SKILL.md | 59448/521 ↔ 63679/629 | 5 | 131 | 390 | 498 | 5 | 0 |
| index.js | 223638/4173 ↔ 225422/4205 | 57 | 3687 | 486 | 518 | 56 | **1** |
| test.mjs | 245474/3673 ↔ 194356/2859 | 20 | 2164 | 1509 | 695 | 20 | 0 |

**有意差异的归属（每条都指向具体轮次）**：
- **WB 侧独有（本侧红线不做）**：GPU OCR 引擎栈（python-dml / DirectML，WB 第五/六/八轮）—— SKILL 的
  GPU 整节、`index.js` 的 `engine:` 段 / `engines` 记账 / `noteBatchMeta` / `engineRole` / `localSoftText`
  「降级保留 + uncertain」/ `ocrUncertain*`、`test.mjs` 的 P17 不变量断言；以及 WB 的 CLI 入口件
  `api.mjs` / `doctor.mjs` / `ttf.js` / `repro-evidence.md`（四件 WB 独有）。
- **DSH 侧独有（WB 明确不移植）**：质量门 `structural` 启发（DSH 第四轮；WB 第十一轮评估 +
  第十二轮 §5「故意不移植」，附 741 页 0 真命中 / 3 页误杀证据）。
- **同一能力的各自实现**：第十二轮「阶段二 PDF 产出质量门 + 插图」（DSH 第五/六/七轮 ↔ WB 第十二/十三轮；
  是**同一能力块在两侧的位置/实现不同**，非能力缺口）；AES-256（DSH 第六轮 ↔ WB 第十二轮，事实一致、
  编号与详略不同）；recalc（DSH 有 `DSH_OFFICE_RECALC_MAX_CELLS` 护栏、WB 第十二轮明写「本侧无护栏」）。
- **调用语法 / 溯源标注**：`office_read ocr="always"` ↔ `read --ocr=always`；WB 给移植件加的
  「（自 DSH 侧移植）」前缀 = WB 第十一轮溯源。
- **本轮第九轮**：README §1a/§1b/§1c 改写与「过门 ≠ 可直接引用」新增、SKILL L193（§5）、
  `index.js` 的 `ocrCapNotes` helper + 两处调用、`test.mjs` 的 `R9-3:` 两条 —— 在 WB 侧全部 **0 命中**。

**`R9-3:` 两条断言 = DSH 侧独有**：WB `test.mjs` grep `R9-3` = 0、`降级不丢` = 0；两条断言整体落在
WB **完全不存在**的 DSH-only 块内；对应 helper `ocrCapNotes` 在 WB `index.js` 亦 0 命中
⇒ **本侧新增，不计偏差**。

**SKILL.md 仅 L193 一行（字节级证明）**：live → 工作副本 = 1 del / 1 ins（equal = 520，reconstruct = true），
差异行号清单 **= [193]**；其余 520 行内容字节合计 **56898 = 56898**（+520 LF = 57418，与本条目 §5 同值），
两侧 521 行 / 末尾 `0x0A` / LF / 无 BOM。⇒ 与原单 §4「本批无人改 SKILL.md」的差异，已由 §5
「负责人授权的第 5 项」消解，**非越界**。

**行尾 / BOM（全量 33 名字 × 3 侧）**：本轮 4 个改动文件在三侧均 **LF + 无 BOM + 末字节 `0x0A`
⇒ 批写回 live 无行尾风险**。唯二例外都在本轮写集**之外**：`REMOVED-MODELS.md`（DSH live/工作副本 =
**CRLF×32**，WB = LF 44 行；**日后跨侧搬运必须按字节复制**）、`repro-run.log`（仅 live，UTF-16LE + BOM，
PowerShell 产物）。全程未调用 `git`，不存在 `core.autocrlf=true` 造成的对账噪声。

**附带观察 O1（低危，不列 K，已在批写前落实）**：`test.mjs` 乱码夹具段的注释原写「降级 notice/content
**会覆盖** OCR 记账脚注」，与本轮修复后的行为直接矛盾 ⇒ 已就地补写修订注记（**原文保留**，追加
「第八轮记录；第九轮 R9-3 已修」并指向紧邻的 `R9-3:` 两条断言）。

### 9. 已知缺口（登记给第十轮，**本批不修**）

**`index.js` PDF 分支缺 `out.__fullBody`** —— 唯一真实偏差 K = 1，**第三轮遗留、非本轮引入**：
消费端 `index.js:833-839`（`finishRead` 读 `out.__fullBody`），生产端只有**非 PDF 分支** `index.js:3901`；
PDF 分支的返回对象（`index.js:3619-3627`）没有这一行 ⇒ PDF 走降级时 `sidecarBody` 取的是**已被内联护栏
截断的** `out.content`，而 notice 承诺的是「整篇正文」，**名不副实**。WB 侧在 `index.js:3670-3676`
反向标注了这一点（`if (capped.truncated) out.__fullBody = content`，L3676）并注明
「**⚠ 反向待确认项（WB 侧保留，DSH 第三轮误删）**……建议 DSH 侧下一轮补回」；
**live r8 与第九轮工作副本完全相同 ⇒ 遗留确认**。本批**不硬塞**（ESM 改动需重启 profile + 全量回归，
且超出本批 5 项范围）：登记为已知缺口，**指定第十轮修**（约 1 行改动，可借 WB 侧既有回归用例）。

## 2026-09-24 — 第八轮（DSH 侧）：SKILL 文档与代码对账 / 短句判据回移 / ocrPreview 记账 / pptx 条目修正

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。上一批基线 **563 checks**。
> 本批只做 6 项、逐项回归；**不动 OCR 引擎栈**（本侧 = CPU RapidOCR + 视觉桥），
> **不动** `PDF_PARSER_VERSION`、`DSH_OFFICE_CACHE_DIR`、sidecar 协议与续读协议
> （`offset + content.length === nextOffset`）。补丁留档 `work\dsh-office-r14\`（patch + apply.mjs）。

### 1. SKILL.md「已知边界」的 pptx 条目已过时 → 改为「已修（第七轮 notesMaster 独立主题）」

原条目（写于第五轮）仍写着"**⚠ 未解决**……元凶在 `presentation.xml` / `slideMasters` / `slideLayouts` /
`notesMasters` 及其组合之内，**未定位**"，并给"要交 pptx 就先 `office_convert` 成 docx/md，
或用 PowerPoint 模板另存"的绕行指引。**两条都已失效**：第七轮 bug#2 已定位并修复（根因 =
`notesMaster1.xml.rels` 把备注母版主题指向与 `slideMaster` **共用**的 `../theme/theme1.xml`；
PowerPoint 要求 notesMaster 有独立主题部件）；补丁 = `pptx.js` 新增 `THEME_NOTES` +
`ppt/theme/theme2.xml` + ContentTypes Override + `NOTES_MASTER_RELS` 改指 `theme2`，当日 563 ALL PASS。
改写后保留**历史诊断链一句话注记**（zip 层 / OPC 不变量 / 非必需部件逐个删 / 叶子部件互换全部排除，
与 `presentation.xml` 顺序/空白/`sldId`/`notesSz`/占位符 idx 全无关），删除"未定位"与 convert 绕行指引。

### 2. 补「过门 ≠ 可直接引用」警告（自 WB 侧第七轮 §7.1 移植，落点「已知边界」）

WB 侧有、本侧完全没有的一条：通过质量门 **≠ 文本无噪声**。数字/正负号密集的表格页务必核对原 PDF，
典型症状**负号被丢**（`-9.8 → 9.8`）；依据 510 数据格 4.9% 出错、密集表 10–15%，置信度 0.991 的页
也能错 12.5%，并有"A.3 数成 D.6"的可复现反例（`分析6-p7`）。措辞按本侧实际**反向强调**：
视觉复核升级**同样防不住**这类版面级噪声（它治"没认出字"，不治"字对了但负号丢了"），
**没有**写成"升级视觉即可解决"。

### 3. 短句判据：先核实出"代码只有一半"，按裁定移植（代码 + 文档双补）

**核实结论（与工单描述不符，已先报告并取得裁定）**：本侧 `rapidocr.js` 原只有判据的**前半套** ——
`shortRatio=0.6` / `shortMinBoxes=30` 是**硬编码常量**（不可配）、判据**无条件判硬**（无"置信 ≥ 阈值
即放行"）、`gateResult` 只返回 `{hard,reason}`（**无 `retryable`**），`index.js` 重试循环对判 hard 的页
**一律换倍率重试**（不存在"结构性失败短路"）；全库 `grep DSH_OFFICE_OCR_SHORT_*` = **0 处**。
即"代码在、文档无"只对了一半。

**裁定：移植 WB 侧判据到本侧**（而非"只补文档"），改两处：

| 文件 | 改动 |
| --- | --- |
| `rapidocr.js` | `GATE` 新增 `shortRatio` / `shortMinBoxes` / `shortRatioScore` 三个**环境变量覆盖项**（默认仍 `0.6` / `30` / `0.95`）；`gateResult` 引入 `hard(reason, retryable=true)`，短句分支改为 `… && r.avg < gate.shortRatioScore` 才判硬，且 **`retryable=false`** |
| `index.js` | 本地批次 `hard.push(...)` 带 `retryable: verdict.retryable !== false`；重试段按 `retryable` / `structural` **分流**，只对 `retryable` 渲染与重试，`structural` 直接记 `（结构性失败，换倍率无效）`（页仍照常交视觉桥） |

**未移植**（WB 特有，超本批）：`info.localSoftText` 无视觉桥降级保留、`noteBatchMeta`、
`localBatchCap`（GPU 批次上限）——本侧视觉桥路径不需要，逐字保留原样。

**文档落点**：§「读取的正确顺序」第 6 条补短句语义三条；环境变量总表补三个 `SHORT_*` 行。

### 4. 环境变量总表收拢（目标"总表 = 全集"）

逐项核实代码真实默认值后收进：`DSH_OFFICE_OCR_SHORT_RATIO`(`0.6`) / `..._SHORT_MIN_BOXES`(`30`) /
`..._SHORT_SCORE`(`0.95`) / `DSH_OFFICE_OCR_MIN_SCORE`(`0.88`，`Number(env)||0.88`，`0` 也回落) /
`DSH_OFFICE_RAPIDOCR_DIR`（未设，候选链**最高优先**）/ `DSH_OFFICE_OCR_DISABLED`（未设，
**任何非空值**即关本地引擎）/ `DSH_OFFICE_BLANKS`（未设，`1|true|on|yes`，只改呈现）/
`DSH_OFFICE_PDF_EMBED_CJK`（未设 = **开启**，`=0` 回退旧 STSong 路且被字节级门槛拦）/
`DSH_OFFICE_PDF_SKIP_RENDER_CHECK`（未设，`=1` 只跳渲染级抽检）。

> **纠正工单一处**：`DSH_OFFICE_TEST_AES256_PDF` **已在总表内**（第七轮随 AES 判别实验写入），
> 本批**不重复收**，只复核语义一致（未设时 `test.mjs` 走 skip 分支并打印提示）。

### 5. `skill.js` 虚假元数据（EPUB）

`CANDIDATE.description` 宣传支持 **EPUB**，但 `index.js` / `legacy.js` grep `EPUB` = **0 处**、
SKILL.md 能力清单也不含 EPUB —— 该文件是**每次 turn-0 目录的注册源**，模型据此选技能，属虚假宣传。
改为与实际能力面对齐：删 EPUB，补本侧真实支持的 `HTML (.html .htm)` 与
`Markdown and plain text (.md .txt .json .rtf)`。**只改本目录这份**（两侧共用文件，WB 侧不在本批）。

### 6. 回移 WB 侧 `ocrPreview` / `ocrPagesCapped` 静默预览记账

**兼容性评估：兼容，逐 hunk 回移**。两侧 OCR 页队列构造同构（本侧原为紧凑三元写法）、
`pageRanges()` / `notes` 汇聚点 / `stats` 汇聚点都在，无需改 `ocrPdfPages()` 签名；回移只增加记账，
**不动页序与缓存口径**（故 `PDF_PARSER_VERSION` 按约定**未动**）。

| 位置 | 改动 |
| --- | --- |
| `index.js` OCR 页队列 | 展开为显式分支，记录 `ocrPreviewLimited = {preview,total,skipped}`（未指定 `pages` 只做前 `OCR_DEFAULT_PAGES=3` 页）与 `ocrPagesCapped = {limit,requested,applied,skipped}`（被 `OCR_MAX_PAGES=20` 砍掉） |
| `index.js` notes | 两条脚注 + **可复制的续读命令**（`office_read path="…" ocr="always" ocrEngine="local" pages="…"`） |
| `index.js` stats | `stats.ocrPreview` = `{preview,total,skipped}`；`stats.ocrPagesCapped` = `{limit,requested,applied,skipped}`（形状与 WB 侧逐字一致） |
| `SKILL.md` 第 4 步 | 写明两种"少做"都不再静默 + 续读命令 |

### 7. 收口批次：3 路并发子代理 + 主代理落盘（同日晚）

本批 6 项交付后，按"多任务并发"要求拆出 3 个子任务并发执行；**子代理无沙箱升级权限**
（写 live 插件目录、跑全量 `test.mjs` 均被 EPERM 拒绝，实测确认），故一律由主代理落盘并跑回归。

| 子任务 | 独占范围 | 结果 |
| --- | --- | --- |
| **P1** 补测试断言 | `test.mjs` | 新增 **21 条**（gateResult/阈值覆盖 11 + ocrPreview 3 + 文案通道 3 + ocrPagesCapped 4）；交付补丁 + dry-run + 探针实测，由主代理落盘，`node --check` 通过 |
| **P2** 环境变量总表全集化 | `SKILL.md` | 12 个缺口逐个核实语义与默认值后收进文档：总表 **+2 行**（`OCR_MODELS/_DET/_CLS/_REC/_KEYS` 合一行 + `OCR_NO_ANGLE`）、「排查」**+1 bullet 块含 4 子项**；复核 **code_count=34 / doc_count=34 / gap=0 / doc_only=0**（SKILL.md 54614 → 58868 字节） |
| **P3** 两侧逐 hunk 对账 | 只读 | 逐字一致 **19** / 有意差异 **10** / **真实偏差 1**（见下） |

**P3 的唯一真实偏差（已修）**：OCR 脚注 push 顺序与 WB 相反（DSH 原为 preview→capped，WB 为 capped→preview）。
两条文案本身逐字一致，仅当同一次读取同时命中"预览截断 + 20 页上限"时脚注行序不同。
已交换两个 `if` 块（`index.js` +110 字符），零行为影响。

**P2 查出的死钩子 / 未接线项（如实入档，不是"沉默凑 0"）**：
- `DSH_OFFICE_NATIVE_SRC` / `_NATIVE_PDF` / `_PPTX` = **写入侧 IPC 通道**：插件代码**零读者**，
  唯一读点在 `test.mjs` 内嵌 PowerShell 的 `$env:`，且每次 spawn 无条件覆盖 → 外部 export 无效。
- `DSH_OFFICE_RENDER_QUIET` = **未接线的降噪开关**：`index.js` 只往渲染子进程塞 `DSH_OFFICE_RENDER_SCALE`，
  从不设 QUIET，也不读脚本 stdout（成功判据是 PNG 在盘上）→ 只对手工跑 `pdf-render.ps1` 有用；
  且是全项目唯一的"精确 `'1'`"语义（`true`/`yes` 无效）。
- **新发现**：`DSH_OFFICE_OCR_MODELS` 存在**探测与运行不同源** —— `findEngine` 写死看
  `<引擎目录>/models/` 里有没有 `DEFAULT_MODELS` 四个文件，**不读** `OCR_MODELS`；改模型目录时
  引擎目录仍须保留 `models/` + 四个默认模型，否则候选目录被跳过、安静回退视觉桥
  （`_DET/_CLS/_REC/_KEYS` 同理）。

  > **第九轮更正（实测，见第九轮条目 §4b）**：本项末句「安静回退视觉桥」在**本机不成立** ——
  > `continue` 只跳过**那一个**候选目录，`findEngine` 会**继续往下一个候选找**（本机落到
  > `vendor/RapidOCR-json_v0.2.0`），运行侧仍传 `--models=<自定义>` 指向 **vendor 下不存在的目录**
  > → 引擎空跑 → 逐倍率重试后才升级视觉桥（禁用视觉桥时表现为整单抛错）；只有**全部候选都不合格**
  > 才退回纯视觉路径。**原文保留（修订轨迹不删）**；SKILL.md L193 与 README L118-119 已按本轮实测更正，
  > 「探测与运行不同源」这个**核心发现本身依然成立**。

**P1 暴露的产品缺陷（本批新增用例的价值所在，已修）**：`pdfOutputGate` 的辅助"空白"判据
`chars >= 24 && firstPageBytes < 30KB && m.ink < 0.001` 里 **`chars` 是整篇可见文本量，而渲染只抽首页**
—— 对"首页少字、后续页内容多"的多页文档必然误判。实测：26 页夹具首页一行英文、整篇 407 字
→ 判"文本量可观却渲染近乎全白"（首页 PNG 10116 字节、墨迹 0.039%、544 px）→ **拒绝落盘**。
修法（最小且口径同源）：新增 `pageCount`（由 `model.blocks` 里的 `pagebreak` 计数 +1），
辅助判据改为 `… && pageCount === 1 && …`；**主判据（整页几乎无墨 `< 40 px`）对所有文档仍然生效**，
单页文档行为逐字不变（"渲染为空白 → fail 且不落盘"用例照旧 PASS）。

**P1 修订过程（如实留档）**：第一版落盘跑出 **581 checks / 1 FAILED**，失败的是「页数上限：正文脚注明说…」
—— 真实产品事实是：读 `pages="1-30"` 时 21-30 页不做 OCR、保留乱码文本层，整篇被质量门判"不可读"
走 sidecar 降级，**降级 notice/content 覆盖了该记账脚注**（content 里出现 0 次，`stats.ocrPagesCapped` 仍在）。
故该条改为断言「降级不丢记账（stats 仍给上限、notice 仍给下一步命令）」，并**新增 B2 组**用
"26 页有文本层文档 + 缓存命中"把脚注**文案通道**真正测到。

> **已知缺口（本轮未修，待决策）**：`fallback="sidecar"` 降级路径下，OCR 上限脚注在**正文**不可见
> （只剩 `stats.ocrPagesCapped`）。不是完全静默，但可观测性弱于正常路径；修它要动降级 notice 的组装，
> 建议单独一批处理。

**沙箱对照实验（结论：本会话无法给出干净对照，如实记录）**：原计划用改动前快照
（`work\office-enhance2\plugin\`）在受限沙箱下复现"561 checks / 2 FAILED"以钉死归因。实测该快照在受限
沙箱下**更早**就崩：`test.mjs:399` 的 PDF 产出质量门报"第 1 页渲染失败（exit=1）"，只跑到 **23 条 PASS**
即 HARNESS ERROR 退出（快照产物目录在工作区内、路径含中文，渲染管线同样受管）。
⇒ `R13-2` 两条 FAIL 的归因仍由**首次受限跑的实测证据**支撑（产物回退
`%TEMP%\dsh-B5SvcN\dsh-office-test-out` + `spawnSync(powershell.exe)` 默认 pipe 的 `w1txt` 为空），
而不是由该对照实验支撑。**"受限沙箱下渲染抽检也会失败"这条记录本身有价值**：以后在受限环境跑套件，
别把早期 `exit=1` 当成代码问题。

**工作副本同步**：本批是 direct-to-live（直接在 live 改），收口时把 live 的
`index.js` / `rapidocr.js` / `SKILL.md` / `skill.js` / `test.mjs` / `CHANGELOG.md` **逐文件**回写到
`work\office-enhance2\plugin\`（禁止整目录覆盖），并以 `Get-FileHash` 逐对确认 `same=True`。

### 回归（同日，逐条实测）

| 命令 / 探针 | 结果 |
| --- | --- |
| `node test.mjs`（插件目录，完整模式） | **584 checks / ALL PASS**（= 基线 563 + 收口批次新增 21 条断言，FAIL 0；⚠ 仅当插件目录可写且子进程管道可用时；见下「沙箱假失败」） |
| `node repro.mjs` | **🟢 GREEN**（退出码 0，43 页样本干净、0 坏码点、结构无损） |
| `office_create` 3 页 pptx → PowerPoint COM `Presentations.Open` | **`OPEN_OK slides=3`**（真开，非插件自读；读出第 1/3 页标题） |
| `test.mjs` 内 R13-7 正向 + 负向控制 | `PPT_OPEN_OK slides=3` / 改回共用 `theme1` → **仍 `0x80070570`**（改对才开、改错必不开） |
| 改判据后既有 3 条质量门断言 | 仍 PASS：分别命中"置信度低"、"版面复杂"、"shortRatio 0.4 < 0.6 放行"，与新的短句联动**不冲突** |

> **⚠ 沙箱假失败（本轮新增经验，值得留档）**：同一套件在**受限沙箱**（会话 `workspace-write`）下跑出
> **561 checks / 2 FAILED**，失败项是 `R13-2 判别实验` 两条，`plainPng=0`。**不是代码问题**，两个环境成因：
> ① 插件目录不可写 → `test.mjs` 把产物回退到 `%TEMP%\dsh-XXXX\dsh-office-test-out`（基线写在插件目录
> `test-out`）；② `test.mjs:3368/3377` 用 `spawnSync('powershell.exe', …, {encoding:'utf8'})`（**默认
> `stdio: 'pipe'`**）跑 WinRT 渲染探针 —— 受限模式**禁止子进程开命名管道**（文档化边界，EPERM），
> 探针实际没跑、`w1txt` 为空。解法：在允许插件目录写 + 子进程管道的环境下复跑，即 **563 ALL PASS**
> （同基线的 `test-out` 路径）。**判据**：`R13-2` 这两条 FAIL 若伴随 `产物：…\Temp\dsh-…`，先查环境，
> **别改加密代码**。

> 生效提示：`SKILL.md` 由 `skill.js` 每次读盘，**改完即时生效**；本轮改了 `index.js` / `rapidocr.js`，
> **需重启 profile** 才生效。

## 2026-09-24 — 第七轮（DSH 侧）：docx 插图两处真 bug / 图片解码面 / 公式子集扩展 / AES-128 对照

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。工作副本
> `<work>/deepseek1\work\office-enhance2\plugin\`（`vendor/` = junction → live）。
> 上一批（第六轮）基线 **516 checks**；本轮跑绿后一次性回写。需求来自《dsh-office 插件增强 ·
> 第二轮收口批次》九项。诊断留档 `work\office-enhance2\probe\`；先报告清单与接口草案
> `work\office-enhance2\DRAFTS-r13.md`。

### 病灶 / 根因（本批挖出**两条真 bug**，都在"从来没被原生 Office 验证过"的路径上）

| # | 病灶 | 根因 | 补丁位置 |
| --- | --- | --- | --- |
| 1 | **含图 `.docx` 100% 被 Word 16.0 拒开**（`0x80070570 The file or directory is corrupted and unreadable`）—— 第六轮的插图链路只验过 zip 结构、`rId` 唯一性、OOXML 包体检，**三项全绿**，但 Word 一律打不开 | `docx.js` 的 `WP_NS` 常量是**不存在的命名空间**：`…/drawingWordprocessingDrawing/2006/main`（正确值 `…/drawingml/2006/wordprocessingDrawing`）。Word 解析 `wp:inline` 时找不到元素定义 → 直接判"文件已损坏"。XML 语法合法、包体检不查 URI 拼写，所以这条路径**一路绿灯** | `docx.js::WP_NS`（一行） |
| 2 | `.pptx` 被 PowerPoint 16.0 拒开（第五轮遗留、第六轮未解） | `ppt/notesMasters/_rels/notesMaster1.xml.rels` 把备注母版主题指向与 `slideMaster` **共用**的 `../theme/theme1.xml`；PowerPoint 要求 notesMaster 有**自己独立的**主题部件（它自己保存时分配 `theme2.xml`）。与 `presentation.xml` 元素顺序/空白/`sldId`/`notesSz`/占位符 idx **全部无关**（已单变量排除） | `pptx.js`：新增 `THEME_NOTES` 常量与 `ppt/theme/theme2.xml` 部件、`contentTypes()` 加 Override、`NOTES_MASTER_RELS` 改指 `theme2.xml` |

### 诊断链（bug #1 是怎么定位的 —— 值得留档，因为所有"结构体检"都骗过了）

1. 观察：无图 docx 能开、含图的全不能开；把图加到**真实 Word 文档**副本上同样打不开。
2. 单变量变体（`probe/mkvariants.mjs`，每个变体用独立 Word 实例打开）：删掉 `w:drawing` 段 → **OK**；
   换成空 `<w:drawing/>` → FAIL；换图片字节（夹具 PNG / GDI+ 重编码 PNG / 真实 Word 的 PNG）→ **都 FAIL**；
   把 `wp`/`a`/`pic` 声明补到根元素 → FAIL；补 `word/theme/theme1.xml` + theme 关系 → FAIL；
   media 改 store 压缩 → FAIL；调换 `[Content_Types].xml` 里 png `Default` 的位置 → FAIL。
   ⇒ 与图片字节、关系、主题、压缩、内容类型顺序**都无关**。
3. 决定性对照（`probe/authordrawing.mjs`）：把**真实 Word 文档里那段 `w:drawing` 原文**塞进真实包 → **OPEN_OK**；
   换成我们生成的 `w:drawing` → FAIL。⇒ 元凶锁定在**我们写出的 drawing XML 文本**里。
4. `probe/dump-ns.mjs` 对比两边根元素的 `xmlns:wp`：真实是
   `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`，我们是
   `http://schemas.openxmlformats.org/drawingWordprocessingDrawing/2006/main` —— **URI 拼错**。
5. 改这一行 → `Word 16.0 OPEN_OK shapes=2`（`probe/fixcheck.mjs`：`office_create` 含图产出 + `office_edit` 再插一张）。

### 本批交付（九项需求）

| 需求 | 交付 | 关键点 |
| --- | --- | --- |
| 1 docx 写出端 image 分支 | `docx.js::writeDocx` / `docBlocksToXml` / `appendBlocksToDocument` 支持 `image` 块 | `word/media/imageN.<ext>` + image 关系 + `w:drawing`；`IMAGE_EMBED_TARGETS` 加入 `docx`；内容 SHA-256 去重（复用部件与关系 id）；尺寸规则显式化；**读不到 / 格式不认识**退化为字面文本 + `imagesSkipped` + notice |
| 2 AES-256 可信度 | AES-128（R4/V4+AESV2）**对照夹具**（与 AES-256 同一骨架）+ 判别实验 + 真实样本槽位 | 见下「交叉验证」 |
| 3 公式重算 | 新增 `SUMIF/AVERAGEIF/COUNTIFS` + 文本函数族 11 个；区域文本/布尔按真 Excel 口径；整列/外部/`_xlfn`/循环引用显式记账；`DSH_OFFICE_RECALC_MAX_CELLS` 护栏 | 见下「公式口径」 |
| 4 插图链路收口 | PNG 索引图（`PLTE`/`tRNS`）、BMP（8/24/32-bit）、GIF 首帧（透明索引 + 交错）；媒体内容去重；PDF `align` | 见下「解码面」 |
| 5 PDF 合并 / 拆分 | **只交接口草案**（红线 4：先报告） | `DRAFTS-r13.md` §4 |
| 6 OMML 公式写入 | **结论：不做** + 量化依据 + 重开条件 | `DRAFTS-r13.md` §7 / SKILL「已知边界」 |
| 7 pptx 拒开 | 见 bug #2 | PowerPoint 冒烟 + 负向控制 |
| 8 `.html` 结构化编辑 | **只交草案**（红线 4） | `DRAFTS-r13.md` §5 |
| 9 WB 分叉清单 | 见文末「分叉清单」 | — |

### 交叉验证（需求 2）：AES-128 判别实验

第六轮记录的"夹具与解析器同源"风险，本轮用**同一骨架的 AES-128 夹具**做判别（`test.mjs::buildAes128Fixture`：
Algorithm 3 算 `/O`、Algorithm 5 算 `/U`、Algorithm 1 对象密钥、AES-128-CBC + PKCS#7）：

| 探针 | 结果 |
| --- | --- |
| 自造 R4/V4+AESV2 夹具 → `office_read` | 读出正文（加密方向与既有 AES-128 解析路径自洽） |
| 同夹具 → WinRT（`pdf-render.ps1`，微软实现） | **能打开并渲染出 PNG**（`aes128Png=1`）⇒ **骨架正确**（xref/trailer/`/Encrypt` 位置与 R4 字典都没问题）⇒ 第六轮 AES-256 的 `OPEN_FAIL` 因此**指向"WinRT 不支持 R5/R6"**，而不是我们夹具的 V5 字典细节。该条用例的**判据是"探针本身有效"**（明文对照能渲染出 PNG），AES-128 的成败作为**结论**记录 —— 两条路径都算成功 |
| 自造 AES-256 R5/R6 夹具 → WinRT | 第六轮起已知 `OPEN_FAIL`，本轮复测维持 |
| 真实样本槽位 | `DSH_OFFICE_TEST_AES256_PDF` 未设 → skip 并打印原因；主人手上有真实 R5/R6 样本时一条命令即可关闭这条自证风险 |

**结论口径（需求 2a 的判别实验已给出明确结论）**：判据是"同一骨架的 AES-128 **能**被微软实现打开"，
所以 **① 骨架与加密字典位置正确**、**② WinRT 不支持 AES-256（R5/R6）**（而不是"夹具写错了"）。
本机仍然没有支持 AES-256 的第三方实现（WinRT 不支持 R5/R6、Edge/Chrome 无头不渲染本地 PDF、
无 qpdf/mupdf/Acrobat CLI），所以"与第三方逐字节一致"这条**仍未取得**；可复验的开关已内置（真实样本槽位）。
"绝不猜"的边界不变：`/Perms` 校验不过就报需要密码；口令候选**只有调用方给的那组**（默认空串）；
V5 下对象密钥 = 文件密钥（不按 obj/gen 派生）。

### 公式口径（需求 3，全部按真 Excel 语义）

- **区域里的文本与布尔一律忽略**：`SUM(B2:B6)` 不把 `"90"`/`TRUE` 算进去；直接写在参数里的仍换算
  （`SUM("3",TRUE)` = 4）。实现上 `flatArgs()` 给每个值带上 `range`/`scalar` 来源标记，`numericOf()` 据此过滤。
- **数值条件只命中数值单元格**：`COUNTIFS(range,"A",range2,">75")` 不把文本型 `"90"` 计入。
- **求和区同样忽略文本 / 布尔**：`SUMIF` / `AVERAGEIF` 的第三参数按区域处理。
- `MIN`/`MAX` 区域忽略空值、文本、布尔（空单元格返回 `''` → 被过滤）。
- 文本函数族按 Excel 语义（`LEFT`/`RIGHT`/`MID` 的负数与越界、`TRIM` 压内部空格、`TEXT` 数字格式子集、
  `TEXT` 的**日期格式显式 unsupported**、`VALUE` 解析千分位与括号负数）。
- **显式记账的边界**：整列引用 `A:A`、外部工作簿 `[Book1]Sheet1!A1`、`_xlfn.` 前缀、未加载工作表（`#REF!`）、
  循环引用（`#CIRC!`）—— 全部带单元格地址进 `stats.recalc.details`。
- **大表护栏**：`DSH_OFFICE_RECALC_MAX_CELLS`（opt-in）超限 → 整本原样返回 + `stats.recalc.skipped`
  （不静默截断、不半算）。

### 解码面（需求 4）

| 形态 | 处理 |
| --- | --- |
| PNG `colorType 3`（调色板） | `png.js` 解出索引 + `PLTE`/`tRNS`；`image.js::indexedPngToRaw` 展开 RGB/RGBA（`tRNS` 有 <255 才挂 alpha） |
| BMP 8-bit（调色板）/ 24-bit / 32-bit | `image.js::bmpToRaw`（自下而上与负高度两种行序；1/4-bit 与 RLE 返回 null 交调用方记账） |
| GIF 首帧 | `image.js::gifToRaw`（LZW 变长码解码、全局/局部调色板、Graphic Control 透明索引、交错行序重排） |
| CMYK JPEG / 1·4-bit BMP / RLE | 明确 `imagesSkipped` + 原因（不猜、不静默） |
| 内容去重 | docx `word/media/` 与 PDF `/XObject` 按内容 SHA-256 复用（部件/资源名 + 关系 id 都不新增），`stats.imageReused` 报次数 |
| PDF 对齐 | `block.align = center/right`；**缺省仍是 `marginX`，字节不变** |

### 实测对照（本机，工作副本；用例前缀 `R13-`）

| 用例 | 实测结果 |
| --- | --- |
| `R13-1` 无图 docx | 不多写 media / 图片关系 / 图片内容类型 / `w:drawing`（逐字节口径） |
| `R13-1` 尺寸策略 | `imageSizeFor(1200,800)` → 原始 900pt、超可用宽缩到 451.3pt；显式 `width` 优先、高度按纵横比 |
| `R13-1` 内嵌（翻转第六轮用例） | `md → docx` / `含图 HTML → docx` 的 `word/media/` **非空**、`rId` 唯一递增、包体检通过、`stats.imageMedia` + notice 尺寸 |
| `R13-1` 去重 | 同图两次 → **1 个媒体部件 + 1 条关系 + 2 个 `w:drawing`**，`imageReused` 1 条 |
| `R13-1` 失败路径 | 读不到 / 格式不认识 → 字面 `![alt](path)` 保留 + `imagesSkipped` + notice |
| `R13-1` 开箱 | **Word 16.0 打开含图 docx：`inlineShapes=2`**（修复前 100% OPEN_FAIL） |
| `R13-1` `.odt` 兜底 | 仍未内嵌的写出端继续降级 + 双账（负向回归） |
| `R13-4` 解码面 | 调色板 PNG / 24-bit BMP / GIF 首帧（含透明）都能进 PDF；坏字节返回 null 且记账 |
| `R13-4` PDF 去重 / 对齐 | 同图两次 → 1 个 `/XObject`、` Do ` 仍是 2 次；`align=center` 的 cm 位移不同、缺省不变 |
| `R13-3` 公式（直连，真 Excel 口径） | 区域文本/布尔忽略 → SUM=240 / COUNT=3 / MAX=90；SUMIF=160 / AVERAGEIF=80；`SUM("3",TRUE)=4`；`COUNTIF` 的 `?` 通配符=5 |
| `R13-3` 公式（端到端） | `SUMIF=220` / `AVERAGEIF=73.33` / `COUNTIFS=1`；`TEXT(1234.5,"#,##0.00")="1,234.50"`；`TEXT(…,"yyyy-mm-dd")` unsupported；`VALUE("1,234.5")=1234.5` |
| `R13-3` 记账边界 | 整列 `A:A` / 外部工作簿 / `_xlfn.CONCAT` / `NoSheet!A1`（`#REF!`）/ 循环引用（`#CIRC!`）全部带地址 |
| `R13-3` 护栏 | 5000 行 / 2500 公式跑完；设 `DSH_OFFICE_RECALC_MAX_CELLS=100` → `skipped` + 原值保留；删掉 env 恢复 |
| `R13-2` AES-128 对照 | 自造 R4/V4+AESV2 夹具读出正文；真口令报"需要打开密码"；不改写原文件 |
| `R13-2` 判别实验 | 明文对照渲染出 PNG（探针有效）；同骨架 AES-128 的 WinRT 结果作为结论记录 |
| `R13-7` pptx | notesMaster→`theme2`、slideMaster→`theme1`、CT 有 Override；**PowerPoint 16.0 能开（slides=3）**；负向控制（改回共用 `theme1`）**必须仍 FAIL**（实测 `0x80070570`）；两版读回逐字一致 |

### 基准

- `node test.mjs`（插件目录，完整模式）：**563 checks / ALL PASS**，退出码 0。
  轨迹：**516**（第六轮收口）→ 加 `R13-` 用例组（45 条）与两处既有断言改写 → **563**。
  准确数量以运行输出的最后一行为准。
- `node repro.mjs`：🟢 GREEN（退出码 0）—— 返回边界哨兵在新增字段
  （`stats.imageMedia` / `imagesSkipped` / `imageReused` / `imageSizing` / `imageFallback` 等）之后仍然全绿。

### 行为边界

- 既有参数语义不变；新能力全是"只增参数 / 只增 stats / 只加 notice / opt-in"。
- **无 image 块的 docx / PDF 产物逐字节不变**（`R12-1` 既有断言 + `R13-1` 的"不多写任何东西"断言）。
- 不静默：图片跳过 / 降级 / 复用 / 尺寸换算 / 未支持函数 / 循环引用 / 护栏跳过 —— 全部同时进 `stats` 与 `notice`。
- 零新增第三方依赖（PNG/BMP/GIF 解码只用自家代码 + `node:zlib`）。
- 未碰 `~/.workbuddy/`、`~/.dsh/cordis.patch.yml`、宿主装配、`dsh-badge` 合同、中文字体内嵌管线（`pdffont.js`）。

### 红线偏离 / 先报告清单

| 触碰点 | 处置 |
| --- | --- |
| **改写了第六轮 4 条 `R12-1` 降级断言** | 需求 1d 明确要求翻转；已改写为"内嵌成功 + 包体检 + Word 能开"，并**新增 `.odt` 降级负向用例**保住兜底覆盖 |
| **改写了 `R12-3` 的夹具公式** `CONCATENATE("a","b")` → `LET(1,2)` | 需求 3b 要求支持 `CONCATENATE`，与"把它当 unsupported 例子"的既有断言**互斥**；断言结构与统计口径未动，只换夹具公式（如实记录） |
| `editDocx` 插图**默认尺寸**由 450pt 改为"原图像素 × 72/96" | 需求 4e 要求的显式化；既有用例都显式传 `width`，零回归 |
| 需求 3a `recalc_formula` 写回 / 需求 5 PDF 合并拆分 / 需求 8 `.html` 结构化编辑 / 需求 4d 落点扩展 | **停在草案，未实施**；⚠ **主人已于 2026-09-24 明确确认：本批不做、草案保留**（下批若要动手，从 `DRAFTS-r13.md` §3–§6 起） |
| 需求 1c `.odt` / `.pptx` 同批内嵌图片 | **本批只做 docx**，其余继续降级兜底；⚠ **主人已确认暂不做** |

### 本批新增的 DSH↔WorkBuddy 分叉（只声明，不回移）

| 文件 | DSH 侧有什么而 WB 侧没有 | 性质 |
| --- | --- | --- |
| `docx.js` | `writeDocx`/`docBlocksToXml` 的 **image 分支** + `imageSizeFor()` + 内容去重（`{part, rid}` 表） | DSH 侧补齐写出端；**建议回移** |
| `image.js` | `bmpToRaw` / `gifToRaw` / `imageToRaw` / `indexedPngToRaw`（调色板 PNG 展开） | DSH 侧解码面扩展 |
| `formula.js` | `SUMIF`/`AVERAGEIF`/`COUNTIFS` + 文本函数族 + 区域文本/布尔口径 + `recalcMaxCells()` 护栏 | DSH 侧扩子集 |
| `pdf.js` | 图片 XObject **内容去重** + `align` + `imageSizing` 记账 | DSH 侧记账面 |
| `docx.js` | **`WP_NS` URI 修正**（WB 侧很可能有同样的错，回移前先核对） | 真 bug 修复 |
| `pptx.js` | notesMaster 独立主题部件（`theme2.xml`） | 真 bug 修复 |
| `index.js` / `test.mjs` / 文档 | `applyImageWriteInfo()`、`R13-` 用例组、`probe/` 诊断脚本、`DRAFTS-r13.md` | DSH 侧记账与测试 |

⚠ 反向分叉（WB 有、DSH 没有，本批**不移植**）：`ttf.js`（字体子集重写）、`api.mjs`/`doctor.mjs`（CLI）、
GPU OCR 链路、`ocrEngineMode`/`engineDecision` 观测面。是否回移由主人决定。

## 2026-09-23 — 第六轮（DSH 侧）：插图链路三端 / AES-256 读取 / 公式重算 / 稿纸网格

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。工作副本
> `<work>/deepseek1\work\office-enhance\plugin\`（`vendor/` = 指向 live 的 junction），
> 既有基线 **467 checks 全绿**，改完跑绿后一次性回写。移植源：**WorkBuddy 侧第十二轮**（只读取，
> WB 目录一个字未改）。需求来自《dsh-office 插件增强批次》提示词的四条。

### 病灶 / 起点：本批是"能力缺失"，不是既有 bug

第五轮之后 DSH 与 WB 分叉：WB 第十二轮已经实现了这四条，DSH 侧**一条都没有**。
提示词列的"现状事实"逐条核实后成立：

| # | DSH 侧现状（核实结论） | 补丁位置（本批） |
| --- | --- | --- |
| 1a | `model.js:132` 的 `markdownToDocument` **没有** image 分支：`![alt](path)` 走段落 → `parseInlineRuns` → 变成"`!` + 链接"，图片语义丢失（`html.js:408` 能产 image 块、`model.js:270` 能写 `![]()`，中间这一跳断着） | `model.js` 段落通道前插 image 解析（独立成行才认） |
| 1b | `index.js` 的 `editDocx` 只有 `replace_text` / `append_markdown` / `set_meta`，**没有插图操作** | `docx.js` 新增 `imageParagraphXml` / `insertImageParagraph` / `ensureContentTypeDefault`；`index.js` 的 `editDocx` 加 `insert_image` / `append_image` 分支（media 部件 + image 关系 + `w:drawing`） |
| 1c | `pdf.js` **零图片支持**（`grep image\|png\|jpeg\|XObject` 无匹配） | `pdf.js` 的 `writePdf` 加图片 XObject：JPEG `/DCTDecode` 原样内嵌、PNG 解成原始采样 `/FlateDecode`（带 alpha 挂 `/SMask`）；新增 `image.js`（嗅探 + PNG→raw），`png.js` 导出 `decodePng` |
| 2 | `pdfcrypt.js:103` 对 `R>=5` 直接返回"AES-256（R5/R6）加密暂不支持" | `pdfcrypt.js` 新增 `hash2B`（算法 2.B）/ `checkPerms` / `collectV5Keys` + `createDecryptor` 的 V5 分支 |
| 3 | 无任何公式求值能力（`xlsx.js` 只读出 `f` 字段当文本） | 新增 `formula.js`（子集求值器）；`office_read` 加 `recalc` 参数（默认 false）+ `stats.recalc` |
| 4b | `office_edit` 对 `.html` **本来就通**（`resolveKind` 把它判成 `text` → `editTextFile`），但文档写着"不支持" | 补回归用例 + 三份文档口径改成"文件级文本替换" |
| 4c | `writeDocx` 无 `w:docGrid` | `docx.js` 新增 `docGridXml` / `normalizeGrid`；`office_create` 加 `grid` 参数 |

### 移植策略（只搬目标能力，不整体覆盖）

- **独立模块直接对齐**：`image.js`（新）、`formula.js`（新）；`pdfcrypt.js` / `docx.js` 复制 WB 版后
  用 `git diff --no-index` 复核过：**差异全部是本批目标改动**，没有 DSH 侧独有内容被覆盖。
- **`pdf.js` 只加图片链路**：WB 同批还把字体管线整体换成了 `ttf.js`（`resolveFontChain` / `buildSubset`）；
  DSH 侧继续用第五轮的 `pdffont.js` + `FontChain` + `opts.info` 报告契约 ——
  不把未审计的字体重写带进 DSH，也不改变第五轮产出质量门的输入契约。
- **`html.js` 只改 image 分支**：从占位文本改成真 `<figure><img src="…" alt="…">`，
  保留 DSH 第五轮 `documentToHtml` 的其余行为（`escapeHtml` 仍是 export、`page-break` class 不变）。
- **`index.js` 逐处接线**（导入 / `editDocx` / `recalc` / `grid` / `.html` 编辑 / notice 传递），
  不用 `git apply`（第三轮教训：`a/`≠`b/` 时它会把文件写到 `b/` 路径并把原文件移走）。

### 本批新增的一条 DSH 专属硬兜底：`degradeImageBlocks()`

输入侧现在会产出 image 块，而 `.docx` / `.odt` / `.pptx` / `.md` 写出端**还没有** image 分支
（`docx.js` 的 `writeDocx`/`docBlocksToXml` image 分支被本任务的范围隔离明确排除，属另一批次）。
把 image 块原样交出去会落进 `blockToXml` 的 default 分支 → **什么都不输出 = 静默丢图**，
直接违反"不静默"红线（WB 侧第十二轮没处理这一跳）。

`index.js` 新增 `degradeImageBlocks(model, targetExt)`：非 `pdf`/`html` 目标上把 image 块降级为
**字面文本** `![alt](path)`（信息不丢），并写进 `stats.imageFallback` + `notice`。

### 实现里两个真坑（移植自带，已确认识别/修复状态）

1. **`COUNTIF` 的自匹配**：`flatArgs` 会把所有参数摊平 → 判据字符串自己成了"总是匹配自己"的一项
   → `COUNTIF` 结果永远 +1。WB 侧已修（`formula.js` 的判据与区域分开处理），本批原样带入。
2. **`ROUND` 的重复乘方**：把已经算好的 `10^d` 又乘了一次 `10^d` → `ROUND(1.0666,2)` 返回 1.0666。
   同样已由 WB 侧修好；`R12-3` 用例用 `ROUND(AVERAGE(C2:C4),2) = 1.07` 把它钉死。

### 实测对照（本机，工作副本；用例前缀 `R12-`）

| 用例 | 实测结果 |
| --- | --- |
| `R12-0` 包结构体检（新增 4 条） | docx 合规；把病灶内容类型塞回去**必须**被判不合规（负向控制）；`ContentType` 写在 `Extension` 前面的属性顺序不得误报 |
| `R12-1` md 图片行 → image 块 | `heading,image,paragraph`（旧版 `heading,paragraph`，图片只剩 `!` + 链接） |
| `R12-1` md → html | `<figure><img src="…r12-img-rgb.png" alt="红色方块"></figure>`（旧版是 `<p><em>[图片：…]</em></p>`） |
| `R12-1` html → md | 往返仍是 `![红色方块](…r12-img-rgb.png)` |
| `R12-1` docx 插图 2 张 | `word/media/image1.png` + `image2.png`；`rId` 唯一且 `= 插前 + 2`；`w:drawing` 在锚点段落之后；`[Content_Types].xml` 补了 `Default Extension="png"`；插图后包结构体检通过；读回见 2 个 `[图片]` |
| `R12-1` 含图 PDF | `/XObject` + 恰好 2 个 ` Do ` + `/SMask`；`imagesSkipped` 记下 `missing.png` 与原因；含图 PDF 正文仍可读 |
| `R12-1` convert → pdf | `stats.pdfQuality.images === 1` |
| `R12-1` md → docx 含图 | `stats.imageFallback` 1 条 + `notice` 明说"不内嵌"；docx 正文里保留字面 `![红色方块](…)` |
| `R12-1` 含图 HTML → docx | `word/media/` 为空（**范围隔离**：writeDocx 的 image 分支不归本批），但 `notice`/`stats` 有账 |
| `R12-2` AES-256 R5/R6 | 空口令透明解密读出正文，原文件 size/mtime **不变** |
| `R12-2` 失败路径 | 真口令 → "需要打开密码"+ 下一步；`/Perms` 已知明文破坏 → 同样报需要密码（**绝不接受错误密钥**）；截断文件 → 四要素错误或"无文本层"，不抛裸异常 |
| `R12-3` 公式重算 | `SUM=240`、`ROUND(AVERAGE,2)=1.07`、`COUNT=3`、`COUNTIF(>75)=2`、`IF→达标`、`VLOOKUP` 命中 `80` / 未命中 `#N/A`、跨表 `240`、`B5&"元"→240元`、`1/0→#DIV/0!`、`CONCATENATE→unsupported`（保留原值） |
| `R12-3` 兼容性 | 不设 `recalc` 时 `stats.recalc` 不出现、正文里公式原样是 `=SUM(B2:B4)`；非表格文件只记 `{skipped}` |
| `R12-4` 稿纸网格 | `grid="20x25"` → `w:linePitch="558"` / `w:charSpace="211"`；非 docx 与非法取值都抛 `【稿纸网格】` 四要素式错误；包结构仍合规；读回内容不受影响 |
| `R12-4b` `.html` 编辑 | `office_edit` 替换文本成功且 `<p>` 结构原样保留（`format` 仍报 `text`，与旧口径一致） |

### 独立交叉验证：AES-256 夹具的"自证风险"与本次尝试的结论

夹具与解析器**同源**（都用 `pdfcrypt.js` 的 `hash2B`），所以 `R12-2` 证明的是"读取链路端到端可用
+ 失败路径不猜"，**不是**"与第三方实现逐字节一致"。本轮尝试用本机可得的独立实现交叉验证：

| 探针 | 结果 |
| --- | --- |
| Windows WinRT `PdfDocument`（即 `pdf-render.ps1` 用的那套，微软实现）打开**明文** PDF `r12-img.pdf` | `OPEN_OK pages=1`（口令参数传 `''`/`'secret'`/`'wrong'` 都 OK）⇒ **验证脚本本身有效** |
| 同一脚本打开自造 `r12-aes256-r5.pdf` / `r12-aes256-r6.pdf` / `-locked.pdf` | **全部 `OPEN_FAIL`**（`AggregateException`；口令 `''`/`'secret'`/`'wrong'` 全试过） |
| Edge/Chrome（PDFium）无头 `--print-to-pdf` | 无输出（无头模式不渲染本地 PDF 文件，此路不通） |
| 扫描本机真实样本找对照（`~/.dsh/attachments` 等 33 个 PDF） | **未找到任何 R5/R6 加密样本**（唯一的 `/Encrypt` 命中是正文里字符串的误匹配） |

**结论（如实记录，不过度推论）**：本机没有支持 AES-256 的第三方实现可用于交叉验证，因此这条
自证风险**仍然存在**。两种解释**尚未区分开**：① Windows 的 WinRT PDF 组件不支持 R5/R6
（明文对照通过说明调用方式无误，故这条概率更高）；② 夹具的 V5 字典仍有个别规范细节偏差。
**后续判定方法**（留给下一轮）：拿一份真实世界的 R5/R6 加密 PDF（或用支持 AES-256 的第三方工具
生成一份）跑 `office_read` —— 能读出正文 ⇒ 夹具与实现都对；读不出而第三方能开 ⇒ 立刻定位到实现。
在此之前，`R12-2` 的结论只应读作"**DSH 侧读取链路自洽、失败路径不猜**"。
探针脚本留档在工作副本上级目录：`work\office-enhance\aes-xcheck.ps1`（WinRT 带口令重载）与
`scan-enc.mjs`（加密标记扫描），均为只读探针。

### 基准

- `node test.mjs`（插件目录，完整模式）：**516 checks / ALL PASS**，退出码 0。
  轨迹：**467**（第五轮收口）→ 改完 `pdfcrypt/docx/pdf/html/index` 主体后先跑一次
  → **467 checks 仍 ALL PASS**（零回归）→ 加 48 条 `R12-` 用例 → **515** → 补 `.html` 编辑 1 条 → **516**。
  准确数量以运行输出的最后一行为准。
- `node repro.mjs`：🟢 GREEN（退出码 0）—— 返回边界哨兵在新增字段（`stats.imageFallback` /
  `stats.recalc` / `stats.pdfQuality.images*`）之后仍然全绿。

### 行为边界

- **既有参数语义一律不变**；新能力全是"只增参数 / 只增 stats / 只加 notice / opt-in（`recalc`、`grid`）"。
- **干净文件的 `as=meta/markdown/json` 与产出字节逐字不变**：`R12-0`/`R12-1`/`R12-3` 直接断言
  "不设 recalc 时公式原样保留"、"没有 image 块时 PDF 资源字典仍长成 4 个 `>`"。
  （`markdownToDocument` 对**含图** markdown 的块类型变化是本批需求本身，见下。）
- **不静默**：图片跳过 / 图片降级 / 锚点未命中 / 缓存 stale / 未支持函数 / 错误值 —— 全部同时进
  `stats` 与 `notice`（或 `recalc.details` 的单元格级 error）。
- **绝不猜**：AES-256 的候选密钥必须过 `/Perms` 校验才被采纳；公式里不支持的函数标 `unsupported`
  并**保留原值**；口令候选只有调用方给的那组（默认空串）。
- 零新增第三方依赖：`image.js` / `formula.js` / 图片 FlateDecode 只用 `node:zlib`；
  AES-256 只用 `node:crypto`；PNG 解码复用自家 `png.js`。
- 未碰 WorkBuddy 目录（只读取其实现做移植源）、未碰 `~/.dsh/cordis.patch.yml`、
  未碰宿主装配与 `dsh-badge` 合同。

### 红线偏离：**无**

本批没有动质量门判据、没有打破"干净文件逐字不变"（既有 467 条用例一条不改、全部保持通过）、
没有改 `cordis.patch.yml`、没有引入新依赖。**唯一与任务约定的冲突**记在下面（按"先报告"处理）。

### ⚠ 验收项与范围隔离的冲突（如实记录）

提示词里需求 1 的验收写着"含图 HTML → `office_convert` → docx 的 `word/media/` 真有图片"，
但同一条提示词的「范围隔离」又明确写着 **`docx.js` 的 `writeDocx`/`docBlocksToXml` 增加 image 块分支
由另一会话负责、本任务禁止触碰**。两者互斥，处置：**不碰 `docx.js` 的写出分支**，改为在
`index.js` 里显式降级 + 记账（`stats.imageFallback` / `notice`），并用
`R12-1 降级：含图 HTML → docx 时 word/media 仍为空…` 把现状钉死 ——
该验收项要真正达成，需等负责 `writeDocx` image 分支的那一批落地（届时这条用例应改成
断言 `word/media/` 非空，`degradeImageBlocks` 的 `IMAGE_EMBED_TARGETS` 里加上 `docx` 即可）。

### 本批新增的 DSH↔WorkBuddy 分叉（只声明，不回移）

| 文件 | DSH 侧有什么而 WB 侧没有 | 性质 |
| --- | --- | --- |
| `index.js` | `degradeImageBlocks()` / `IMAGE_EMBED_TARGETS` + `office_create`/`office_convert` 里的 `imageFallback` 记账 | DSH 侧兜底（治"image 块在无 image 分支的写出端静默丢图"）；**建议回移** |
| `index.js` | `pdfOutputGate` 的 `withImages()`（`pdfQuality.images` / `imagesSkipped` + notice），并接进 `office_create`/`office_convert` 的返回 | DSH 侧的产出质量门与 WB 侧（`writePdfWithGate` + `writePdfReport`）实现不同，属各自的记账面 |
| `test.mjs` | `R12-0/R12-1/R12-2/R12-3/R12-4/R12-4b` 六组 48 条 + `checkOoxmlPackage()` + `makeTestPng()` + AES-256 自造夹具 | 仅测试 |
| `png.js` | 导出 `decodePng()`（供 `image.js` 复用），**同时保留** DSH 第五轮的 `pngInkCoverage()`（WB 侧没有这个函数） | DSH 侧多一个消费方 |

⚠ 反向分叉（WB 有、DSH 没有，本批**不移植**）：`ttf.js`（WB 第十二轮的字体子集重写）、
`api.mjs` / `doctor.mjs`（CLI）、GPU OCR 链路（Python/DirectML）、
`ocrEngineMode` / `engineDecision` 等引擎角色观测面。是否回移由主人决定。

## 2026-09-23 — 第五轮：中文 PDF 嵌字体子集 / PDF 产出质量门 / Word 拒开根因 / .html 写出端

> 对象：本目录（`${DSH_HOME}/plugins/dsh-office\`）。四条已知缺陷 + 一条顺带发现。
> 全程单变量、可复现：每个根因都落到"改哪一处才变绿"，并在真机 Word 16.0 / PowerPoint 16.0 上亲验。

### 1. 中文 PDF 不嵌字体 → 内嵌真字形（`pdffont.js` 新增）

- **症状**：WinRT 渲染空白 + 问号、Edge/Chrome 乱码、`office_read` 提取 `☆`/`①` 变成 `?`；
  字节级证据是 `/FontFile=0`、`/ToUnicode=0`（只有 `STSong-Light` + Adobe-GB1 预定义 CMap）。
- **改法**：新增 `pdffont.js`（零依赖 TTF/TTC 解析：`simsun.ttc` 集合取 font[0]、cmap 0/4/6/12、
  hmtx/metrics、**恒等 GID 子集**——不重排 GID，未用字形 `glyf` 清零 + `loca` 重建长格式 + 重算
  `checkSumAdjustment`；字体链 cjk `simsun→simhei→msyh` / emoji `seguiemj` / symbol `seguisym`）。
  `pdf.js` 的 `writePdf` 改为 `Identity-H + FontFile2 + 完整 ToUnicode CMap + /W + CIDFontType2`，
  `BaseFont` 带 `XXXXXX+` 前缀；`segmentText` 分段（latin 走 base-14 / embed 走 Identity-H /
  `DSH_OFFICE_PDF_EMBED_CJK=0` 回退旧 STSong 路径）。测量与绘制共用同一分段，`widthOf ≡ drawText`。
- **顺带的体积根因**：子集化原本整表拷贝 `COLR`/`CPAL`/`SVG `——PDF 的 CIDFontType2 只吃 `glyf` 轮廓，
  这三张表渲染器根本不读，而 Segoe UI Emoji 的 `COLR` 单项就有 7.4 MB。加入 DROP_TABLES 后：
  **9000 字 / 445 字形样本 4.35 MB → 0.39 MB**（小样本 4.33 MB → 0.18 MB），渲染与提取逐字不变。

### 2. 自己产出的坏 PDF 一路绿灯 → 出站质量门 `pdfOutputGate()`

- **两级判据**（不 OCR、不消耗视觉额度）：①字节级——正文含 CJK 却无 `/FontFile*` → hard-fail
  （`DSH_OFFICE_PDF_EMBED_CJK=0` 的显式回退同样拦）；②渲染级——bytes 先写**同目录临时文件** →
  `pdf-render.ps1` 渲染第 1 页 → 判"渲染为空白"。
- **修完自己踩的坑（两处，均已修）**：
  - 渲染通过后**忘了把临时文件改名到目标**（`office_create` 报 `pdfQuality.renderCheck=pass`，
    但目标文件根本不存在，套件里表现为"创建 pdf 含中文 FAIL + 之后 `文件不存在` 的 HARNESS ERROR"）；
  - "空白"判据最初用**覆盖率**（`ink < 0.0002` 或 PNG<30KB 且 `ink<0.001`），会把合法近空白文档误杀：
    `document:{blocks:[{type:'hr'}]}` 渲染出 8668 字节 PNG、0.058% 墨迹 → 被拒。
    现改为**绝对黑像素数**（`pngInkCoverage` 增加返回 `dark`）：整页 `dark < 40` 才判空白；
    辅助判据要求"可见文本 ≥ 24 字"；并把 `hr`（`---`）与 `<!-- pagebreak -->` 排除出"可见文本"。
    实测：真正的中文空白探针 0 px → 拦下；hr-only → 放行。
- **通过才原子改名到目标；失败删半成品 + 四要素错误 + 逃生通道**（md → odt → Word COM
  `ExportAsFixedFormat($out,17)`）。`DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` → `renderCheck="skipped"`
  （仍写字节级结论）。`office_create`/`office_convert` 返回 `stats.pdfQuality`。

### 3. `office_create` 的 docx 被 Word 16.0 拒开（`0x800A1401`）→ 一行 Content-Type

- **先推翻的错误假设**：子代理给的"`styles.xml` 里 `pPr` 顺序违规"经 12 个单变量变体证伪
  （改 styles 顺序 / 删 styles 部件 / 最小正文全部仍 FAIL）。
- **根因**：`[Content_Types].xml` 里 `docProps/app.xml` 的 Override 内容类型写成了
  `application/vnd.openxmlformats-officedocument.**wordprocessingml**.extended-properties+xml`；
  OPC 规定 `app.xml` 是**包级**部件，正确值是
  `application/vnd.openxmlformats-officedocument.extended-properties+xml`（不带 `wordprocessingml` 段）。
- **决定性证据**（`docx-diag/variants2.mjs` + Word COM 逐个开）：
  | 变体 | 内容 | 结果 |
  | --- | --- | --- |
  | v2-01 基线 | — | `OPEN_FAIL 0x800A1401` |
  | v2-02 | **只把 app.xml 的 Override 类型改对** | **`OPEN_OK`** |
  | v2-04 | 只删 app.xml 部件（Override+rels 留着） | `OPEN_OK` |
  | v2-08 | 只把 app.xml **内容**换成 Word 自产的（类型仍错） | `OPEN_FAIL` |
  | v2-03 | 只删 Override（部件留着） | `OPEN_FAIL` |
  | v2-05/v2-07/v2-09 | 只删 core.xml / core 换最小内容 / app 换空壳 | 全 `OPEN_FAIL` |
  ⇒ 内容无辜、类型是根因。`pptx.js` 里 `presentationml.extended-properties+xml` 同类错误一并改对。
- **复验**：新产出的 docx → Word COM `OPEN_OK paragraphs=14 chars=81`（无"发现不可读内容"提示）→
  `ExportAsFixedFormat` 导出 224,583 字节 PDF → `office_read` 回读与修复前（同模型、只回退那一行
  Content-Types 的模拟包）**逐字一致**。

### 4. `test.mjs`：原生 Office 开箱冒烟（新增 3 条）

- Windows 上 `office_create` 出 docx/.html → `powershell.exe -Command` 驱动 Word COM：先
  `Stop-Process WINWORD` + 删 `%APPDATA%\Microsoft\Word\*.asd`，`DisplayAlerts=0`，**不调 `$word.Quit()`**，
  结束 `Stop-Process`；路径只走环境变量（避免中文/引号被命令行切碎），PS 片段 ASCII-only。
- 断言：`Documents.Open` 成功（不再 0x800A1401）、`ExportAsFixedFormat($out,17)` 产出 >2 KB PDF、
  Word 能打开产出的 .html。Word 不可用（COM 工厂未注册）时按套件惯例 `ok(…, true, '跳过')`。

### 5. `.html` 写出端（`office_create` / `office_convert` 新目标）

- `html.js` 新增 `documentToHtml()`：语义化 HTML5（`h1`-`h6`、`<p>`、嵌套 `ul`/`ol`、`<table>` +
  `<thead><th>`、`<blockquote>`、`<pre><code>`、`<hr>`、行内 `strong`/`em`/`u`/`s`/`code`/`a`），
  UTF-8 + `<meta charset="utf-8">`，**中文与 emoji 原样输出、不转数字实体**；接进 `WRITERS` /
  `KIND_TARGET`（document）/ 两个工具 description。
- **顺带修读取端**：`<pre>` 原本只收集**直接**文本子节点，导致语义化写法 `<pre><code>…</code></pre>`
  读出**空代码块**（实测往返丢正文）；改为 `preTextOf()` 递归取子树文本（含嵌套 `code`/`span` 与 `<br>`）。
- **往返实测**：md → .html → md 与 md → md **逐字一致**（含嵌套列表层级、引用、表格、代码块里的
  `<tag> & "quoted"`、行内 `code`）；产出的 .html 能被 Word 16.0 打开（`OPEN_OK`）。

### ⚠ 本轮发现但未解决：产出的 `.pptx` 被 PowerPoint 16.0 拒开

`Presentations.Open` 报 `0x80070570 The file or directory is corrupted and unreadable`
（PowerPoint 自产的对照包 `OPEN_OK`，所以不是 COM/权限问题）。已排除：

- **zip 层**：把 PowerPoint 自产 pptx 用我们的 `makeZip` 原样重打包 → 仍能打开；
- **OPC 不变量**：无缺失部件、无悬空关系目标、全部 XML 可解析、所有 `r:id` 都能解析、无重复 Id/Override；
- **非必需部件**：逐个删掉 app.xml / tableStyles / viewProps / presProps / notesMaster / core.xml → 仍失败；
- **叶子部件本身**：把我们的 theme / presProps / viewProps / tableStyles / app.xml / core.xml，以及
  `ppt/slides/slide1.xml` + 其 rels 换进 PowerPoint 自产包 → **全都开得开**。

⇒ 元凶在 `ppt/presentation.xml` / `slideMasters` / `slideLayouts` / `notesMasters` 及其组合之内，未定位。
本轮已按规范修好两处（`docProps/app.xml` 的 Override 类型、`aRuns` 里 `<a:solidFill>` 写成 `<a:rPr>`
兄弟的 DrawingML 违规——`a:r` 的子节点只能是 `(rPr?, t)`），但**不足以**让 PowerPoint 接受。
`test.mjs` 只用插件自带解析器回读 pptx，所以套件全绿掩盖了这条；已在 `SKILL.md`「已知边界」与
README 如实登记，需要 pptx 时先转 `.docx`/`.md` 或用 PowerPoint 模板另存。

## 2026-09-23 — 挂载方式改为裸包名：插件列表显示友好名

> 对象：`~/.dsh/cordis.patch.yml` + 本目录 `package.json` / 新增 `locale/`。起因：设置 → 内置插件 →
> 插件列表里 `tool-office` / `skill-office` 两行的标题是一长串 `file:///…/index.js?v=12`，看不出是哪个插件。

- **根因**：harness 的 `readPluginMeta()` 第一行就是
  `if (barePackageName(specifier) === void 0) return void 0;`，而 `barePackageName()` 对**任何含 `:` 的
  specifier**（`file://` 正好命中）直接返回 undefined → 条目拿不到 `meta`，UI 回落到
  `moduleShortName(moduleName)`，该函数只剥 `@scope/`、`cordis:`、`cordis-plugin-`、`dsh-` 这些**开头**
  前缀，对 `file://…` 一个都不匹配 → 原样显示整个 URL（连 `?v=12` 一起）。
- **改法**：`cordis.patch.yml` 的两行改用裸包名 `dsh-office` / `dsh-office/skill`；各 profile 的
  `node_modules/dsh-office` 链接到本目录；`package.json` 补 `exports`（`.`、`./skill`、`./package.json`、
  `./locale/en.json`、`./locale/zh.json`，外加 `./skill/…` 别名——`readPluginMeta()` 是按
  `${specifier}/package.json` 与 `${specifier}/locale/<lang>.json` 拼路径解析的，子路径条目必须有别名才解析得到）。
- **友好名**：新增 `locale/en.json` + `locale/zh.json`（结构 `{"meta":{"title","description"}}`）。有字典时
  `localizedText()` 返回多语言**对象**，UI 走 `resolveText()` 直接按语言取值，**不再**经过会剥掉 `dsh-`
  前缀的 `moduleShortName()` → 中文界面显示 "dsh-office（办公文档）"，英文界面 "dsh-office"。
- **代价**：裸包名不能带 `?v=N`，**热换能力消失**——改任何插件源码（含 `html.js` 等兄弟模块）都要重启
  profile 才生效；原先"改 index.js 后把 `?v` 加一"的用法作废。
- **验证**：web / headless 两个 profile 的 `dsh --profile <p> --dump-config` 均 exit=0 且两行 name 为裸包名；
  直接调用 `readPluginMeta("dsh-office", <profile>/package.json)` 返回
  `{"title":{"en":"dsh-office","zh":"dsh-office（办公文档）"},"description":{…}}`；
  `import("dsh-office")` / `import("dsh-office/skill")` 均导出 `name`/`apply`/`inject`；
  改后热重载下 `office_read` 仍正常返回。
- **回滚**：备份在 `~/.dsh/backup-office-friendly-20260923-151128/`（含改前的 `cordis.patch.yml`、
  本目录 `package.json`、两个 profile 的 `package.json`）。

## 2026-09-16 — 第四轮补丁：读取链路修补（HTML / 批量盘点 / 批间衔接 / 字符级质量启发 / CSV 编码）

> 对象：`<work>/deepseek9\work\office-read-fix\plugin\`（可写工作副本），跑绿后一次性回写
> `${DSH_HOME}/plugins/dsh-office\`。需求来源：一次重度实战（25 个中文 PDF + 17 个 HTML
> 学习笔记）暴露的 5 个缺口，按 1、2 → 3、4 → 5 优先级全部交付。
> 全局约束遵守：不改变既有 as/meta/markdown/json 返回结构（只加增量字段）；零新增第三方依赖
> （HTML 用插件自研解析器，无许可证/体积负担）；每项带接口草案、错误处理与兼容性说明。

### 需求 1【高优】HTML 读取与转换 —— 新文件 `html.js`

- **实现路线**：DOM 解析优于正则（需求方实测正则剥标签丢嵌套列表语义）。零依赖三段式：
  `tokenize()`（词法：标签/注释/DOCTYPE/**原始文本元素** script·style·textarea·title·xmp）→
  `buildTree()`（语法：隐式闭标签 li·dt·dd·tr·td·th·option·optgroup·thead·tbody·tfoot、
  void 元素、`<p>` 的作用域关闭、大小写无关配对）→ `toDocumentModel()`（语义：块级元素冲刷
  段落、行内样式 run 合并、嵌套 ul/ol 压平进同一 list 块以 level 记层级、表格行收集兼容
  thead/tbody/tfoot 与游离 tr、blockquote→quote、pre→code、a→link run、img→image 块、
  dt/dd→缩进列表）。输出与 docx/odt/md 同构的 document 模型 → 全部 as 与全部转换目标自动继承。
- **接口**：`htmlToDocument(html) → {kind:'document', meta:{title}, blocks}`；
  `plainTextOf(doc)`（as="text" 的纯文本渲染：表格→制表符行，不残留 `#`/`|`/`**`）；
  `htmlPlainText(html)`。接入点：`loadModel()` 的 html/htm 分支（read/create-from/convert 共用）；
  office_read execute 的 `as="text"` 分支按 kind 分派（其他格式维持 modelToText 逐字不变）。
- **实体解码**：named 表 ~200 项（HTML4 Latin-1/符号/希腊/箭头/数学）+ 数字实体 + 浏览器同款
  **Windows-1252 数字别名**（`&#150;`→`–` 而非控制符）；越界/孤代理→U+FFFD。
- **剥离**：style/script/noscript/template/svg/canvas/iframe/表单控件整体跳过；
  `hidden` 属性、`display:none`/`visibility:hidden` 内联样式、`aria-hidden="true"` 子树跳过。
- **边界（有意不做）**：CSS 伪元素 `content:attr(...)` 不在 DOM、不包含；colspan/rowspan 只取
  单元格文本；`<ol start>` 起始编号 markdown 层无法表达；office_edit 不支持 .html。
- **热换**：`index.js` 以 `'./html.js?v=1'` 引入（与 cordis.patch.yml 同款 ?v= 爆破），改 html.js
  后 bump 该参数即可随 index.js 热生效，不必等宿主重启。

### 需求 2【高优】批量文件 stats 探测 —— `office_read` 新增 `paths`

- **接口**：`paths: string[]`（条目=文件或**目录**；目录展开受支持扩展名、排序、去重，
  **自动排除 `.ocr.md`/`.read.md`/`.ocr.json` 插件缓存**）；`path` 与 `paths` 二选一（同给报错，
  `paths: []` 显式报错）。返回 `{format:'scan', total, ok, failed, files:[…], skipped?, truncated?, notice, content}`。
- **行字段与 meta.stats 同名**：`format` / `pages`(PDF) / `characters` / `textLayerUsable` /
  `scannedPages`(数) / `garbledPages`(数) + `garbledPageSpec`（页码明细，仅乱码时）+
  `suggestedBatches`（建议分批数：乱码 PDF ⌈pages/20⌉、干净 PDF pages>12 时 ⌈pages/15⌉、
  文本类 ⌈chars/内联上限⌉）+ 行内 `error`（单文件失败不连累整批）。**绝不返回正文、不触发 OCR**
  （PDF 复用解析 memo；质量门纯 CPU）。`content` 是管道表（文件|格式|页|字符|文字层|扫描页|乱码页|建议批次）。
- **错误处理**：全部路径无效 → 四要素报错（含逐路径原因）；`as="markdown"` 与批量形态冲突 → 显式报错；
  单文件 >200MB/不可读 → 行内 error。上限 500 文件（超出记 `truncated`，不静默截断）。
- **实测**：`<corpus>/科目A\真题集` 23→21 个文件（排除 2 个
  sidecar 缓存后）一次调用 1.1-1.2 秒，逐行 stats 完整。

### 需求 3【中优】跨批读取连续性 —— 两个方案都交付

- **页级续读**：`pageFrom`/`pageTo`（integer，含端点）。与 `pages` **互斥**（同给报错）；
  `pageFrom` 缺省=1、`pageTo` 缺省=最后一页；`pageFrom>pageTo` 或越界报"页码范围"四要素错误。
  实现=readPdf 顶部合并成 `pages="N-M"`，因此质量门/OCR/缓存/分批逻辑零改动。
  `stats.pageFrom`/`stats.pageTo` 回显实际覆盖区间；`stats.nextPage` = 下一未读页（读完不出现）。
- **批间衔接**：凡 `pages` 指定且未覆盖全书，`stats.prevTail`（上一批末页末行 ≈100 字符，空白/扫描页不给）、
  `stats.nextHead`（下一批首页开头 ≈100 字符）+ notice 一行摘要（`boundary:true` 时省略——正文已内联）。
  `boundary: true`（boolean，默认 false）→ 以 `> 〔dsh-office 批间回看｜第 N 页末尾〕…` /
  `> 〔dsh-office 批间预览｜第 N 页开头〕…` **内联**进 markdown 正文首尾；`as="text"`/`as="json"`
  永不内联（原文/合法 JSON 不破坏）。**默认不改 content** —— 分批拼接与整本读取逐字 diff
  （验收用例：5 页书 `pages="1-2"`+`pages="3-5"` 剔标记拼接 ≡ 整本，368=368 字符一致）。

### 需求 4【中优】字符级启发 —— `textQuality(text, baseline, {structural})`

- **模式分离**（误报不升的关键）：`structural` 只在 **PDF 逐页**质量门启用（`textLayerProfile` 与
  readPdf 的 suspect 循环——两处同口径，meta 说的=read 遇到的）；整本口径（finishRead sidecar 兜底、
  非 PDF meta 探针、convert 出站探针）不启用——合法的重复表格行/短字段清单不会把整本推进兜底。
- **5 条启发**（原 4 条判据逐字保留，全部叠加）：
  1. 重复行率 ≥0.5（≥10 行"像正文"的行，≥4 字符含字母/CJK，点前导/纯数字行不算）；
  2. 单字/双字行率 ≥0.75（≥12 非空行，纯数字行不算）；
  3. Latin-1 高带（0xC0–0xFF，×÷ 除外）占可见 ≥0.5 —— UTF-8 按 Latin-1 解的 mojibake；
  4. 无元音字母串（[A-Za-z]{6,} 内 ≥8 连排不含 aeiou）占字母 ≥0.25 且字母 ≥60 —— 连续非词典字符；
  5. 非词典区相对异常：希腊/西里尔（0x0370-0x04FF）+制表几何（0x2500-0x25FF）+Latin-1 合计
     占比 ≥0.12 且全书基线 <0.03 且可见 ≥200 —— **俄文原文档 baseline 本身高，绝不误伤**。
- 触发页计入 `garbledPages`/`qualityGate.reasons` → `ocr="auto"` 自动送去重识别 + notes 给可复制
  的 `ocr="always"` 参数串（沿既有 suspect 管线，无新分支）。`stats.quality` 增量加 `oddCharRatio`。
- **验收**（test.mjs R4 组）：人为降质样本（Latin-1 高带/无元音串/重复行/单字行）全部检出；
  正常中文/英文/俄文/代码页/目录点前导页全部不误报；旧判据（PUA）回归通过。

### 需求 5【低优】CSV/文本写出编码与转义收敛

- **核实**：`WRITERS.csv` 原本就是 UTF-8 with BOM + RFC 4180 转义（逗号/引号/换行字段自动加引号、
  内部引号成对双写）——需求里"自动处理转义"已具备，本轮**参数化**并补齐真缺口：
- `office_create`/`office_convert` 新增 `encoding: "utf-8-sig"|"utf-8"`（仅 csv/tsv/md/txt 目标；
  默认 csv 带 BOM、其余不带=旧行为逐字不变；其他目标显式报错不静默忽略）。
  gb18030 写出因 Node Buffer 只内建 utf8/latin1 系、非零依赖不可达 → 维持不提供（读取侧
  decodeTextBytes 本就自动识别 GBK/Big5，不受影响）。
- **office_edit 修复**：decode→写回曾静默剥掉 BOM（"编辑一次、Excel 再打开就乱码"）。现在
  编辑 csv/tsv/md/txt 前探测 BOM、原样写回。

### 兼容性与验证

- **不变式**：干净文件的 `as="meta"/"markdown"/"json"` 返回值逐字节不变（R3 整本读取、R4 干净页
  等用例直接断言）；新字段全部为增量（stats.pageFrom/pageTo/nextPage/prevTail/nextHead、
  quality.oddCharRatio、scan 整体、html format）；`path` 参数 required 化解除但调用方全兼容
  （单文件调用照旧，都缺才报错且文案含 path/paths）。
- `node test.mjs` **431 checks / ALL PASS**（新增 R1-R5 五组 86 项；第三轮 361 项全部保持通过）；
  `node repro.mjs` 🟢 GREEN。真实样本：两份 68-73KB 学习笔记 HTML（tables 3/36、嵌套列表、
  标题层级）语义转换验证；23 个真题 PDF 批量盘点 1.1 秒。
- 工作副本：`work\office-read-fix\plugin\`（`vendor/` 为指向 live 的 junction，供本地引擎跑测试）。

## 2026-09-15 — 第三轮补丁：pdf.js 解析层三条先存缺陷 + 质量门点前导假阳性

> 对象：`<work>/deepseek9\work\pdfjs-fix\plugin\`（可写工作副本），跑绿后一次性回写
> `${DSH_HOME}/plugins/dsh-office\`。备份：`dsh-office.bak-20260915-185050`。
> 这是三轮里**改动面最大**的一次：`pdf.js` 与 `index.js` 同时动，且**故意打破了两条既有红线**
> （见文末「红线偏离」）。

### 病灶（三条，本轮亲自复核，非采信探路报告）

| # | 位置 | 缺陷 | 后果 |
| --- | --- | --- | --- |
| **A** | `pdf.js` `expandObjStms` | 头部索引 `nums` 从 `text.slice(0, first)` 读出（`off`/`end` 是**区域内相对偏移**），却被当成 `text` 的**绝对**下标 | 每个 ObjStm 内对象整体前错 `First` 字节。第一个对象恰好读到 **ObjStm 的头部索引表本身**（`"2 0 3 342 4 609 …"`）→ `/Pages` 解析不出 `/Kids` → `pages()` 落到"扫描全部 `/Type/Page`"兜底分支、**页序退化成对象号升序**；其余对象整体错位 → 页 `/Resources` 错位 → `fontsOf` 空表 → 2 字节 CID 被拆成两个单字节字符 = **"整本中文乱码"的真因** |
| **B** | `pdf.js` `mul()` | e/f 两行三个独立错误：读越界的 `b[6]`、该用 `a[4]/a[5]` 处写成 `b[4]/b[5]`、以及 `0 * undefined = NaN` | **连纯平移都炸**（`mul(ident,[1,0,0,1,5,7])` 的 e/f 都是 `NaN`；旧 e/f **在任何输入下都是 NaN**，不存在某个有限错值）→ 每个 run 的 y 都是 NaN → `Math.round(NaN/2.5)` 是合法 Map key → **每页所有 run 塌成一行** |
| **C** | `pdf.js` `runPara` | 换行循环把 `fit` 夹在 ≥1；一旦"本行剩余宽度装不下**一个**字符"，判定条件恒为真 | **同步死循环**，事件循环被占死（`Promise.race` 连超时都触发不了，**整个宿主进程冻住**）。触发条件**与字符种类无关** |

**A、B 必须同批**：只修 A 时"看着有换行"是乱码里混进行分隔符的假象，单独合入会让病灶样本
"看起来更坏"（换行 718 → 0）。

**C 的触发条件**（勘误，见下方「前提更正」第 1 条）：不是"私用区专属"。`_`.repeat(200)、汉字 50 个、
CJK 扩展 B、ASCII 长串**一样冻宿主**。

### 连带修复（同一批，否则修完反而暴露新问题）

| 项 | 位置 | 行为 |
| --- | --- | --- |
| `cm` 合成顺序 | `pdf.js` | 规范是 `CTM_new = M_cm × CTM_old`，旧版写成 `mul(a, ctm)`（反了，带缩放的嵌套 `cm` 会算错）。**注意：`pdf.js:712/719` 的 `mul(ctm, Tlm)` 本来就是对的**（在 `mul(A,B) ≡ apply(A, apply(B,·))` 语义下 `Tlm × CTM` 就该这么写）——真正的错在 `cm` 那一行 |
| **q/Q 图形状态栈** | `pdf.js` | 旧版 q/Q 只清操作数栈，CTM 整页单调漂移、永不复位 |
| `Td` / `T*` / `'` / `"` | `pdf.js` | 旧版 `Td` 把上一个 `Tlm` **整个丢掉**（连续 `Td` 的多行文本全落在同一点）；现在按规范 `Tlm_new = T_translate × Tlm_old` 累积，且位移发生在**文本空间**（旋转行也对） |
| `Tz / Tc / Tw / Ts / Tr` | `pdf.js` | 旧版按"未知算子"清栈，状态永远读不到（`Tr` **不是**"被分词但从不读取"，而是根本不在 `OPS_TEXT` 白名单里）。`3 Tr` 是**不可见**文字，不生效就会把水印/隐藏层当正文 |
| 有效字号 | `pdf.js` `pushRun` | 存**设备空间的有效字号**（`\|Trm\|` 纵向量长），不再存裸 `Tf`（某样本 209 vs 真实 10.45pt，宽度估算差 20 倍）。新增 `rawSize` 字段 |
| CJK 空格 | `pdf.js` | 行内 run 之间补空格改为 **CJK 感知**（两侧有一侧是 CJK 就不补），旧规则只对恰好以 `一` 开头的 run 特判 |
| `as="json"` 收口 | `index.js` | PDF 的 json 分支过去**忽略 `pages`**，且超 `READ_CAP` 时 `capText` 在**字符串中间**切断 → 返回**无法 parse 的半截 JSON**。页序修好之前碰不到，修好后立刻暴露。现在 honor `pages`，仍超限就**按页丢尾部**并写进 `stats.truncateNote`（永远是合法 JSON，绝不静默丢）。新增 `splitDocByPage()` / `capJsonPayload()` |
| **缓存迁移 `parser:` 版本戳** | `index.js` | `.ocr.md` 按页码记，而页码来自插件自己的 `pages()`；A 修好之后"页码 → 哪张纸"的映射变了（见下），旧缓存会把**另一张纸**的识别结果贴上这一页。新增 `PDF_PARSER_VERSION = 2`：`writeOcrCache()` 写 `parser:`（位置 `covered \| total \| parser \| src \| retry`），`readOcrCache()` 读到**缺失或不符 → 整份视为未覆盖**（`total` 保留），外显 `stats.ocrCacheStale` / `stats.ocrCacheNote` / `sidecarCoverage()` / `convertRefusalError()` |
| **质量门"点前导"假阳性** | `index.js` `textQuality()` | 见下节（本条是**红线偏离**） |

**渲染缓存目录不处理**：它按 `<basename>-<mtime>` 键控、与解析器版本无关，但它缓存的是**图像**，
不受文字层解析器影响 —— 判定为**无需迁移**（这是个判断，不是遗漏）。

### 页序迁移（A 的行为后果，必须单独记一笔）

① 的 `/Pages`（对象 2）住在 ObjStm 里，且它是该 ObjStm 的**第一个**对象（`off = 0`）——旧代码恰好
把 ObjStm 的头部索引表当成了它的主体。实测映射（`out/pagediff-s1.txt`）：

- 真身：`/Kids = [432, 6, 22, 25, 134, … 236, 433]/Count 35` —— **封面是 obj 432，第 1 页**；
- 旧代码枚举顺序：`[6, 22, 25, …, 236, 432, 433]` —— **封面跑到第 34 位**；
- 两者**集合完全相同**（不丢页、不重复），是**置换**：旧 `idx0..32` = 新 `idx1..33`，旧 `idx33` = 新 `idx0`；
- ② 的页序**完全没变**。

**危害是实打实的**：文字层走插件 `pages()`，OCR 渲染走 `pdf-render.ps1` 的 WinRT `GetPage(n-1)`
（**物理页序**）—— 旧代码下 ① 第 1–33 页的 OCR 成果**整体错一位**。已用真 OCR 交叉验证：
对 ① 请求 `pages="6-7"`，WinRT 渲染的物理第 6 页 OCR 文本 = 修后解析的第 6 页文本，逐段对齐。
这也是 `parser:` 版本戳必须存在的原因。

### 质量门"点前导"假阳性：**本轮唯一的显式红线偏离**

**症状**：① 修好 A/B 之后仍有 3 条 FAIL —— `textLayerUsable=false`、`garbledPages="4"`、
`qualityGate.reasons=["第 4 页：CJK 覆盖率 12.2% 远低于全书 73.8%（且非英文/代码页）"]`。

**诊断**：第 4 页是**目录页**（`# 目录 / # Contents CONTENTS`），整页由**点前导**（`..........1`）
+ 页码组成，所以逐页 CJK 覆盖率只有 **12.2%**，撞上"相对 CJK 覆盖率"判据。
`out/cjk-s1.txt` 显示：除 p4 外，有文字的页全在 **65%–85%** 之间，p4 是**唯一离群点**。

**这不是"报告难看"，是功能回归**：`ocr="auto"` 会把目录页送去白跑一次 OCR；
`office_convert` 会因为第 4 页"文字层不可信"而**拒绝转换整个 ①**（明明整本可读）。

**决策（三选一，取推荐方案）**：在 `textQuality()` 的可见字符循环里，把**连续 ≥3 个 ASCII 句点**
与 `\t\n\r`、空格同样处理 —— **不计入可见字符**。

**前后对照（`out/dot-decision.txt`，本机实测）**：

| 口径 | 全书 baseline cjkRatio | 阈值 baseline/3 | p4 visible | p4 cjkRatio | p4 garbled | 整本命中页 |
| --- | --- | --- | --- | --- | --- | --- |
| 旧（点号计入） | 73.8% | 24.6% | 1,833 | **12.2%** | **true** | **[4]** |
| 新（点号=装饰） | 79.1% | 26.4% | **281** | **79.7%** | **false** | **[]** |

p4 的点前导总长 **1,552**（1,833 − 281 = 1,552），全部是 `\.{3,}` 匹到的 ASCII 句点。

**反向护栏（新口径不得放过真乱码，全部实测）**：

| 输入 | 新口径结果 |
| --- | --- |
| 乱码夹具（PUA 成片 ×600） | `garbled=true`，`私用区码点占 100.0%` |
| 替换字符成片 ×200 | `garbled=true` |
| 控制符成片 ×200 | `garbled=true` |
| 非点号低 CJK 页（`¡¢£…¿`×20 + 中文基线） | `garbled=true`（相对判据**没被废掉**） |
| 纯点号页 ×400 | `visible=0`、`garbled=false`（点号是装饰，判"无文字"比判"乱码"更准确） |
| 点号 ×400 + 替换字符 ×60 | `garbled=true`，`替换字符占 96.8%`（装饰剔除不影响其他判据） |

**为什么这不是"放宽闸门"**：真缺 `ToUnicode` 的文档表现为 **PUA / 替换字符成片**，
不会长成"一片点号"；私用区 / 替换字符 / 控制符三条判据**一字未动**，且彼此独立命中。

**对照的侧证**：点前导修复前后，① `as="markdown"` 的正文差异已**逐行核对** ——
**唯一差异是第 4 页那两行质量门脚注（80 个汉字）及其附属的 2 个空行**
（`out/ab-diff2.txt` 行频差异 OLD x240/NEW x238 空行 + 那 2 行脚注，别无他物；
`out/ab-diff3.txt`：剔除脚注与空行后正文 **916 行逐字一致**）。
所以这条改动**只影响质量门判定，不碰一个字正文**。

### 实测（前后对照）

| 样本 | 指标 | 修前 | 修后 |
| --- | --- | --- | --- |
| ① `sample-A.pdf`（35 页） | 字符 / 控制符 / CJK / 填空段 / 含填空页 / 换行 | 45,692 / 19,085 / **0** / 0 / 0 / 718（假） | 23,992 / **0** / **16,141** / **305** / **29** / **1,084** |
| ① 的 `as="meta"` | `textLayerUsable` / `garbledPages` | **false** / `"4"` | **true** / `[]` |
| ② `sample-B.pdf`（43 页） | 字符 / CJK / 填空段 / 换行 | 31,200 / **19,879** / **700** / **0** | 31,200 / **19,879** / **700** / **1,738** |
| ③ 私用区 / 汉字 / 扩展 B / `_` / ASCII × 长度 1..200 | `writePdf` | ≥47 字符**永不返回**（同步死转） | 全部 **1–2 ms** 返回 |

① 无文字页（`scannedPages`）：修前 `[1,2,4,34,35]` → 修后 **`[1,2,3,5,35]`**（页序修正的直接后果）。

② 的"字符数"必须拆开看：**换行 0 → 1,738** 是修复本身（旧版每页塌成一行），
而**去换行后 29,400 → 29,420（+20）**，增量全是**行分隔 + markdown 标记**（`# ` 行 7 条）；
扣掉标记与换行后只多 **6** 个字符（`# ` 之后的空格与 run 间确有空隙时补的空格）。
**汉字 19,879 与填空 700 段逐字不变**。
（旧版基线 29,400 = total 29,442 − 换行 42，**用备份的旧 `pdf.js` 重新实测**得出。）
⚠ 这**故意打破了第二轮"干净文件逐字不变"的红线**，是修复不是回归。

口径说明（避免三份文档"打架"）：

- 换行 **1,738** = **逐页内**换行数（`probe.mjs`）；用例 `断行：② 换行总数…` 数的是逐页正文
  拼接后的换行，含 42 个页间连接符，所以它报 **1,780**。
- ① 的 **16,141** = 原始文本层（`probe.mjs`）；用例 `解析器：① 修后 CJK > 16,000` 数的是
  `as="markdown"` 落盘正文（含页标记与脚注），本轮实测 **16,311**。
- ② 的非换行字符 **29,400 → 29,420（+20）**：旧 29,442 字符含 **42** 个页间连接换行，
  新 31,200 字符含 **1,780** 个换行。旧基线用**备份的旧 `pdf.js`** 实测（`out/oldmod/`），
  不是从提示词抄的。

### 本轮前提更正（对照交接提示词 §1 / §4.4，**原提示词这几处是错的**）

1. **§1-E「死循环是私用区专属」错**。触发条件是 `runPara` 的换行循环把 `fit` 夹在 ≥1，
   一旦"本行剩余宽度装不下**一个**字符"判定恒真 → 同步死转。**与字符种类无关**：
   `_`.repeat(200)、汉字 50 个、CJK 扩展 B、ASCII 长串**一样冻宿主**。
   原提示词"普通汉字/扩展 B 均正常"不成立。
2. **§1-H「`pdf.js:712/719` 把 `mul(ctm, Tm)` 写反了，规范是 `Tfs × Tlm × CTM`，163.85 vs 536.4」错**。
   在 `mul(A,B) ≡ apply(A, apply(B,·))` 的语义下，`Tlm × CTM` 就该写成 `mul(ctm, Tlm)`，
   **712/719 本来就是对的**；真正写反的是 **691 行 `cm` 的 `mul(a, ctm)`**。
   另外旧 e/f 在**任何输入下都是 `NaN`**，**不存在 536.4 这个有限错值**。
3. **§1-H「`Tz/Tc/Tw/Ts` 被分词但从不读取」局部错**：它们确实读不到，但 `Tr` 也**不是**被忽略，
   而是根本不在 `OPS_TEXT` 白名单里 → 落进"未知算子"分支清栈。
4. **§4.4「② 字符数 29,442 不变**且**换行数从 0 变正」在数学上互斥**（换行也是字符）。
   实测拆解（用**备份的旧 `pdf.js`** 重新对账）：**CJK 19,879 与填空 700 段逐字不变**；
   旧 total 29,442 / 换行 42 → 非换行 **29,400**；新 total 31,200 / 换行 1,780 → 非换行 **29,420**；
   **差额 +20（不是提示词写的 +12）**，且全是行分隔 + `# ` 标记。
   ⚠ 提示词的 `+12`（`29,408 → 29,420`）来自 `test.mjs` 里一个**硬编码常量** `29442 - 34`，
   那不是实测值（真实旧基线是 `29442 - 42 = 29,400`）；本轮已把该断言改成
   **显式常量 `OLD_NO_NL = 29400` + 命名容差**，并在用例输出里打出真实口径。
   报告/文档一律按 **29,400 → 29,420（+20）** 的"带说明的偏差"记。
5. **新发现（连带）**：`office_create` / `office_convert` **生成**的中文 PDF 用的是**未内嵌**的
   `STSong-Light`，**WinRT 渲染为空白**（整页 OCR `code=101`）。所以乱码夹具的**可见文字必须用
   Helvetica**，PUA 只作不可见的乱码载体。已写进 `README.md`「已知边界」与 SKILL.md。

### 用例增删明细（`node test.mjs`）

**新增 33 条**：`解析器：` 3（`3 Tr` 不可见 / `Ts` 生效 / ObjStm 端到端）、`ObjStm：` 5
（夹具踩中路径 / 对象本体 / `/Resources` 字体表 / 越界不抛 / `/First` 缺失不抛）、`矩阵：` 6
（平移不 NaN / 恒等 / 不读 `b[6]` / "先 b 后 a" / Trm 合成顺序 / `cm` 顺序）、`死循环：` 3、
`解析器：` ① 修后 4（`textLayerUsable=true` / `garbledPages=[]` / `reasons=[]` / 无文字页 `[1,2,3,5,35]`）、
`解析器：`+`填空：` ① 正文 5、`断行：` 5、`缓存迁移：` 5 + 4、`填空：` 4。

**迁移 24 条**：原先把 ① 当"乱码样本"的用例全部改用**自造夹具** `buildGarbledFixture()`
（前缀 `乱码夹具：`），这是 A 修好之后的**必然结果** —— ① 的文字层完全可用，再拿它当"乱码样本"
只会让这段失去意义。其中 4 条"第 17 页换倍率救回"实测**在 ① 上依然成立**，于是迁到 `缓存迁移：`
保留（**删除 0 条**）。

**改写 5 处**（不算新增也不算删除）：两处 P0 manifest 正则加 ` \| parser: \d+`；
`样本 PDF：convert 落盘正文与页 1 文本一致` 改为"忽略空白与标题标记后比对"（断行/标题修好后
前缀不同是**修复的必然结果**）；测试里 6 处**手写 sidecar** 全部补 `parser: 2`。

### 本轮 8 条 FAIL 的处置（收尾记录）

| FAIL | 根因 | 处置 |
| --- | --- | --- |
| 1/2/3 `ObjStm：…` | **夹具写错了**：`minimalObjStmPdf()` 的循环 `for (i = 1..6)` 把故意不设的 obj 4 写成了 `4 0 obj\nundefined\nendobj`，于是文件里**真有一个 obj 4**，`scanObjects()` 先扫到它、`expandObjStms()` 又因 `objs.has()` 跳过 → 夹具**根本没走到被测路径** | 循环改为 `for (const i of [1,2,3,5,6])`（**测试问题，实现无 bug**） |
| 4 `` 解析器：`Ts` 升起量 `` | **测试写错了**：`Ts` 按规范**不随 BT/ET 重置**，第二条 `BT` 之后 `Ts` 仍是 100 → 两条 y 相同、被合理合并成 `"RS"` | 第二条显式 `0 Ts` 复位（**测试问题，实现无 bug**） |
| 5 `样本 PDF：convert 落盘正文与页 1 文本一致` | 旧版 ② 第 1 页塌成一行恰好让前缀相同；断行/标题修好后前缀自然不同 —— **修复的必然结果** | 改为忽略空白与 markdown 标记后比对前 60 字 |
| 6/7/8 ① 的质量门假阳性 | 目录页点前导 | 见上「点前导」节（**实现改动 + 红线偏离**） |
| —（修 FAIL-6..8 后新暴露的 1 条） | `缓存迁移：convert 拒绝文本解释"缓存为何不算数"` 原本拿 ① 当"乱码源"；① 变可读后 convert **不再拒绝** | 改用自造乱码夹具（3 页）当乱码源，断言不变（**测试迁移**） |

### 基准

- `node test.mjs`（插件目录，完整模式）：**378 checks / ALL PASS**，0 FAIL，退出码 0。
  轨迹：**258**（第一轮基线）→ **296**（自 WorkBuddy 同步的第二轮）→ **334**（DSH 专属 +38）
  → **378**（第三轮 **+44**）。本轮收尾**没有增减用例数**：8 条 FAIL 里 5 条是测试/夹具自身的错、
  1 条是测试迁移（①不再能当乱码源）、2 条是实现修复，`378 → 378`。
- `node repro.mjs`：🟢 GREEN（默认样本 ②，未改动），退出码 **0**；
  输出 `提取层：raw 文本层干净（0 个坏码点）` / `组装层：结构合法`。

### 红线偏离（**两条，都是"先报告再动手"**）

1. **动了质量门判据**（点前导不计入可见字符）—— 原提示词明列"不改质量门判据"。
   偏离理由、前后对照、反向护栏见上「点前导」节；`SKILL.md`「读取的正确顺序」第 3 步与「排查」
   两处都加了这条，供日后排查时先排除该假阳性。
2. **打破了第二轮"干净文件逐字不变"** —— ② 去换行字符 **+20**（原因逐项列出）。
   这不是回归：旧版"逐字不变"里包含"整页塌成一行"这个缺陷本身。

### 行为边界

- 既有参数语义一律不变；新行为全是"纯兜底 / 只增字段 / 只加文本 / opt-in"。
- **不静默**：替换 / 跳过 / 降级 / 回退 / 缺页 / 作废缓存 全部进 `stats` 与 `notice`。
- 干净文件（②）除"断行修好"这一项外逐字不变（不写 `stats.quality`、`stats.sanitized = 0`）。
- 不引入 Python / LibreOffice / pdftotext / npx；不改 `dsh-badge` 署名合同。
- 未碰 WorkBuddy 目录、未碰样本原件与附件目录（一律先 `copyFile` 到产物目录再操作）。
- **未给 `./pdf.js` 加 `?v=`**（源码分叉，范围外，先报告）。

### 本轮新增的 DSH↔WorkBuddy 分叉（回移前要逐个文件比，别整目录覆盖）

WorkBuddy 侧 `<home>/.workbuddy\skills\dsh-office\` **本轮一个字没改**
（最后写入 17:59，全在上一会话开始之前）。两侧 `pdf.js` 修改前**逐字节相同**
（sha256 同为 `706E1C4E446233C968C0C1C1CF67B0F68D6F446571DE29732E68085B752F6A96`），
所以**两个 bug 两侧都有**。本轮全部改动落在 DSH 侧，等于**新增分叉**：

| 文件 | 分叉内容 | 性质 |
| --- | --- | --- |
| `pdf.js` | A：`expandObjStms` 加 `First`；B：`mul()` e/f 修正 + 新增 `export function matrixMul()`、`OPS_TEXT_STATE`、q/Q CTM 栈、`Td`/`T*` 文本空间合成、`Tz/Tc/Tw/Ts/Tr` 读取、`3 Tr` 不可见跳过、`pushRun` 存有效字号（新增 `rawSize`）、`isCjkLike()` 与 CJK 感知补空格；C：`runPara` 换行循环防死循环 | **行为修复**，**WorkBuddy 侧同样有缺陷，建议回移** |
| `index.js` | `PDF_PARSER_VERSION=2` + `parser:` 版本戳 + `staleCacheNote()` + `readOcrCache`/`writeOcrCache`/`sidecarCoverage`/`convertSourceGuard`/`convertRefusalError`/`readPdf` 接线；`blanksPresentation()`/`renderBlanks()`/`renderBlanksDeep()`（任务 H）；`splitDocByPage()`/`capJsonPayload()`（`as=json` 收口）；`textQuality()` 点前导装饰 | DSH 侧新增能力，**建议回移** |
| `test.mjs` | 自造乱码夹具 `buildGarbledFixture()` + 33 条新用例 + 24 条迁移 + 6 处手写 sidecar 补 `parser:` | 仅测试 |
| `SKILL.md` / `README.md` | 两种乱码成因的判别表、"缓存迁移"段、填空呈现、死循环历史注记、第三轮修复注记、点前导假阳性排查条目 | 文档 |

⚠ **分叉是双向的**：WorkBuddy 侧另有 DSH 侧没有的东西（`api.mjs`、`repro-evidence.md`、`test-out/`），
且它的 `SKILL.md`(15,285) / `CHANGELOG.md`(16,310) / `test.mjs`(94,256) 与本侧尺寸不同。

---

## 2026-09-15 — DSH 专属补丁：P0 同步+接线 / 失败通道可观测性 / sidecar 全文

### 问题（DSH 宿主会话实测，编号 F1–F7）

被补丁对象 `${DSH_HOME}/plugins/dsh-office\`，病灶样本
`<samples>`（35 页 / 4,891,408 字节 /
sha256 前缀 `1b73dac1` / WPS 产物，30 页文字层整本 CID 乱码）：

| # | 症状 | 实测 |
| --- | --- | --- |
| F1 | **`office_read` 内容通道在 DSH 侧全灭**，与页、引擎、模式全部无关 | a–f 六条调用矩阵（整本三档 limit / 单页 text / 单页 markdown / 本地 OCR / 视觉 OCR / 纯扫描页）**全部**返回 `value is not lossless JSON`，只有 `as=meta` 幸存 |
| F2 | 工具报错时**副作用已发生**：sidecar 照样落盘，错误文本对此只字不提 | 宿主插件进程对"会话只读"的附件目录**有写权限**（WorkBuddy 侧"写失败"的经验在 DSH 不成立） |
| F3 | **批内失败页对调用方完全静默** | `ocr=always pages=5-24` 报错后 sidecar 实际 `covered: 1-16,18-24`，第 17 页缺页在返回值、错误里都看不到 |
| F4 | `office_convert` 零告警垃圾落盘复现 | 56,341 字节（NUL 5135 + C0 控制符 13666），`read` 判 binary 拒读，返回 JSON 只有 `{source,target,bytes}` |
| F5 | 宿主批处理下渲染缓存放点"看着消失" | `%TEMP%\dsh-office-ocr` 下按"当日修改 + ≥30 张 PNG"未定位到页渲染目录 |
| F6 | 填空下划线在 OCR 产物里不带任何字符 | 整本样本的填空呈"缺词"状态（本次人工标注 302 处才可用） |
| F7 | sidecar 污染只读附件目录 | `.ocr.md` 落在 sha256 命名的附件目录里，清理要专门越权 |

### 根因（逐条核对代码后）

1. **第一轮补丁的「返回边界」段在 DSH 版是死代码** —— `finalizeToolValue` 定义了但 `defineToolLite()`
   的 `execute` / `render` 都没调用，`catch` 也没过 `sanitizeThrown`。所以
   `stats.ocrCovered: undefined` 这类属性原样出站到 host 判定器 → **整条作废**。
   实测佐证：修复前的 `node repro.mjs` 在**干净对照样本**上就报 🔴
   `$.stats.ocrCovered 是 undefined` —— 与"某页乱码"无关，是全局结构问题（F1 的推论成立）。
2. 质量门 / 降级链整段在 DSH 版**完全不存在**（`textQuality` / `finishRead` / `writeReadSidecar` /
   `readSidecarPath` / `readErrorHint` / `readSuggestion` / `capWithOffset` 全无）。
3. `office_convert` 的 execute 链路没有 textQuality、没有 readOcrCache → 写盘前零校验（F4）。
4. `ocrPdfPages` 里被 `gateResult` 判不合格的页直接标 `failed`，无重试、且失败页不进任何对外字段（F3）。
5. 没有 `DSH_OFFICE_CACHE_DIR` → 成果只能落在源文件同目录（F7）。

### 补丁位置

| 文件 | 位置 | 内容 |
| --- | --- | --- |
| `index.js` | 整文件 | **以 WorkBuddy 版为源同步**（35 hunk / +715 −94：返回边界段、质量门+降级链段、第二轮五任务、全部接入点） |
| `index.js` | `defineToolLite()` | **接线（本轮关键）**：`execute` 返回走 `finalizeToolValue(await options.execute(…))`、`render` 投影同样过 `finalizeToolValue`、`catch` 走 `throw sanitizeThrown(e)` |
| `rapidocr.js` | `engineArgs(rawScale)` / `ocrImages(…,{scale})` | 倍率可显式传入。**同步源要求**：不给这个参数，换倍率重试渲染出 2× PNG 后引擎仍按默认 1024 长边缩回，重试等于空转（见下方"范围外改动"） |
| `index.js` | 新增 `sidecarCoverage()`，接进 `readErrorHint` | **DSH 补充（任务六-1 / F2）**：抛错分支一并给出 sidecar 路径 + `covered` + 缺页 + `grep 'covered:'` 提示 |
| `index.js` | `convertSourceGuard` 返回 `missing` / `covered` / `total` / `cachePath`；`convertRefusalError` 外显 | **DSH 补充（任务一-2）**：拒绝文本除四要素外给出 sidecar 现状；`下一步` 的 `pages` 只列**缺页** |
| `index.js` | `readPdf` 的 stats 新增 `ocrFailedPages`；正文脚注「批内缺页：…」 | **DSH 补充（任务三-3 / F3）**：每页失败带"页码 + 一句话原因"常驻外显，空数组 = 本批无缺页，调用方不必再 grep sidecar |
| `index.js` | `readPdf` / `office_read` 经临时字段 `__fullBody` → `finishRead` | **DSH 修正**：旧版把被内联护栏截过的**前缀**写进 sidecar，而 notice 承诺"整篇正文"。现在 sidecar 恒为全文，另加 `stats.sidecarChars` 可核对 |
| `index.js` | `readOcrCache()` 的 `src:` 正则收紧为 `[^|\n]+`、新增 `retry:` 回读；`ocrPdfPages()` 写盘前合并 retry | **宿主实测抓出的真 bug（同步源同样存在）**：manifest 各段是同一条注释行里用 ` \| ` 拼的，旧正则 `[^\n]+` 一路吃到行尾 → 第二轮新加的 `retry:` 被吞进 `src` 值 → 逐段匹配失败 → **整行 src 解析不出来**。后果是下一批写 sidecar 时 `srcOf` 为空，跨批的**引擎来源**与**倍率账**都被静默抹掉（实测：跑完 1-20 再跑 21-35，头部退化成 `src: rapidocr=21-35` 且 `retry:` 消失）。现在两段都跨批累积，且历史 retry 只进 manifest、不冒充本轮 `stats.ocrRetried` |
| `repro.mjs` | 新增（自同步源复制） | 返回边界常驻哨兵，退出码 0/1 = 绿/红 |
| `test.mjs` | 新增 4 个 DSH 测试段 | 接线 / F1 矩阵 / 批内缺页 / sidecar 提示 |
| `SKILL.md` / `README.md` / `CHANGELOG.md` | 文档同步 | 环境变量表新增 `DSH_OFFICE_OCR_RETRY_SCALES`、`DSH_OFFICE_CACHE_DIR` 两行 + 排查节新症状条目 |

### 用例名（在 `node test.mjs` 输出里可直接 grep）

- 接线：`接线：defineToolLite 的 execute / render / catch 三个出口都挂上了边界`、
  `接线：finalizeToolValue 与 sanitizeThrown 除定义处外确有调用点`、
  `接线：真实调用确实过了边界（stats.sanitized 只有 finalizeToolValue 会注入）`、
  `接线：catch 分支过了 sanitizeThrown（错误文本无游离代理 / 无控制符）`
- F1 矩阵：`F1矩阵：a1 as=markdown 整本 limit=200000 正常返回且无损` … `F1矩阵：f3 混合书中纯扫描页 ocr=always vision 正常返回且无损`（10 条）、
  `F1矩阵：纯扫描页 ocr=never 返回干净的"无文本层"说明（不是错误、不是乱码）`、
  `F1矩阵：整本无文字层 + ocr=never → 抛的是带四要素的说明（非裸 invalid output）`
- 批内缺页：`批内缺页：stats.ocrFailedPages 与 ocrFailed 同页、每页带原因`、
  `批内缺页：正文脚注直接列出缺哪几页 + sidecar 路径 + covered（不必再 grep）`
- sidecar 提示：`sidecar提示：convert 零缓存拒绝文本给出 sidecar 路径`、`…明说 covered: 无`、`…给出可照抄的 grep 提示`、
  `sidecar提示：部分缓存 → 拒绝文本报出 covered 区间`、`…点出仍缺的页`、`…下一步的 pages 只列缺页（不叫已覆盖的页重跑）`、
  `sidecar提示：部分缓存下 convert 仍然绝不产出目标文件`、`sidecar提示：非 PDF / 无 sidecar 时探测返回空串（绝不误报）`、
  `sidecar提示：抛错分支带出 sidecar 路径 + covered + 缺页（成果在盘上看得见）`、`sidecar提示：四要素仍然齐备（新增信息不挤掉旧契约）`、
  `sidecar提示：内联截断时 sidecar 仍拿到全文（sidecarChars > content 长度）`、`sidecar提示：sidecar 正文含护栏截掉的那部分（不重不漏，末尾完整）`

### 红→绿与宿主复验（DSH 专属，WorkBuddy 无法代验）

- `node repro.mjs`：同步前 🔴 RED（`$.stats.ocrCovered 是 undefined`，且**是在干净对照样本上**复现的）
  → 接线后 🟢 GREEN（退出码 0）。
- 宿主重放 F1 矩阵 a–f：热换后 6 条全部正常返回、零 `not lossless JSON`；
  纯扫描页 `ocr=never` 那条返回干净说明「（本页未提取到文本层，疑似扫描/图片页；可用 ocr:"always" 识别）」。
- `as=meta` 在病灶样本上给出 `textLayerUsable:false` / `garbledPages:"3,5-33"` / `qualityGate.reasons`
  / `suggestion.copy`（含 `ocr="always"`），且 `pages=35` / `pagesWithText=30` / `scannedPages=[1,2,4,34,35]`
  / `characters=45141` 与补丁前**逐字一致**（既有字段没改名没改值）。
- 第 17 页换倍率重试：`stats.ocrRetried=17`、`stats.ocrRetryScale={"17":1}`、
  `stats.ocrFailedPages` 为空、sidecar manifest `retry: 17=1` 且 `src` 仍记 `rapidocr`。

### 基准

- `node test.mjs`（插件目录，完整模式）：**334 checks ALL PASS**，0 FAIL。
  轨迹：258（提示词所载第一轮基线）→ 296（自 WorkBuddy 同步的第二轮）→ **334**（本轮 DSH 专属 +38）。
- `node repro.mjs`：🟢 GREEN（退出码 0）；`--pages "1-35"` 同样 GREEN。
- 慢段可跳过：`DSH_OFFICE_TEST_SKIP_BROKEN_OCR=1` 会跳过病灶样本的 OCR 段（其中含最慢的 8 条端到端用例）。
- 宿主侧复验：F1 矩阵 a–f 六条全部返回；`as=meta` 质量画像、`convert` 拒绝 / 转出、
  报错文本自带 sidecar 现状、换倍率重试救回第 17 页、manifest 跨批累积 —— 均在**运行中的宿主**上实测。

### 行为边界

- 既有参数语义一律不变；新行为全是"纯兜底 / 只增字段 / 只加文本 / opt-in"。- 不静默：替换 / 跳过 / 降级 / 回退 / **缺页**全部进 `stats` 与 `notice`。
- 干净文件逐字不变（不写 `stats.quality`、`stats.sanitized = 0`、不加脚注、不产新文件）。
- 重试**串行**（逐倍率一批），渲染次数计入现有渲染记账；`ocrEngine:"local"` 禁回落视觉的语义不变。
- 不引入 Python / LibreOffice / pdftotext / npx；不改 `dsh-badge` 署名合同（121×20、不换色、不改 Logo、不换链接）。
- 未改动 WorkBuddy 目录、未改动病灶样本原件。

### 范围外改动（先报告，再动手 —— 均已在会话中说明）

1. **`rapidocr.js`**：不在"只许改这六件"清单内，但第二轮任务三的换倍率重试**硬依赖**它新增的
   `engineArgs(rawScale)` / `ocrImages(…,{scale})`；不同步则重试渲染出的放大 PNG 会被引擎按默认
   1024 长边缩回，重试等于空转（第 17 页救不回来，验收 §3.4 第 4 条必挂）。改动为 5 行、
   向后兼容（不传参时读环境变量，与旧行为逐字一致），且 WorkBuddy 侧 CHANGELOG 本就把它列为
   第二轮补丁的一部分。
2. **`~/.dsh/cordis.patch.yml`**：把 `tool-office` 行的 `index.js?v=8` 提为 `?v=9`（后又提为 `?v=10`）。
   该文件自带注释即要求"编辑源码后 bump 版本爆破号"，否则运行中的宿主仍在跑旧模块，
   验收 §3.3 / §3.5 的宿主复验无从执行。**只改缓存爆破号，未动任何装配语义**（行、id、路径、顺序全不变）。
   注意：热换只重新导入 `index.js` 本身，它 `import` 的兄弟模块（含 `rapidocr.js`）走 Node ESM 缓存，
   **要等宿主重启才生效** —— 因此任务三的引擎侧倍率联动本轮只在 `node test.mjs`（新进程）中验证。

### F5 结论（任务六-3）：**改文档，不改代码**

实测宿主路径的渲染缓存写入 / 清理与既有描述**一致**，"找不到"是两处环境错觉，都不是 bug：

| 观察 | 数据 |
| --- | --- |
| 宿主插件写的缓存根 | `<TEMP>\dsh-office-ocr`（41 个目录） |
| 整本渲染目录**跑完仍在** | `broken-layer.pdf--1338544444` png=**35**；`____-____.pdf--1633024077` png=**36** |
| 倍率重试目录各自独立、也在 | `-s1` / `-s1.5` / `-s2` 各 png=1（与 `renderDirFor(file, tag)` 的设计一致） |
| 当日新建目录被删了几个 | **0 个**（>7 天线未到）——"批尾自清"的猜测不成立 |

两个错觉的来源：

1. **`%TEMP%` 语义不同**：会话的 `pwsh` 工具里 `$env:TEMP` 被改写成**每次进程一份**的
   `…\Temp\dsh-XXXXXX\`（本机实测 `dsh-KVdewp`），而宿主进程用的是**不带这层前缀**的用户 Temp。
   在 pwsh 里展开 `%TEMP%\dsh-office-ocr` 找的是 node 子进程那份，永远找不到宿主那份。
2. **中文 basename 被 sanitize**：`[^\w.-] → _` 且 JS 的 `\w` 只认 ASCII，
   所以 `sample-A.pdf` → `____-____.pdf-<mtimeMs|0>`，按名字搜必然落空。

处置：**代码不动**（合同本来就对），把这两点写进 `SKILL.md` 的渲染缓存条目 + 本节。

## 本轮前提更正（重要，影响第三轮的修法方向）

> **【第三轮结账 · 2026-09-15】本节是第二轮结束时写下的判断。为保留修订轨迹，原文一字不动，
> 只在下面标注"后来怎么变的"。** 第三轮已把本节列出的先存缺陷全部修掉，并**显式作废**一条前提：
>
> - ❌ **「① 文字层整本 CID 乱码（缺 `ToUnicode`）、整本必须 OCR」—— 显式作废。**
>   实测证明 ① **有完整文字层**；那段"乱码"是我方 `expandObjStms` off-by-`First` 造成的**假乱码**。
>   第三轮修好后：`textLayerUsable=true` / `garbledPages=[]` / CJK **16,141** / 控制符 **0** /
>   填空 **305 段 / 29 页** / 无文字页 `[1,2,3,5,35]`。**对 ① 不要再 OCR**
>   （见顶部「第三轮补丁」与 `README.md`「已知边界」）。验收口径同步改为"① 直接可读"。
> - ✅ 先存缺陷 ①（`expandObjStms` 少加 `First`）= 第三轮**任务 A**，已修；并连带落地
>   `PDF_PARSER_VERSION=2` 缓存迁移（修 A 会改变"页码 → 哪张纸"的映射，旧 `.ocr.md` 必须整份作废）。
> - ✅ 先存缺陷 ②（`mul()` e/f）= 第三轮**任务 B**，已修。但本节下方"规范合成应为 `Tfs × Tlm × CTM`：
>   `mul(flip, tm)` 正确值 `y=163.85`，现有写法给 NaN"这段**描述有误**，已在下文就地更正：
>   `pdf.js:712/719` 本来就对，真正写反的是 **`cm` 那一行**；旧 e/f **恒为 `NaN`**，
>   **不存在 163.85 / 536.4 这组有限值对照**。
> - ✅ 「附带发现 ②：私用区码点死循环」= 第三轮**任务 C**，已修（1..200 扫描全部 1–2 ms 返回）。
>   但"**私用区专属**"这个描述也要更正 —— 触发条件与字符种类**无关**。
> - ⏳ 「附带发现 ①：混合型乱码 PDF 的纯扫描页在 `convert` 里输出空白」**本轮未修**（不在本轮范围），
>   仍建议进下一轮：缓存里有正文的**无文字层页**应与 `garbledPages` 一并替换。

提示词 §1.1-1 / §1.3 把病灶样本判定为「WPS 产物、文字层**整本 CID 乱码**（缺 `ToUnicode`），**整本必须 OCR**」，
并把 ② 干净对照样本当作"既有健康行为"的基准。**这两个前提的因果都不成立**，本轮任务七探路 + 我方独立复现如下。

### 复现结果（只改 `pdf.js:161` 一行，未碰插件目录）

| 样本 | 版本 | 字符数 | 控制符 | CJK | `_` 段数 | 含填空页 | 每页换行数 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| ① 病灶 | shipped | 45,692 | **19,061** | **0** | 439(单字符) | 0 | 718（5 页为 0） |
| ① 病灶 | 仅修 ObjStm | 22,863 | **0** | **16,141** | **305** | **29** | **0（全塌成一行）** |
| ② 对照 | shipped | 29,442 | 0 | 19,879 | **700** | 43 | **0（全塌成一行）** |
| ② 对照 | 仅修 ObjStm | 29,442 | 0 | 19,879 | 700 | 43 | 0 |

修好后 ① 第 6 页原文可直接读出：`…习近平新时代中国特色社会主义思想的________。1982年召开的中共________，邓小平提出…`

### 两个先存缺陷（都在 `pdf.js`，均**不在**本轮"只许改六件"范围内）

1. **`pdf.js:161` `expandObjStms` 少加 `First`**：`nums` 是从 `text.slice(0, first)` 读的，故 `off`/`end`
   都是**区域内相对偏移**，却被当成 `text` 的绝对下标用 → 每个 ObjStm 内对象整体前错 `First` 字节。
   样本① 407 个对象里 **219 个来自 ObjStm**，其中页 `/Resources` 被解析成字符串 `"Object"`
   → `fontsOf`(`pdf.js:210`) 返回空表 → `decodePdfString`(`pdf.js:766`) 走非 Type0 分支，
   把 2 字节 CID 拆成两个 1 字节字符（`<0024>Tj` → `U+0000` + `$`）。
   **所以"整本乱码"是自家解析器坏了，不是文档缺 `ToUnicode`**；`<0042>`→ToUnicode→`U+005F _` 实测存在。
2. **`pdf.js:663-667` `mul()` 的 `e/f` 两行**：既读越界的 `b[6]`（`0*undefined → NaN`），又把 `a[4]/a[5]`
   误写成 `b[4]/b[5]`。实测 `mul(ident,[1,0,0,1,5,7])` 的 `e/f` 都是 NaN —— **连纯平移都坏**。
   后果：任何 `cm`/`Td`/`Tm` 之后 x/y 全 NaN，而 `Math.round(NaN/2.5)` 是合法 Map key
   → `runsToLines` 把整页所有 run 归到同一桶 → **每页塌成一行**。
   规范合成应为 `Tfs × Tlm × CTM`：`mul(flip, tm)` 正确值 `y=163.85`，现有写法给 NaN。

   > **更正（第三轮实测）**：这一段的**后半句是错的**，已作废。
   > `expandObjStms`/`mul` 修好后实测：`mul(A,B) ≡ apply(A, apply(B,·))`，
   > 所以 `Tlm × CTM` 就**该**写成 `mul(ctm, Tlm)` —— **`pdf.js:712/719` 本来就是对的**。
   > 真正写反的是 **`cm` 的 `mul(a, ctm)`**（规范 `CTM_new = M_cm × CTM_old`）。
   > 另外旧 e/f 在**任何输入下都是 `NaN`**，**不存在 536.4 或 163.85 这组有限错值**。
   > 「x/y 全 NaN → 每页塌成一行」的因果链则**完全成立**，是任务 B 的核心。

### 对第三轮的直接影响

- **不要只修一处**：上表显示仅修 ObjStm 会让 ① 的换行从 718 掉到 **0**（两个 bug 此前互相掩盖）。
  ① 现在"看着有换行"纯粹是乱码字符里混进了行分隔符。→ **两处修复必须同批发布**。
- **② 不能当"行为完好"的基准**：它的 43 页输出其实全是单行块。本轮把 ② 行为"逐字不变"当验收是对的
  （不能夹带改动），但它证明的是"没改坏"，不是"本来就是好的"。
- **本轮补丁仍然成立**：质量门判"文字层不可信 → 走 OCR"在现实现下是**正确且必要**的兜底（它救出了
  16,028 汉字）；修了解析器之后它退化为"真扫描件与真缺 ToUnicode 文档"的兜底，不会白做。
- **建议改写 F6**：从「填空下划线不带任何字符」改为「ObjStm 解析错位使文字层乱码 + `mul` 越界使坐标全 NaN，
  二者叠加让本来就是 `_` 字符的填空看起来消失」。已同步改进 `SKILL.md` 的已知边界。

### 附带发现（本轮未修，建议进第三轮）

**① `convert` 对"混合型乱码 PDF"仍会输出空白页**（实测，非缺陷而是规格边界）：
出站质量门按规格只替换 `garbledPages`（文字层存在但不可信的页）。纯扫描页（无文字层）不在
`garbledPages` 里，即使 `.ocr.md` 已经完整覆盖了它们，`convert` 依旧给这些页输出空正文 ——
宿主实测本样本第 1/2/4/34/35 页在转出的 md 里就是空的。第三轮建议：缓存里有正文的**无文字层页**
也一并替换（同一套 `readOcrCache`，只是把"替换范围"从 `garbledPages` 扩到 `garbled ∪ 已有缓存的空白页`）。

> **第三轮结账**：本轮**未修**（不在第三轮范围：A/B/C + 连带修复）。建议原样进下一轮。
> 注意第三轮修好 A 之后，① 的 `garbledPages` 变成 `[]`，这条的触发场景收窄为"**本身就有真乱码页
> 或真无文字层页**的混合型 PDF"，回归保护仍在（`乱码夹具：…` 一组用例）。

**② `office_create` → PDF 遇私用区码点会死循环卡死事件循环**（pre-existing，非本轮引入）：
（**第三轮已修**，见顶部「第三轮补丁」任务 C。但"私用区专属"要更正：触发条件与字符种类**无关**，
`_`.repeat(200) / 汉字 50 个 / CJK 扩展 B / ASCII 长串**一样冻宿主**。）
markdown 正文含约 ≥50 个 U+E000–U+F8FF 字符时，`writePdf` 永不返回，且因为它跑在事件循环上，
**整个宿主进程一起冻住**（`Promise.race` 超时都触发不了）。实测阈值：30 字符 / 40 字符正常（2–3 ms），
50 字符起挂；对照同类长度的一眉汉字（U+4E00）与 CJK 扩展 B（U+20000）均正常。
最小复现：`office_create path=x.pdf markdown="# x\n\n" + "\uE0A1".repeat(50)`。
根因在 `pdf.js`（不在本轮"只许改六件"清单内，故只报不修）。

---

## 2026-09-15 — 第二轮补丁（自 WorkBuddy 版同步落地）：出站质量门 / meta 质量画像 / 换倍率重试 / 缓存位置 / 截断协议

### 问题（四类静默失败）

病灶样本（35 页，WPS 产物，**30 页文字层整本 CID 乱码**，逐页控制字符占 32%~43%）：

| 调用 | 旧版结果 | 判定 |
| --- | --- | --- |
| `as=meta` | `pagesWithText=30 / scannedPages=[] / characters=45141 / fallback="none"` | ❌ **看起来完全健康**，实际整本必须 OCR |
| `convert → full.md` | 写出 56341 字节，其中 **NUL 5135 个 + C0 控制符 13666 个**；返回 JSON 只有 `{source,target,bytes}` | ❌ **静默坏输出**，零告警 |
| `read pages=17 ocr=always ocrEngine=local` | 「本地引擎未通过质量门，且 `ocrEngine:"local"` 禁止回落视觉桥」 | ❌ 该页正文丢失（同页 `RENDER_SCALE=2` 重跑 → 成功） |
| 分段续读 | `content` 尾部掺"已达上限"说明，消费方必须 `slice(0, nextOffset - offset)` | ⚠️ 协议脆弱，写错就"又重又漏" |

### 根因

1. `office_convert` 的 execute 只有 `hostPath → WRITERS 检查 → loadModel → adaptModel → WRITERS[ext]() → saveBuffer`，
   **没有 textQuality、没有 readOcrCache**。
2. `as=meta` 只统计"**有没有**文本层"，不判"文本层**能不能用**"。
3. `ocrPdfPages` 里被 `gateResult` 判不合格的页直接标 `failed`，**无任何重试**。
4. 截断说明与正文混在同一个 `content` 字段里。

### 补丁位置

| 文件 | 位置 | 内容 |
| --- | --- | --- |
| `index.js` | `office_convert.execute` + `convertSourceGuard` / `convertRefusalError` | 写盘前跑质量门：乱码页有完整 `.ocr.md` 覆盖 → 用 OCR 文本替换后继续（`stats.fallback="ocr"` / `garbledPages` / `notice`），否则**拒绝落盘**并抛四要素 |
| `index.js` | `textLayerProfile` / `garbledSpec`；`readPdf` 的 `as=meta` 分支 | 纯 CPU 质量画像：`textLayerUsable` / `garbledPages` / `qualityGate{testedPages,garbledPages,reasons≤3}`；命中乱码时 `suggestion` 改为 `{note, copy}` |
| `index.js` | `office_read.execute` 的非 PDF `as=meta` 分支 | 同样补 `textLayerUsable` |
| `index.js` | `ocrRetryScales` / `currentRenderScale` / `scaleTag`；`ocrPdfPages` 本地批次 | 未过门页按 `DSH_OFFICE_OCR_RETRY_SCALES`（默认 `2,1.5,1`）**串行**换倍率重试；记账 `stats.ocrRetried` / `ocrRetryScale` / sidecar `retry:` |
| `index.js` | `renderDirFor(file, tag)` / `renderPdfPage(s)` / `runRenderScript(…, scale)` | 重试渲染落独立目录 `<safe>-<mtime>-s<倍率>`（否则 `renderedPng()` 命中缓存 → 重试空转）；倍率经**子进程 env** 传给 `pdf-render.ps1`（并发安全）；`sweepRenderCache` 正则同步放宽 |
| `index.js` | `cacheDirState`；`ocrCachePath` / `readSidecarPath` / `writeReadSidecar` / `finishRead` | `DSH_OFFICE_CACHE_DIR` 生效（文件名规则不变）；未设逐字一致；不可写则回退并把原因写进 `stats.cacheDirNote` |
| `index.js` | `capWithOffset` / `finalizeToolValue` 内联护栏 | `content` 只放**纯前缀**，说明移入 `notice` + `stats.truncateNote`；硬不变式 `offset + content.length === nextOffset` |
| `rapidocr.js` | `engineArgs(rawScale)` / `ocrImages(…,{scale})` | 倍率可显式传入（渲染侧与引擎侧必须同值） |
| `test.mjs` | 新增 5 个测试段（30 条用例） | 截断协议 / 重试倍率纯函数面 / 缓存目录 / meta 画像 / 病灶样本端到端 |

### 用例名（同步段自带，可直接 grep）

- `护栏：capWithOffset 满足 offset + content.length === nextOffset`、`护栏：截断说明写在 notice 与 stats.truncateNote，不掺进 content`、
  `护栏：自动截断满足 offset + content.length === nextOffset`、`护栏：仅用 offset + content.length 驱动分段拼接，不重不漏`
- `重试倍率：默认表 2,1.5,1；空串关闭；垃圾值被滤掉且保序去重`、`重试倍率：当前生效倍率的接受窗口与 pdf-render.ps1 一致（0.5–4）`、
  `重试倍率：重试渲染目录与默认目录分离（否则 PNG 命中缓存，重试变空转）`
- `缓存目录：设置后 .ocr.md / .read.md 都落到该目录，文件名规则不变`、`缓存目录：降级 sidecar 真的写到该目录`、
  `缓存目录：不可写时回退默认位置并给出原因`、`缓存目录：回退原因进 stats.cacheDirNote（不静默）`
- `meta 画像：干净 PDF → textLayerUsable=true / garbledPages=[]`、`meta 画像：非 PDF 也有 textLayerUsable`、`meta 画像：乱码文件 → textLayerUsable=false`
- `病灶样本：meta 一眼看出文字层不可用（旧版此处"看起来完全健康"）`、`病灶样本：convert 无缓存 → 拒绝且四要素齐备`、`病灶样本：拒绝时绝不产出目标文件`、
  `病灶样本：有完整 .ocr.md → 转出成功且 fallback=ocr`、`病灶样本：转出内容确实来自 OCR（无 NUL / 控制符垃圾）`、
  `病灶样本：第 17 页自动通过（不再出现"未完成"）`、`病灶样本：stats.ocrRetried 记下换倍率的页`、
  `病灶样本：sidecar manifest 记 retry: 17=…，且 src 仍记 rapidocr`

### 行为边界

- 新行为一律"纯兜底或 opt-in"；不静默；干净文件逐字不变；重试串行。
- WorkBuddy 侧基准：`node test.mjs` **296 checks ALL PASS**（基线 258 → +38；跳过病灶 OCR 段为 288）。

---

## 2026-09-15 — 返回边界根治 + 读取降级链（自 WorkBuddy 版同步落地 · 第一轮）

### 问题

样本 `sample-B.pdf`（43 页，WPS 文字产物，文本层完整）：`as="meta"` ✅ 正常，
但 `pages="1"` / `"1-4"` / `"1-12"`、以及 `ocr="always" ocrEngine="local"` 全部 ❌
`tool "office_read" returned invalid output: value is not lossless JSON`；`office_convert` ✅。
即：**任何含正文的 office_read 全军覆没**。DSH 侧实测更彻底 —— 连干净单页也过不了（见本文 F1）。

### 根因（与最初"游离代理"的推断不同）

出站判定由 host 侧在工具体返回值离开 `execute()` 之前做（DSH 走 `@deepseek-ai/dsh-util-values`
的 `snapshotJsonValue`）。它**不检查字符串内容**（字符串只判 `typeof`），所以游离代理不会触发这条报错。
真正的元凶是**结构**：

```
$.stats.ocrCovered 是 undefined（host 会整条拒收）
```

`stats.ocrCovered: ocrInfo ? … : undefined` 这类"有则给、无则 `undefined`"的写法，
`JSON.stringify` 会安静丢掉，判定器却把它当成不可序列化值 → **整条结果作废**。
`as="meta"` 不含正文、走另一条分支所以幸免。一并收口的还有 NaN/±Infinity/-0、稀疏数组、
非纯对象（Date/Map/Set/Buffer/函数）、循环引用，以及会被跨进程替换的游离代理与 C0/C1 控制符。

### 补丁位置

| 文件 | 位置 | 内容 |
| --- | --- | --- |
| `index.js` | 「返回边界：码点消毒 + lossless-JSON 收口」段 | `sanitizeTextForReturn` / `losslessJsonProblem` / `losslessValueOf` / `finalizeToolValue` / `sanitizeThrown` |
| `index.js` | `defineToolLite()` | `execute` 与 `render` 接入 `finalizeToolValue()` —— 唯一出站终点（**DSH 版此前从未接线，是 F1 的直接原因**） |
| `index.js` | 「正文质量门 + 读取降级链」段 | `textQuality` / `finishRead` / `writeReadSidecar` / `readErrorHint` / `readSuggestion` / `capWithOffset` |
| `index.js` | `readPdf()` | 质量门 → 自动本地 OCR 升级；`stats.fallback` / `stats.ocrFromQualityGate`；CID 乱码附可复制参数串 |
| `index.js` | `isProtectedCacheName()` | `.read.md` 归入受保护缓存名 |
| `repro.mjs` | 新增 | 返回边界最小复现 + 常驻哨兵（红→绿，退出码 0/1） |
| `test.mjs` | 新增 5 个测试段 | 消毒 / 判定器 / 收口 / 护栏 / 质量门 / 降级链 / 全格式矩阵 |

### 用例名（可直接 grep）

`消毒：游离代理（高位）→ U+FFFD`、`消毒：NFC 规范化（e + U+0301 → é）`、`判定器：认出 undefined 属性（本次事故形态）`、
`收口：undefined 属性被剔除（回归本次整条拒收的元凶）`、`收口：正文脚注 + stats.sanitized 记账（绝不静默）`、
`护栏：超阈值改为首段 + nextOffset 续读协议`、`质量门：中文书里的英文/代码页不被误判（相对判据带 ASCII 字母护栏）`、
`降级：正文不可读 → stats.fallback = sidecar`、`错误四要素：页码 / 格式 / 根因 / 下一步 齐备`、
`矩阵：report.docx 的 meta/markdown/text/json 四形态全部成功且无损`（14 个格式夹具）、
`样本 PDF：pages="1" 返回无损（旧版此处整条拒收）`

### 附带修掉的一个真 bug

非 PDF 读取路径的 `truncated` 计算是 `offset + body.length < content.length`，而
`body = content.slice(offset)` → 左边恒等于 `content.length` → **`truncated` 恒为 `false`**。
现改为 `capWithOffset()`，`truncated` 真实生效并给出**精确** `nextOffset`。

> **更正（第二轮补丁）**：上面这段的上半部分仍然成立，下半部分已作废 —— 第二轮把截断协议收敛了，
> `content` 现在只放**纯前缀**（不带任何后缀），说明改放 `notice` 与 `stats.truncateNote`，
> 于是可以**直接**用 `content.length` 续读拼接。旧描述保留在此以保留修订轨迹。

### 行为边界

- 既有参数语义不变；消毒边界属 bugfix，**默认生效**。
- 新增行为一律是纯兜底（质量门命中才动、sidecar 只在正文不可读时写、护栏只在超阈值时触发）。
- 干净文件逐字不变：不写 `stats.quality`、`stats.sanitized = 0`、不加任何脚注。
- 不引入 Python / LibreOffice / pdftotext / npx；不改 dsh-badge 合同。
