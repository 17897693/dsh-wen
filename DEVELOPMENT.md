<!-- 本项目原名 dsh-office，后更名为 dsh-wen；本文档保留当时的名称与实测记录。 -->

# dsh-office 开发参考（Development reference）

> 这份文档装的是**逐轮修复史、内部实现细节与诊断过程** —— 插件的调用者不需要读它。
> 工具怎么用、边界在哪、失败后怎么办看 [SKILL.md](SKILL.md)；快速入门与格式能力矩阵看
> [README.md](README.md)。历史结论以**当轮的实测**为准；**当前行为**一律以 `index.js` + `test.mjs` 为准。

> ⚠ **验收硬规则（第十九轮 R19 固化；有 R13 / R16 的教训）**：**在"工作副本"里跑绿不算验收。**
> 任何改动的验收都必须走完三步：① 工作副本跑绿 → ② 一次性同步到目标目录
> `${DSH_HOME}/plugins/dsh-office/` → ③ **在目标目录再跑一次** `node test.mjs`，且这次
> **必须用 `danger-full-access`**（目标目录在会话工作区之外；受限沙箱下会**静默跳过**两类用例）。
> 受限沙箱下被跳过、只有提权才真跑的两类是：**原生 Office COM 冒烟**（Word / PowerPoint / Excel 真开箱）
> 与**非 `%TEMP%` 目标目录的 PDF 产出质量门用例**。历史教训：R13 / R16 出现过"套件全绿、交付物被
> Word 16.0 拒开（`0x80070570` / `0x800A1401`）"—— 绿的是跳过后的套件，**盖不住 Office 拒开**。
> 因此：**工作副本的 `N checks` 只是草稿；目标目录提权跑出的 `N checks` 才是验收数字。**
> 同步前要有全量 SHA-256 清单 + 被改文件的备份；跑完测试要再取一次清单证明"没有污染目标目录"。

## 目录

- [第十九轮（R19）：把 R18 的「剩余风险」钉成硬证据或明确结论](#第十九轮r19把-r18-的剩余风险钉成硬证据或明确结论)
- [第十八轮（R18）：PDF 产出质量门落点 / ZIP64 / 真实锁回归 / 输出边界 / 文档去重](#第十八轮r18pdf-产出质量门落点--zip64--真实锁回归--输出边界--文档去重)
- [第十七轮（R16）：缓存身份 / 原子落盘 / ZIP 限额 / schema 一致性](#第十七轮r16缓存身份--原子落盘--zip-限额--schema-一致性)
- [未实现能力：草案、结论与重开条件](#未实现能力草案结论与重开条件第十八轮-r18-从-skillmd-迁出)
- [从 README.md 迁出的实现细节与"当轮实测、未自动化"数字](#从-readmemd-迁出的实现细节与当轮实测未自动化数字第十八轮-r18-整理)
- 更早的逐轮修复注记（第七轮、第五轮、第六轮 …）见下方同名小节

## 第十九轮（R19）：把 R18 的「剩余风险」钉成硬证据或明确结论

> 对象：`${DSH_HOME}/plugins/dsh-office/`。**本轮不做新功能、不改任何既有判据** —— 只处理 R18 交付报告
> 「剩余风险」的六条，逐条给出**硬证据**或**明确结论**（不再以"风险"名义留白），并把验收流程固化成硬规则。
> 工作副本 `work/r19-src`（改动与测试都在那里），一次同步回目标目录；命令与数字见「交付报告」。
> 原计划"零代码改动"，但任务 D2 的**全仓正则审计**查出 **4 类真实可复现的 `RangeError`**（与 R18 修的那两条
> 同病），于是把范围收在"修掉这 4 类 + 补回归用例"（见 §D2）。
> 改动落在六个**解析层**文件（`model.js` / `index.js` / `xml.js` / `odf.js` / `docx.js` / `pptx.js`）
> 加 `zip.js` 的一处 JSDoc；**正常输入逐字不变**（每条都有等价性哨兵），
> 只有"输入大到旧写法会崩"时行为不同（旧：宿主级 `RangeError`；新：正常返回）。

### A（最高）ZIP64「>4 GiB 数据尺寸」分支：能证的那一半**已经证了**，剩下的落成**已知限制**

R18 的原文是"该分支只由 `zip64ExtraField` 的布局单测 + `baseOffset` 接缝（覆盖**偏移**分支）间接覆盖"。
本轮把它拆成三个可独立回答的问题：

| 问题 | 结论 | 证据 |
| --- | --- | --- |
| ① 我们写出的"≥4 GiB 尺寸声明"（哨兵 + ZIP64 extra 的 8 字节真值）**会被主流实现接受吗？** | **会** | 手工包交给 `.NET ZipArchive`：`Entry.Length` 读成 `4294967396`（§A1 / §A2） |
| ② "数据尺寸 ≥ 4 GiB"这条**代码路径**能不能端到端跑？ | **本机不能 → 已知限制** | 量化门槛：`makeZip` 峰值 ≈8.0 GB（输入 4.00 + 输出 4.00），本机 15.9 GB 总内存 / 空闲 2.9–3.6 GB（§A3） |
| ③ `opts.baseOffset` 到底是什么？ | **测试接缝、`@internal`**（JSDoc 已标注） | 本仓**零**生产调用点（§A4） |

#### A1 外部实现的行为边界（探针 `work/r19-probe/probe-zip64-declared.mjs`）

构造三种"中央目录 / local header **声明** ≥4 GiB、**实际数据只有 5 字节**"的包
（ZIP64 extra 直接调本仓 `zip64ExtraField()`，字段序与 `makeZip` 的 unc64 / comp64 分支一致），交给三方读取：

| 包形态 | 本仓 `openZip` | `.NET ZipArchive`（`Expand-Archive` 的引擎） | `Expand-Archive` 本体 |
| --- | --- | --- | --- |
| ① 未压缩尺寸 ≥4 GiB（compSize = 5，真值） | 打开 ✓ 列名 ✓；`get()` 按**声明**拦在单条目上限：`zip 条目解压后过大：big.txt 声明 4096.0 MiB，单条目上限 256 MiB` | `OPEN=ok`、**`LENGTH=4294967396`**、`COMPRESSED=5`；`READ=ok read=5` | `EA=ok files=1`（解出 5 字节） |
| ② 未压缩 + 压缩尺寸**都** ≥4 GiB（store 4 GiB 条目的真实形态） | 打开 ✓ 列名 ✓；`get()` 报 `zip 结构损坏：条目 "big.bin" 声明压缩长度 4294967396 超过文件长度 233` | `OPEN=ok`、`LENGTH=COMPRESSED=4294967396`；真读流报 `本地文件头已损坏` | 同样报 `本地文件头已损坏` |

**三条结论（A 最硬的收获）**：

1. **我们的 ZIP64 尺寸声明格式被 .NET 正确解析**（`Entry.Length` = 4 GiB+）⇒ "声明格式不被主流实现接受"
   这条风险**不存在**。
2. **.NET 不校验"未压缩尺寸 vs 实际解压量"**：形态①它按 `compSize` 只读 5 字节就返回成功。
3. **本仓读侧比 .NET 更严**：形态①按**声明**拦上限（R16 定的"绝不先解压"）、形态②直接判"声明压缩长度
   超过文件长度"。这条差异是**有意的**，写进文档以免下轮误判成"不一致"而放宽断言。

#### A2 真实 >4 GiB 合法条目（探针 `work/r19-probe/probe-zip64-real4gib.mjs`）

deflate 的极限压缩比 ~1032:1 ⇒ **4 GiB 的全 `A` 输入只产出 ~4 MB 压缩流**，可以 1 MiB 块循环流式喂入、
增量算 CRC32，**内存 O(1)** 地造出一个真实合法的 >4 GiB 条目：

```
输入         4,294,971,392 字节（4 GiB + 4 KiB，越过 0xffffffff 哨兵线）
deflate      12,902 ms → 4,174,511 字节（1028.9:1；CRC32 = 0xA5FF…）
包           4,174,731 字节；uncSize 写哨兵 0xffffffff + ZIP64 extra 真值，compSize 写真值
.NET         OPEN=ok COUNT=1 LENGTH=4294971392 COMPRESSED=4174511
             STREAM=ok read=1048576 nonA=0     ← 前 1 MiB 全为 'A'，解压内容正确
本仓 openZip  按声明拦在 256 MiB 单条目上限（可读错误，不先解压）
```

⇒ "ZIP64 声明 ≥4 GiB **且**数据真的 ≥4 GiB"这条组合**被主流实现正确接受、并能正确解压**，
不再只是布局级证据。

#### A3 为什么 `makeZip()` 的"数据尺寸 ≥4 GiB"分支**仍然端到端不可行**（明确结论）

`makeZip()` 对 `entry.data` 的处理链是：`crc32(data)`（**逐字节遍历**）→ 可选 `deflateRawSync(data)` →
`concatBytes([...])`（**再复制一份**成连续缓冲）。要让它走 `unc64 = data.length >= 0xffffffff`：

- 输入 `Uint8Array` 必须 **≥ 4.0 GiB 且真实可遍历**（`crc32` 会触碰每个字节 ⇒ 不能用"假 length"的对象）；
- `concatBytes` 还要再分配 **≈4.0 GiB** 输出；
- ⇒ **峰值 ≈8.0 GB 连续内存**（输入 4.00 GB + 输出 4.00 GB，未计 header / 中间件）。

本机实测（`work/r19-probe/probe-zip64-makezip-4gib.mjs` 的预检输出）：物理内存 **15.9 GB**、
探针运行期间空闲 **2.9–3.6 GB**（`os.freemem()` 两次实测；预检打印"需求 8.00 GB vs 门槛 9 GB，SKIP"）。
⇒ 本轮**明确不尝试**（强上会进换页、拖垮会话），把这条从"剩余风险"升级为 **已知限制**：

> **已知限制（R19 定性）**：`makeZip()` 的 `data.length ≥ 0xffffffff` 分支（以及 `compSize ≥ 0xffffffff`
> 分支 —— 它要求**压缩后** ≥4 GiB，更不可能）**没有端到端产物验证**，只有布局级 + 声明级证据
> （§A1 / §A2 + `zip64ExtraField` 单测 + `baseOffset` 的**偏移**分支）。
> **触发条件**：调用方把 ≥4 GiB 的单个条目交给 `office_create` / `office_edit` / `office_convert`
> —— 现实里到不了（OOXML / ODF 包的单部件不会有这个量级；插件还有 256 MiB 单条目**读取**上限）。
> **影响面**：若真有 4 GiB 级单条目，写出端会按规范写 ZIP64（§A1 / §A2 已证明声明被接受）；
> 读回先被 `DSH_OFFICE_ZIP_MAX_ENTRY_BYTES`（默认 256 MiB）拦下 ⇒ **不存在"静默产出废包"的路径**。
> **重开条件（可复现）**：在**空闲内存 ≥10 GB** 的机器上跑 `work/r19-probe/probe-zip64-makezip-4gib.mjs`
> （自带内存预检；加 `R19_ALLOW_4GIB=1` 才真分配），即可端到端复现，不需要改任何代码。

#### A4 `opts.baseOffset` 的定性（任务 A.3）

`zip.js` 里 `makeZip()` 的 JSDoc 已改为：**测试接缝、不是稳定 API（`@internal`）**，并写明
"本仓没有任何生产调用点（写出的包永远自包含、从 0 开始）"、保留理由（在自动化里覆盖 >4 GiB
**本地头偏移** → ZIP64 这条分支）、以及"将来真要嵌进别的容器时才提升为公开参数 + 补跨容器读回用例"。
核实：`grep -n 'baseOffset' index.js zip.js test.mjs` → 只有 `zip.js` 的签名 / 赋值 / 判据与 `test.mjs` 的用例。

### B `Expand-Archive` 真解包：**跑到底了**（+ 每文件开销对比）

R18 只跑到 900 s 超时（54,357/65,536，无错误）。本轮用 `work/r19-probe/probe-expand.mjs`
把同一个 65536 条目 ZIP64 包（6,269,334 字节）**真解包到底**，并同时量了 `.NET` 三种用法的每文件开销：

| 方式 | 结果 | 每文件开销 |
| --- | --- | --- |
| `.NET ZipArchive` `OpenRead` + `Entries.Count`（**套件默认用的那条**） | `COUNT=65536`，1,590 ms | 不解包 |
| `.NET ZipFile::ExtractToDirectory`（一次性全解） | 成功，65,536 个文件，75,046 ms | **1.145 ms** |
| `.NET OpenRead` + 逐条 `ExtractToFile`（抽 2,000 条） | 成功，2,000 个文件，4,544 ms | **2.272 ms** |
| **`Expand-Archive` 本体** | **status=0，65,536/65,536 个文件，2,180,154 ms（36.4 分钟）** | **33.27 ms** |

**结论（"套件默认只用 `ZipArchive` 读取"的量化依据）**：

- `Expand-Archive` 真解包**确实能跑完**（status=0、65,536 个文件齐全、零错误）⇒ 在 R18 的"能打开并列出"
  之外，"官方命令行工具能真解包我们写出的 ZIP64 包"现在也是实测过的。
- 但它的每文件开销是 `.NET ZipFile::ExtractToDirectory` 的 **29.05×**（33.27 vs 1.145 ms/文件）——
  PowerShell cmdlet 的每文件参数绑定 / 路径解析 / 管道开销吞掉了几乎全部时间。
  ⇒ 套件默认改用等价的 `.NET ZipArchive`（`Expand-Archive` 内部就是它）**读取**不是"偷工"，
  而是有量化依据的取舍；真解包保留为 opt-in（**发版前**手跑，见 SKILL「排查」的测试变量节）。
- 性能数字是**同一进程、同一时段**测的；跑的时候机器上还有套件与其它探针在跑（并发负载），
  所以绝对耗时偏高 —— **相对量级（29×）不受影响**。R18 记的"≈60 文件/秒"是空载时的数字，
  本轮负载下只有 ~0.5 文件/秒，两者不矛盾（都是同一 cmdlet）。

### C 现实世界的锁与网络盘
#### C1 瞬时锁：**退避窗口内释放 → 重试成功**（新用例 `R19-原子写：`）

R18 只证明了"锁一直不放 → 失败 + 退避窗口被走到（395 ms）"。本轮补另一半：子进程用
`FileShare.None` 拿住目标文件，**在主进程开始写之后 ~70 ms 释放**，`writeFileAtomic` 必须在
退避表（60 / 120 / 180 ms）走完之前重试成功。

握手设计（**R19 实测修正过一次 —— 那是夹具时序问题，不是实现问题**）：

- 主进程看到子进程写下的"已持锁"标志后，**先挂一个 70 ms 后写 release 标志的定时器、再立刻开始写** ——
  这样"首次 `rename` 一定落在锁持有期内"（否则测不出重试），而"释放时刻"由主进程控制，
  不依赖 PowerShell 的 `Start-Sleep` 精度；
- 第一版让子进程自己 `Start-Sleep 250 ms` 后关闭，实测释放晚于 360 ms 的窗口上界 ⇒ 用例红。
  改法只在测试夹具里（`test.mjs`），**实现一行没动**。

断言（全过）：`writeFileAtomic` 返回成功、终态 = `瞬时锁-新内容`（不是旧内容、也不是半截）、
不留 `.part` 临时件、耗时 ≥ 60 ms（证明**至少一次退避真的发生过**）。实测：**耗时 203 ms**（锁在主进程开始写之后 70 ms 释放 ⇒ 走完 60 + 120 两次退避后成功；
用例 detail 里同时打印了退避表与临时件残留检查）。

#### C2 跨卷 `rename` → `EXDEV`（新用例）

把代码注释里"临时件必须在目标同目录"这条约束变成**实测断言**：把 `%TEMP%`（C:）下的文件
`renameSync` 到另一个卷的可写目录 → 期望 `EXDEV`。网络盘 = 另一个卷 ⇒ 这条同时解释了
"为什么临时件不能放 `tmpdir()`"（R16 注释的第 ③ 条）。实测：`vol C: → D:`、`code=EXDEV` ✓（用例名 `R19-原子写：跨卷 rename → EXDEV…`）。

#### C3 网络盘 / OneDrive：**本机无法验证**（明确结论，不用模拟代替）

- `net use` → **无映射网络驱动器**（空列表）；`Get-PSDrive` 只有 `C:` / `D:`（本机两块盘）。
- `<home>/OneDrive` **目录存在，但里面只有一个 `desktop.ini`**（未登录 / 未同步）
  ⇒ 它不代表同步盘语义，**不能拿来冒充"网络盘实测"**。
- 因此本轮对"网络盘 / OneDrive 上的 rename 覆盖 + fsync 行为（含 `DSH_OFFICE_ATOMIC_FSYNC=0` 的差异）"
  的结论是：**本机无法验证**。已留 opt-in 用例 —— 有环境的机器上一条命令即可验证：

```
set DSH_OFFICE_TEST_NET_DIR=<网络盘目录>
node test.mjs    →  R19-原子写：网络盘 / OneDrive 上"写临时件 → rename 覆盖"成功且内容正确（opt-in）
```

- 能验证的邻近面本轮都补了：跨卷 `EXDEV`（§C2）、`DSH_OFFICE_ATOMIC_FSYNC=0` 与默认在本地卷上
  **语义一致且都成功**（新用例 `R19-原子写：DSH_OFFICE_ATOMIC_FSYNC=0 与默认在本地卷上语义一致…`）。

### D 大输入：全仓正则审计 + 阈值量化 + 端到端守门

#### D1 溢出阈值：逐档定位 + `--stack-size` 矩阵（探针 `work/r19-probe/probe-textquality-threshold.mjs`、`probe-textquality-stacksize.mjs`）

R18 只把溢出区间收窄到 `(1 MiB, 8 MiB]`。本轮用**真实旧实现**（精确照抄
`work/r18-run/backup-before/index.js` 的两个热点）逐档定位：

| 输入 | 1 MiB | 2 MiB | 4 MiB | 5 MiB | 6 MiB | 7 MiB | 8 MiB |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 字母 `a×N`（`matchAll(/[A-Za-z]{6,}/g)` **只迭代**） | ok 9 ms | ok 16 ms | ok 21 ms | ok 42 ms | **THROW `RangeError: Maximum call stack size exceeded`** | THROW 12 ms | THROW 11 ms |
| 字母 `a×N`（旧实现整体：迭代 + `split(正则)`） | ok 11 ms | ok 5 ms | ok 16 ms | ok 42 ms | THROW 10 ms | THROW 13 ms | THROW 12 ms |
| 点号 `.×N`（`matchAll(/\.{3,}/g)` 只迭代） | ok 4 ms | ok 1 ms | ok 2 ms | — | — | — | ok 4 ms |
| 点号 `.×N`（旧实现整体：把**整个匹配**的每个下标 `Set.add`） | 112 ms | 286 ms | 663 ms | — | — | — | **1,626 ms** |
| **当前实现**（手写扫描，零正则） | 22 ms | 74 ms | 66 ms | — | — | — | **105 ms**（字母）/ **75 ms**（点号） |

**结论（比 R18 的表述更准）**：

- 精确阈值落在 **(5 MiB, 6 MiB]**：5 MiB 通过、6 MiB 抛 `RangeError`。**是"单个巨大匹配"的长度**，
  不是"总输入长度"（1 MiB 的 20 组 `......` 完全正常）。
- 崩溃发生在 **`matchAll` 的迭代器本身**（"只迭代、不进循环体"那一档同样抛）—— 与 R18 的定位一致。
- 点号那条**不抛**，但旧实现要对整个匹配做 O(len) 次 `Set.add`：8 MiB → **1,626 ms**（R18 报 2,564 ms，
  同量级、机器/负载差异）；新实现 75 ms（≈**21×**）。
- **R18 写的"与机器栈大小相关"需要修正**：用 `--stack-size` 扫 **100 / 200 / 400 / 984 / 2000 / 8000 KB**
  六档 + `--no-opt`，**阈值一次都不动**（都是 5 MiB 过、6 MiB 崩）。
  ⇒ 这个 `RangeError` **不是 JS 调用栈深度**、也不是 JIT 优化状态带来的，而是 **V8 对"单个超长正则匹配"
  的内部限制**（随 V8 版本 / 平台 / 正则引擎实现变化）。因此它**不是可移植常数**：
  **不能写死一个"安全长度"，只能在实现层消除"整串正则"** —— 这正是 R18 的修法，本轮把它从
  "修复方式"升级为"唯一正确策略"。

#### D2 全仓正则审计（子代理独立审计 + 本轮复核；报告 `work/r19-probe/regex-audit.md`）

口径：只看"**可能超大长度**的字符串"上的 `match(/…/g)` / `matchAll` / `split(正则)` / `replace(正则)`；
判定标准是本轮**实测**出来的两类崩溃形态（比"正则长相"更准）：

| 会抛 `RangeError` 的形态 | 不会抛的 |
| --- | --- |
| ① `{n,}` 且 **n ≥ 4**（`a{4,}` 在 8 MiB 输入上就抛；`a{1,3,}` 不抛） | 简单原子上的 `+` / `*` / `{1,3,}` / `\.{3,}`（8 MiB ~ 64 MiB 单匹配实测安全） |
| ② 量词循环落在**复合体**上：体内含捕获组 / 交替 / 内层量词 / 回引用（`(a)+`、`(.)*`、`(?:a|b)*`、`(?:a{4})+`、`(?:[ \t]*\1){2,}` 全抛） | 非复合体的 `(?:a)+`、`(?:ab)*`、`(?:a){2,}` |

**统计：危险 14 处 / 需注意 27 处 / 已判定安全 25 处**（按"文件 + 符号"聚合）。

**本轮修掉（4 类，全部补了 `R19-大输入：` 回归用例）**：

| 点位 | 旧写法 | 修复前实测 | 现在 |
| --- | --- | --- | --- |
| `model.js::markdownToDocument` | `/^\s*([-*_])( *\1){2,}\s*$/` | 单行 2 MiB `-` → `RangeError` | 手写扫描（首字符 `-`/`*`/`_`、其后只允许空格或同字符、该字符 ≥3 个） |
| `index.js::pdfOutputGate` | `/^[ \t]*([-*_])(?:[ \t]*\1){2,}[ \t]*$/gm` | 整篇 8 MiB 单行 `-` → `RangeError` | 逐行 `stripMarkdownHrLines()`（命中行**只删内容、保留换行**，与旧 `replace` 等价） |
| `xml.js::parseXML` | `((?:[\s\/]+[^=>\s\/]+\s*=\s*(?:"[^"]*"\|'[^']*'\|[^\s>]+))*)` | 单标签内 2e6 个 ` a=a` → `RangeError` | `parseTagBody()` 手写扫描（name / 分隔符 / 属性名 / 值逐条对齐旧语义；不合法仍整条标签跳过） |
| `odf.js::decodeNumeric`（**有真实调用点**）、`docx.js::decodeEntitiesRaw`、`pptx.js::decodeRaw` | `String.fromCodePoint(+d)` | `&#` + 400 位数字 ⇒ `Infinity` ⇒ `RangeError: Invalid code point` | `codePointOrRaw()`（`Number.isFinite && 0 < c ≤ 0x10FFFF`，越界**原样保留**；与 `xml.js::decodeEntities` 同口径）。后两处当前**零调用点**（`grep` 只有定义行），守卫是为将来接上调用点时不再踩 |

修复前后对照（**同一份探针** `work/r19-probe/probe-biginput-sites.mjs`，断言一个字没改）：
修复前 **4 THROW**（`model` ×2 / `pdfOutputGate` / `parseXML`），修复后 **10/10 ALL OK**，
三条等价性哨兵全过（`---` 仍是 hr、`<w:p a="1" b='2' c=3>` 属性解析正确、hr-only 仍算"无可见文本"、
`&#66;` 仍解码成 `B`）。

**已判定安全（下轮不必重查）**：

- `zip.js` / `skill.js` / `pptx-edit.js` / `pdffont.js` / `pdfcrypt.js` / `png.js`：逐文件 grep **零命中**
  （没有任何 `match(/g)` / `matchAll` / `split(正则)` / `replace(正则)` / `exec` / `test`）。
- `textQuality()` 的两条（R18 已改手写扫描）：本轮复核确认**零正则**、O(n) 零分配。
- 19 个真实 `replace(正则)` 点位在 8 MiB 输入上实测 0.6–7.3 ms（全是简单原子，不抛）。

**需注意但本轮不动（登记 + 重开条件）**：

- **6 处"只为计数就物化用户正则匹配数组"的 `text.match(re)`**：`index.js::editTextFile`（**整文件**）、
  `index.js::replaceTextInOdfContent`、`index.js::applyReplacement`（xlsx 单元格 / sharedStrings）、
  `odf.js::replaceTextInOdfContent`、`docx.js::replaceTextInDocument`、`pptx.js::replaceTextInPptxPart`。
  `re` 来自 `op.find`：普通调用会经 `escapeRegExp` 成字面量（简单原子 ⇒ 安全）；只有调用方显式
  `regex: true` 且传 `{4,}` 这类形态才会崩 —— **那是调用方给的正则本身**，不由本仓决定。
  **重开条件**：若要消除"物化"，把 `text.match(re).length` 改成 `re.exec()` 循环计数即可（每处 ~3 行）；
  但要清楚这**不会**消除 `{4,}` 形态的崩溃（`exec` 同样抛），它只解决内存 / 耗时量级。
- **4 处"与输入同量级的数组 / Set 物化"**：`pdf.js::parseToUnicode`（`[...matchAll]`，8 MiB → 2,097,152 元素 /
  455 ms）、`index.js::editDocx`（Set）、`formula.js::formatTextValue`（8.4M 元素）、
  `html.js`（`exec` 2,097,152 次 / 153 ms，**不**物化数组）。真实文档里这些输入有自然上界
  （CMap 条目 / rels 条目 / 数字位数），本轮判定"有上界、量级可接受"；
  **重开条件** = 在真实语料里观测到量级问题。

#### D3 端到端守门（`test.mjs` 前缀 `R19-大输入：`）

| 用例 | 断言 | 实测 |
| --- | --- | --- |
| 8 MiB 单行 `.md` 走 `office_read` | 不抛 `RangeError`、耗时 < 30 s | 2,358 ms，content 120,000 字符（内联护栏截断） |
| 同上：截断必须**带说明**（不静默丢） | 要么完整、要么有 `truncateNote` | `正文超过内联上限 120000 字符，已截断 80000 字符；用 offset:120000 继续读取…` |
| "8 MiB 单行 XML 部件"的 `.docx` 走 `office_read` | 不抛、耗时 < 60 s | 96 ms |
| `markdownToDocument` 单行 2 MiB `-` | 不抛、且 `---` 仍识别为 hr | 10 ms |
| `parseXML` 单标签内 2e6 个属性 | 不抛 | 217 ms |
| `parseXML` 正常属性 | 与旧口径逐字一致 | `{"a":"1","b":"2","c":"3"}` |
| `readOdt` 含 400 位数字实体 | 不抛 | 5 ms |
| `pdfOutputGate` 整篇 8 MiB 单行 `-` | 不抛（hr-only 仍算无可见文本） | ok |

> 端到端那条（`office_read` 读 8 MiB 单行）是**工具边界**的守门：旧版在这里会抛宿主级 `RangeError`
> （整条 `office_read` 作废）；解析层那 5 条是**点位级**守门，直接钉住 D2 审计发现的四个修点。

### E 收尾卫生与流程固化

#### E1 `test-out/` 的 133 个历史文件 → **清理**（本轮同步时删除）

`test-out/` 是**历次 `node test.mjs` 在插件目录里直接跑（没设 `DSH_OFFICE_TEST_OUT`）留下的产物**：
`ensureOut()` 未设变量时选 `resolve(process.cwd(), 'test-out')`，而 SKILL 一直允许"在插件目录里直接跑"。
同步前实测清点：

| 维度 | 实测 |
| --- | --- |
| 规模 | **133 个文件、98.45 MB**，另 9 个子目录（`cache-root` / `r13-winrt-aes128` / `r13-winrt-plain` / `render-fix-png` / `scale-1` / `scale-1.5` / `scale-2` / `scale-default` / `scan-dir`） |
| 时间分布 | 2025-08-16（2）· 2026-08-27（1）· 2026-09-14（3）· **2026-09-25（126）** · 2026-09-26（1） |
| 扩展名 | pdf 39 · md 25 · docx 14 · png 14 · txt 8 · xlsx 8 · html 7 · csv 6 · pptx 4 · odt 3 · odp / ods / bmp / gif / xyz 各 1 |

**结论：与本轮无关，全是历史测试垃圾** —— 量最大的 2026-09-25 是 R13 / R14 轮次的 `node test.mjs`
产物；唯一一个 2026-09-26 08:48 的 `scan-dir/scan-garbled.ocr.md` 也是 R18 之前某次在插件目录里
跑测试留下的识别缓存。
**处置：清理**（连目录一起删）。理由：它们不属于交付物、会被 `office_read paths=[插件目录]`
之类的盘点扫到（其中 39 个 `.pdf` / 14 个 `.docx` 都是可读格式），保留只会让"插件目录 = 交付物"失效。
**防复发**：测试产物一律落 `%TEMP%`（显式设 `DSH_OFFICE_TEST_OUT`）—— SKILL「排查」已写，
R19 起目标目录复核也一律显式设它。

#### E2 工作区副本的归档策略 → **保留原名，不改名**

`work/` 下的目录一律按轮次前缀命名（`r15-*` / `r16-*` / `r17-work` / `r18-*`，本轮新增
`r19-src` / `r19-probe` / `r19-run`）：

| 目录 | 内容 | 策略 |
| --- | --- | --- |
| `rNN-src/` | 该轮**定稿的工作副本**（含 `vendor/`） | 保留 —— 它是"那一轮代码长什么样"的快照（可从目标目录重取） |
| `rNN-probe/` | 该轮独立探针（**证据可复跑**） | 保留 |
| `rNN-run/` | 该轮的 SHA-256 清单 / 备份 / 运行日志 | 保留 |

**不改名**：现有前缀已经能一眼看出属于哪一轮；改名只会破坏 R16 / R18 文档里的既有引用
（`work/r18-probe/probe-zip64.mjs` 之类）。唯一允许的清理是"磁盘压力下删 `rNN-src/vendor`"
（ONNX 模型 + exe，可从目标目录或 `~/.dsh/ocr` 重取）。

#### E3 `dsh-badge` 徽章 → **仍然跳过**（不自造）

R19 复核：`Test-Path ~/.dsh/skills/dsh-badge` → **False**；`~/.dsh/skills/` 下只有
`convert-documents-to-markdown` / `photo-to-comic` / `quark-video-reader` / `token-efficient-workflow`。
⇒ 按约定**继续跳过，不自造徽章**；文末那行署名说明已按本轮复核更新。

#### E4 验收硬规则 → 已写到本文档**顶部**的"⚠ 验收硬规则"框

内容：工作副本跑绿 ≠ 验收；必须同步到目标目录、**用 `danger-full-access` 再跑一次**；
受限沙箱下被静默跳过的是**原生 Office COM 冒烟**与**非 `%TEMP%` 目标目录的质量门用例**；
R13 / R16 的教训是"套件全绿盖不住 Office 拒开"。附同步前 / 后的全量 SHA-256 与跑后复取清单的要求。

### 交付报告（本轮实际执行的命令与结果）
**独立探针（都不在 `test.mjs` 里；全部落 `work/r19-probe/`）**：

| 探针 | 作用 | 关键输出 |
| --- | --- | --- |
| `probe-zip64-declared.mjs` | A1 外部实现行为边界 | 三种"声明 ≥4 GiB"包 × 三方读取（表见 §A1）；`.NET` 报 `LENGTH=4294967396` |
| `probe-zip64-real4gib.mjs` | A2 真实 >4 GiB 合法条目 | deflate 12,902 ms → 4,174,511 字节（1028.9:1）；`.NET` `LENGTH=4294971392`、`READ=1048576 NONA=0` |
| `probe-zip64-makezip-4gib.mjs` | A3 端到端复现探针（默认只预检） | `总 15.9 GB / 空闲 2.9 GB`；需求 8.00 GB ⇒ `SKIP=默认只预检`（要真跑加 `R19_ALLOW_4GIB=1`） |
| `probe-expand.mjs` | B 真解包 + 每文件开销 | `Expand-Archive` 65,536/65,536、2,180,154 ms、33.27 ms/文件（表见 §B） |
| `probe-textquality-threshold.mjs` | D1 阈值逐档（**真实旧实现**的两个热点） | 5 MiB ok / 6 MiB `RangeError`；点号旧实现 8 MiB 1,626 ms（表见 §D1） |
| `probe-textquality-stacksize.mjs` | D1 `--stack-size` 矩阵 | 100 / 200 / 400 / 984 / 2000 / 8000 KB + `--no-opt`：阈值**一次都不动** |
| `probe-biginput-sites.mjs` | D2 / D3 四个修点（修复前后**同一份**探针） | 修复前 **4 THROW** → 修复后 **10/10 ALL OK**（含 4 条等价性哨兵） |
| `regex-audit.md` + `regex-*.mjs` | D2 全仓审计（子代理独立完成，主代理按盘核对） | 危险 14 / 需注意 27 / 已判定安全 25；报告 29 KB |
| `hash-manifest.mjs` | 同步流程的 SHA-256 清单生成 | `176 files`（before） |

**套件（工作副本 `work/r19-src`，受限沙箱）**：

```
cd work/r19-src
$env:DSH_OFFICE_TEST_OUT = %TEMP%\dsh-office-r19-b
node test.mjs   → ✅ ALL PASS  (739 checks)   exit 0            # R18 基线 719 ⇒ 本轮 +20
node repro.mjs  → 🟢 GREEN（exit 0）
```

- **新增 20 项检查**（719 → 739）：`R19-zip：` 6 项、`R19-原子写：` 6 项、`R19-大输入：` 8 项。
- **跳过项**（测试自己打印原因，不算通过也不判红，与 R18 一致）：Word / PowerPoint / Excel COM 冒烟、
  `DSH_OFFICE_TEST_AES256_PDF` 未设、若干"真实样本存在性"夹具槽位未设；
  **本轮新增三类 opt-in**（`Expand-Archive` 真解包、真实 >4 GiB 条目、网络盘）；
  以及受限沙箱下被跳过的 `R18-质量门：非 TEMP 目标目录`（提权后在目标目录跑时**会真跑**，见下）。

**同步回目标目录 + 在目标目录再跑一遍（`danger-full-access`）**：

```
# ① 同步前：目标目录全量 SHA-256 清单（176 个文件 = 43 个插件文件 + 133 个历史 test-out）
node work/r19-probe/hash-manifest.mjs <目标目录> work/r19-run/sha256-before.txt      → 176 files
# ② 备份将被修改的 12 个文件 → work/r19-run/backup-before/（12 个文件）
# ③ 逐文件复制 12 个文件（每个都核对 src/dst SHA-256 相同，不一致即抛错）
# ④ 删除 test-out/（§E1 的决定）：133 个文件、98.45 MB
# ⑤ 同步后：node work/r19-probe/hash-manifest.mjs <目标目录> work/r19-run/sha256-after.txt → 43 files
node work/r19-probe/hash-diff.mjs sha256-before.txt sha256-after.txt
  → changed=12  added=0  removed=133
    changed : zip.js index.js xml.js model.js odf.js docx.js pptx.js test.mjs
              README.md SKILL.md DEVELOPMENT.md CHANGELOG.md
    removed : 全部是 test-out/*（133 个历史测试产物，见 §E1）
# ⑥ 在目标目录再跑一遍（这一次用 danger-full-access，见文首"验收硬规则"）
cd ${DSH_HOME}/plugins/dsh-office
$env:DSH_OFFICE_TEST_OUT = %TEMP%\dsh-office-r19-target
node test.mjs   → ✅ ALL PASS  (744 checks)   exit 0      # 工作副本 739 ⇒ +5（见下方说明）
node repro.mjs  → 🟢 GREEN（exit 0）
# ⑦ 跑完再取一次清单 → work/r19-run/sha256-postrun.txt
node work/r19-probe/hash-diff.mjs sha256-after.txt sha256-postrun.txt
  → changed=0  added=0  removed=0（43 个文件逐字不变）⇒ **跑测试没有污染目标目录**
```

同步的 12 个文件（SHA-256，前 16 位）：

| 文件 | SHA-256（前 16 位） | 文件 | SHA-256（前 16 位） |
| --- | --- | --- | --- |
| `test.mjs` | `5203604b03b0352e` | `docx.js` | `a523061987112439` |
| `zip.js` | `25baf5c81334c9c9` | `pptx.js` | `98d83b308df3b9e1` |
| `index.js` | `5ef54cfa86c5b092` | `README.md` | `0ca0ae5c897e2d8e` |
| `xml.js` | `1438b8b47efe7aaf` | `SKILL.md` | `7819103e8af112c2` |
| `model.js` | `f5cd35f9238345ab` | `DEVELOPMENT.md` | `4313b6bb87a3d3b2` |
| `odf.js` | `278bca39a5ec0285` | `CHANGELOG.md` | `06c0c97404ebc355` |

> 注：`sync.ps1` 因本机 `ExecutionPolicy`（脚本未签名）无法用 `& script.ps1` 直接跑，
> 实际执行用的是**等价的内联命令**（同一套逐文件 SHA-256 核对 + `throw` 保护），
> 输出与脚本一致；脚本本身留在 `work/r19-run/sync.ps1` 作为流程留档。

**目标目录那次比工作副本多 5 项**（`Compare-Object` 逐条确认 target-only 10 条、工作副本里对应的
5 条"跳过"条目被替换为真跑）：受限沙箱下被静默跳过的两类现在真的跑了。

- **非 `%TEMP%` 目标目录的 PDF 产出（4 条，全 PASS）**：`nt=<home>/.dsh\tmp\dsh-office-r18-nontemp`，
  create → `{"embedded":true,"renderCheck":"pass","firstPageBytes":15103,"images":0}`、convert → `firstPageBytes:15010`、
  空白夹具仍被拒且不落盘、目录里不留临时件 / 渲染副本。**这就是任务 A（R18）的正面验收路径**。
- **原生 Office COM 冒烟（6 条，全 PASS）**：`Word 16.0 能打开产出的 .docx`（`paragraphs=3`）、
  `Word 能把产出的 .docx 导出 PDF`（186,699 字节）、`Word 能打开产出的 .html`（`paragraphs=9`）、
  `R13-1 开箱：含图 .docx`（`inlineShapes=2`）、`R13-7 pptx：PowerPoint 16.0 能打开产出的 .pptx`（`slides=3`），
  以及**负向控制仍然 FAIL**（`notesMaster` 改回共用 `theme1` → `0x80070570`）。
  ⇒ 本轮改动（六个解析层文件 + `test.mjs`）**没有破坏任何真实 Office 打开能力**。
- 目标目录那次的 `R19-原子写：瞬时锁` 耗时 **206 ms**（工作副本 203 ms），同样走完 60 + 120 两次退避。

> ⚠ 口径说明（与 R18 同一处理）：**测试是在"代码文件已同步"之后跑的**，所以 744 checks 覆盖的就是
> 最终交付的代码。**文档**（`DEVELOPMENT.md` / `CHANGELOG.md`）在其后补了本轮的数字并**再同步一次**
> （第二次同步只碰这两个文档：`sha256-after` → `sha256-final` 的差异**恰为这 2 个文件**）。
> 代码文件（`zip.js` / `index.js` / `xml.js` / `model.js` / `odf.js` / `docx.js` / `pptx.js` / `test.mjs`）
> 在那一跑之后**一字未动** —— 与上表的 SHA-256 逐字一致。

### 剩余风险 / 已知限制（逐条明确结论，不留白）
对照 R18 的六条"剩余风险"，逐条给出**明确结论**（不再留白）：

| R18 的剩余风险 | 本轮结论 | 证据 |
| --- | --- | --- |
| ① >4 GiB **数据尺寸**分支未端到端 | **已知限制**（不是风险）：`makeZip` 要 ≥4 GiB 真实输入 + 等量输出（峰值 ≈8.0 GB），本机 15.9 GB 总内存 / 空闲 2.9–3.6 GB。**声明格式已被 .NET 接受**（`Entry.Length` = 4 GiB+；真实 4 GiB 条目也能被正确解压）。触发条件 / 影响面 / 重开探针见 §A3 | §A1 / §A2 / §A3 |
| ② `Expand-Archive` 只跑到超时 | **已闭环**：跑到底（`status=0`、65,536/65,536 个文件、36.4 分钟），并量化出 **29×** 的每文件开销 ⇒ 默认用 `.NET ZipArchive` 有据可依 | §B |
| ③ 瞬时锁未被验证（只有"锁不放 → 失败"） | **已闭环**：新增"锁在退避窗口内释放 → **重试成功**"用例（终态新内容、原件未破坏、无残留） | §C1 |
| ④ 网络盘 / OneDrive 的 rename 与 fsync 未验证 | **本机无法验证（明确结论）**：无映射网络驱动器；`~/OneDrive` 未登录 / 未同步。已留 opt-in 用例 `DSH_OFFICE_TEST_NET_DIR`，有环境的机器一条命令即可验证；**没有拿本地目录冒充**。可验证的邻近面已补：跨卷 `rename` → `EXDEV` 实测 | §C2 / §C3 |
| ⑤ CRC 严格拒绝策略下"可用但 CRC 不符"的包会被拒 | **策略保持不变**（R18 的理由仍成立：自动关校验 = 把不可信数据当可信）；两个逃生口（`crc === 0` 放行、`DSH_OFFICE_ZIP_CRC=0`）已由用例守门 | R18 §3；本轮未动 |
| ⑥ `textQuality` 栈溢出的精确阈值未定 / 别处是否同病 | **已闭环**：阈值量化到 **(5 MiB, 6 MiB]**，并证明它**不是可移植常数**（`--stack-size` 六档不移动）；**全仓审计**查出 14 处危险，其中 **4 类真实可复现的 `RangeError` 本轮已修**（含 `xml.js` —— 所有 OOXML / ODF 部件的入口），其余登记判定 + 重开条件 | §D1 / §D2 / §D3 |

**本轮新增的明确限制（诚实登记）**：

1. **`Expand-Archive` 的绝对耗时随负载变化很大**（空载 R18 记 ≈60 文件/秒；本轮并发负载下 ~0.5 文件/秒）。
   它只影响"发版前手跑一次"的等待时间，不影响任何对外行为。
2. **`{n,}`（n ≥ 4）形态的用户正则仍会崩** —— 出现在 `office_edit regex: true` + `find: "a{4,}"` 这类
   "调用方自己给的正则"上：这是 V8 正则引擎对超长输入的限制，本仓**没有**（也不该有）
   "限制调用方正则"的策略。6 处 `text.match(re)` 的"物化"已登记判定与重开条件（§D2），本轮不改。
3. **网络盘 / OneDrive 的实测仍然缺失**（本机无环境）—— 这是"**本机无法验证**"，不是"已验证安全"。
4. **`docx.js::decodeEntitiesRaw` / `pptx.js::decodeRaw` 仍是零调用点**（守卫已加）；守卫让将来接上时
   不再崩，但"接上后行为是否正确"需要另配用例。

---

## 第十八轮（R18）：PDF 产出质量门落点 / ZIP64 / 真实锁回归 / 输出边界 / 文档去重

> 对象：`${DSH_HOME}/plugins/dsh-office/`。本轮是**缺陷修复 + 收尾加固 + 文档治理**：
> 修掉一条"在任何非 `%TEMP%` 目录都 100% 失败"的真实功能缺陷（A）、补上一处**号称修过其实没修好**
> 的栈溢出（E3）、把写出端 ZIP64 补全（B）、用**真实文件系统错误**而不是注入钩子覆盖原子写（C）、
> 加固输出边界与 schema（D）、并把两份文档的重复维护收敛掉（F）。
> 工作副本 `work/r18-src`（改动、测试都在那里），一次同步回目标目录；命令与数字见「交付报告」。

### 1（A，最高优先）渲染校验副本必须落 `%TEMP%`——否则非 TEMP 目标的 PDF 产出 100% 失败

**病灶（探针 `work/r18-probe/probe-winrt.mjs` 实测）**：WinRT 的
`StorageFile.GetFileFromPathAsync` 在本机**只允许 `%TEMP%` 下的文件** ——

| 探针 PDF 位置 | 结果 |
| --- | --- |
| `%TEMP%\<session>\…\probe.pdf` | `png=1`（12213 字节），`status=0` |
| `<corpus>/…\probe.pdf`（工作区） | `png=0`，`status=1`，日志为空（"拒绝访问，该项目没有位于应用程序可以访问的位置"） |
| `…\work\r18-src\test-out\probe.pdf` | 同上 |

而 `pdfOutputGate()` 旧实现把渲染副本写在**目标同目录**（`writeTempBeside(file, …)`）⇒
**在任何工作区 / 用户目录 create 或 convert 一个 PDF 都会 100% 失败**，报"第 1 页渲染失败…逃生通道"。
R16 只是把测试产物目录挪到 `%TEMP%` 绕过了它，**代码没修**。

**修法**（`pdfOutputGate` + 新增 `pdfGateStagingRoot()`）：

- 渲染副本改写到 `tmpdir()` 下（默认 `%TEMP%\dsh-office-pdfgate\<tag>\check.pdf`，
  `DSH_OFFICE_PDF_GATE_DIR` 可覆盖该根目录），渲染完立即删除；
- 校验**通过后**才复用公共两段式（目标同目录唯一 `.part` 临时件 + `rename`）发布 ——
  「未经校验不发布」的语义一字不变，且**不再为了渲染往目标目录写任何东西**；
- 失败路径与旧版一致：清理暂存件与临时件、不产出目标、抛「PDF 产出拒绝｜四要素」；
- 渲染失败时若暂存目录不在 `%TEMP%` 内，错误文本追加一句位置提示（`renderLocationHint()`）——
  否则把 `DSH_OFFICE_PDF_GATE_DIR` 指错地方只会看到一句无信息量的 `exit=1`；
- `stats.pdfQuality` 的字段名与语义**完全未变**（`embedded` / `renderCheck` / `firstPageBytes`
  + `images` / 条件性 `imagesSkipped` / `notice`）。

**验收（`work/r18-probe/probe-gate.mjs`，工作区目录、非 TEMP）**：

```
[1 create]  exists= true  pdfQuality= {"embedded":true,"renderCheck":"pass","firstPageBytes":15103,"images":0}
[2 convert] exists= true  pdfQuality= {"embedded":true,"renderCheck":"pass","firstPageBytes":15010,"images":0}
[3 blank]   exists= false err=【PDF 产出拒绝｜四要素】…根因=渲染为空白（首页 PNG 8606 字节、黑像素覆盖率 0.000%（0 px）、可见文本 10 字…）
[4 stray]   (none)      [4 staging leftovers] 0
[5 override 指到非 TEMP 目录] exists= false err=…第 1 页渲染失败（exit=1；注意=渲染暂存目录 … 不在 %TEMP% 下…）
```

### 2（A 连带）R13-2 两条 FAIL 的**真实根因**与修法

两条用例是 `R13-2 判别实验（WinRT 打开同骨架 AES-128 夹具）` 与 `R13-2 判别实验：探针本身有效`。
它们的失败原因**不止一个**，实测定位到两条（**都不是加密骨架的问题**）：

1. **沙箱下管道 stdio 被拒**。用例自己 `spawnSync('powershell.exe', …, { encoding:'utf8' })`
   （默认 `stdio:'pipe'`）→ 受限沙箱里直接 `EPERM`，`stdout/stderr` 全空、没有 PNG。实测：
   `[TEMP] png=0 exit=null err=spawnSync powershell.exe EPERM`（同一路径换成 `stdio:'ignore'` 就 `png=1`）。
   **基线跑法里产物目录本来就在 `%TEMP%`**（`DSH_OFFICE_TEST_OUT=$env:TEMP\…`），所以"探针 PDF 不在
   TEMP"并不是这两条 FAIL 的直接原因 —— 它是**第二个**（潜伏的）原因。
2. **WinRT 只读 `%TEMP%`**（同 §1）。一旦产物目录退化到工作区（不设 `DSH_OFFICE_TEST_OUT` 时
   `ensureOut()` 会选中插件目录下的 `test-out/`），探针 PDF 就不在 `%TEMP%` 了。

**修法（与实现同源，断言一个字没放宽）**：探针 PDF 改为写进 `pdfGateStagingRoot()`（渲染暂存目录、
默认 `%TEMP%`），渲染统一走新导出的 `runRenderScript()`（它自带"管道被拒 → `stdio:'ignore'` 重来"兜底）。
修后 `plainPng=1 aes128Png=1`，并把结论写进用例名：
**"WinRT 能打开同骨架 AES-128 夹具 ⇒ 骨架正确；AES-256 打不开指向 WinRT 不支持 R5/R6 或 V5 字典细节"**。

### 3（B）ZIP 写出端支持真实 ZIP64 + CRC 误拒面定策略

旧写侧对 `>65535 条目` / `>4 GiB` 一律报错（"当前实现不写 ZIP64"），而更早的版本会用
`Math.min(entries.length, 0xffff)` **静默丢条目**。现在（`makeZip`）：

- 条目数 ≥ `0xffff`、尺寸 / 压缩尺寸 / 本地头偏移 ≥ `0xffffffff` ⇒ 按 APPNOTE 写
  **ZIP64 扩展字段（ID 0x0001）+ EOCD64（56 字节，`size=44`）+ 定位记录（20 字节）**；
  中央目录的扩展字段只列"真的写成哨兵"的那些值、顺序固定（未压缩尺寸 → 压缩尺寸 → 本地头偏移）。
- **阈值取 ≥ `0xffff` 而不是 `> 0xffff`**：`0xffff` 本身就是"见 ZIP64 记录"的哨兵，
  `.NET` 的 `ZipArchive` 见到普通 EOCD 里的 `0xffff` 会去找 ZIP64 记录。留一个恰好 65535 条目、
  只写普通 EOCD 的包会踩这条歧义，所以从 65535 起就写 ZIP64。
- **不需要 ZIP64 的包字节布局与旧版逐字一致**（这条有专门探针，见交付报告）；
  `store` 语义与 ODF 的"mimetype 首个且不压缩"未动。
- **条目名 > 65535 字节仍然显式报错** —— 这条**没有 ZIP64 出路**：`nameLen` 在 local header 与中央目录
  里都是 16 位字段，ZIP64 只扩尺寸与偏移，**规范里没有扩展条目名长度的记录**（APPNOTE 4.5.3）。
  报错文案已改成说清这一点（旧文案只说"装不下"）。
- 新增 `opts.baseOffset`（默认 0）：把整包当成从该偏移开始的片段记账（把 zip 嵌进别的容器时有用），
  同时是"覆盖 >4 GiB 偏移分支"的唯一可科目B试接缝（真造 4 GiB 条目不现实）。

**CRC 误拒面：选择"保持严格拒绝"**（不是"静默自愈重试"）。理由：

- 中央目录声明了非 0 的 CRC、而数据流算出来不符，意味着**流真的与声明的不是同一份内容**
  （等长替换、传输损坏、写包器 bug 都会长这样）。自动以关闭 CRC 重试**等于把不可信数据当可信**，
  与插件"绝不静默"的总体原则冲突（渲染质量门、zip 限额、sidecar 身份都选了"拒绝并说清"）。
- 已经存在的两个逃生口足够：`crc === 0` 一律放行（旧式流式写包器不填 CRC 的包不受影响）、
  `DSH_OFFICE_ZIP_CRC=0` 显式关闭。错误文案**点名条目**并直接写出逃生口，
  本轮补了断言：CD CRC 与数据流不符（数据流被**等长替换**）→ 拒绝 + 点名 + 逃生口；
  置 `DSH_OFFICE_ZIP_CRC=0` 后同一份样本可读；`crc=0` 样本默认放行。

### 4（C）原子写：用**真实文件系统错误**覆盖（不再只有注入钩子）

R16 只有 `DSH_OFFICE_ATOMIC_FAULT=…` 注入式覆盖，没有真实错误。本轮补上（`test.mjs` 前缀 `R18-原子写：`）：

| 场景 | 机制 | 实测结果 |
| --- | --- | --- |
| 只读目标（`chmodSync(0o444)`，Windows 上即"只读属性"） | — | `rename` 报 **EPERM**；抛四要素；**原文件仍在且逐字节不变**；不留临时件；**没有 unlink 只读目标** |
| 同进程持有句柄（`openSync(target,'r+')`） | — | 同样是 **EPERM**（Windows 语义下 rename 覆盖被挡），四要素 + 原内容不变 + 无残留 |
| 跨进程真锁（子 PowerShell 以 `FileShare.None` 打开并保持） | 一次性脚本 + 就绪/释放标志文件（`stdio:'ignore'`） | **EPERM**；四要素；原文件逐字节不变；**退避重试窗口真的被走到**（实测 395 ms ≥ 60+120+180 ms 的上界） |
| `writeFileAtomicSync`（sidecar 路径） | 只读目标 | 四要素 + 原内容不变 + 目标未被删 + 无残留 |

**为什么不用 `chmod` 目录**：Windows 上无效，挡不住 `rename`。**为什么不用注入钩子顶替**：注入只能证明
"错误路径被写了"，证明不了"真实锁下原文件安全、且重试窗口真的跑到"。

### 5（D）输出边界与 schema 加固

- `defineToolLite` 的 `output.schema` 增加 `properties: { content: { type: 'string' } }`（**不加 `required`**）。
  动手前核实了全部 `content` 赋值都是字符串（`capText` / `capWithOffset` 的纯前缀、PDF 的
  `JSON.stringify(bare)`、`office_read` 正文、批量盘点管道表、convert 报告），并用
  `as=json/meta/markdown/text + paths` 五形态断言"一条也没被误伤"。
- `itemViolations` 的边界文案收敛：`null` / `undefined`（含**稀疏数组空洞**）→ "不能为空（null/undefined）"；
  字符串 / 数字 / 数组 → "应为对象…**实际收到**字符串/数字/数组"。校验循环同时从 `forEach` 改成
  **下标循环** —— `forEach` 会静默跳过空洞，旧版 `operations: [ <hole> ]` 会一路滑到执行期变裸 `TypeError`。
- **新收窄**（schema 早就声明 string，运行时却 `String(op.x ?? '')` 静默字符串化）：
  `find` / `replace` / `text` / `markdown` / `cell` / `name` / `newName` 非字符串 → 带下标的参数错误。
  **兼容性影响**：过去传数字（`find: 2024`）会被静默当成 `"2024"` 搜索，现在会被**明确拒绝**；
  迁移方式 = 传字符串。**刻意没有收窄** `value`（数字是常态）与 `regex` / `whole`（truthy 习惯）。
- `.ocr.json` 这个"遗留保护名"复核结论：**保留**（本仓确实没有任何写入点，但历史版本 / 姊妹实现可能已经
  把它写进共享缓存目录；删掉只会带来"清扫器删掉用户识别成果"的单向风险）。已写进 `index.js` 的
  `isProtectedCacheName` 注释，并由既有用例 `P2-5 .ocr.md/.ocr.json 与其所在目录永不碰` 守门。

### 6（E）测试卫生、端到端补面与基准

- 清掉 `test.mjs` 的未使用 import（`pathHashOf`、`writeFileAtomicSync` —— 后者在 §4 的新用例里又重新用上了）。
- 新增**真实大 stored 图片条目**端到端：2000×2000 噪声 PNG（11.45 MiB）→ `office_create` + `insert_image`
  → `office_read` 回读；断言媒体部件**逐字节一致**、**以 `store`（method 0）落地**、
  远低于单条目上限（11.45 MiB vs 256 MiB）。噪声源必须是**不可压缩**的：先用 LCG 生成时被 deflate
  压到 0.10 MiB（"大条目"直接不成立），改用 64 KiB 随机池循环取样（池距 > deflate 的 32 KiB 窗口）才拿到 11.45 MiB。
- 大输入基准（详见 §7 数字表）：`textQuality` 对 8 MiB 连续字母 / 8 MiB 连续句点 / 8 MiB 正常中文 /
  8 MiB 混合各测耗时与 RSS，**并修掉一个真 bug**：
  R16 把 `s.match(/[A-Za-z]{6,}/g)` 换成 `matchAll` 并没有修好栈溢出 ——
  **单个巨大匹配时 `RegExpStringIterator.next` 自己就炸**（8 MiB 连续 `a`；8 MiB 连续 `.` 也一样，
  旧版是 2564 ms 而不是崩）。现在 `textQuality` 里两条整串正则全部改成**手写扫描**
  （点前导就地识别、字母连排一遍数完），零正则、零中间字符串、O(n) 零分配；
  判据口径用小样本与旧"正则 + `split`"参考实现**逐样本对照一致**。
- zip `maxOutputLength` 早停的 RSS 断言：用"声明 1 MiB、实际 64 MiB"的 deflate 流，
  父进程量早停增量、**子进程真解压完**量干净基线，判据 `早停 Δ < 8 MiB` **且** `真解压 Δ > 早停 Δ × 4`
  （自校准阈值，不写死机器相关常数）。实测 0.9 MiB vs 131.5 MiB。

### 7 基准数据（本轮实测；`work/r18-probe/probe-bench.mjs`，单人单机、无并发）

`textQuality`（修复后）：

| 输入 | UTF-8 字节 | 耗时 | RSS 增量 | `visible` |
| --- | --- | --- | --- | --- |
| 8 MiB 连续字母 `a×8MiB` | 8,388,608 | **128.0 ms** | +11.1 MiB | 8,388,608 |
| 8 MiB 连续句点 `.×8MiB` | 8,388,608 | **86.6 ms** | +9.0 MiB | 0 |
| 8 MiB 正常中文 | 7,980,000 | **38.6 ms** | +5.3 MiB | 2,660,000 |
| 8 MiB 混合 | 8,200,000 | **67.6 ms** | +10.3 MiB | 4,600,000 |
| 1 MiB 连续字母 | 1,048,576 | 14.8 ms | +1.0 MiB | 1,048,576 |

修复前（用目标目录里未改动的 `index.js` 跑同一组，`probe-bench-old.mjs`）：

| 输入 | 修复前 |
| --- | --- |
| 8 MiB 连续字母 | **`RangeError: Maximum call stack size exceeded`**（149 ms 后崩） |
| 8 MiB 连续句点 | 不崩，但 **2564.2 ms**（Set 建 800 万条）→ 修复后 86.6 ms（**约 30×**） |
| 1 MiB 连续字母 | 正常（31.6 ms） |

⇒ 溢出阈值落在 (1 MiB, 8 MiB] 区间，且与机器栈大小相关（**不是**一个精确、可移植的常数）。

zip `maxOutputLength` 早停（同一轮实测）：被拒绝的一次 `RSS +0.9 MiB`；把同一段 deflate 流交给
干净子进程真解压完 64 MiB 时 `RSS +131.5 MiB`（超线性是因为 V8 同时保留了输出缓冲与 zlib 内部窗口）。

### 8（F）文档治理：去重、清理不可证实内容、符号化定位、字段补全

- **README.md 瘦身到"快速入门 / 格式能力矩阵 / 常见用法 / 指向"**：删掉与 SKILL 重复维护的长段
  （扫描件 OCR 的 86 行实现细节、加密算法细节、84 行"已知边界"），保留每条一句的**速查边界** +
  指向 SKILL / DEVELOPMENT 的入口；OCR 实现两级、长页面分带等**README 独有的实现细节**迁入本文件。
- **SKILL.md 保留操作性内容**（工具选择 / 读取顺序 / 参数用法 / 仍然生效的边界 / 排查），
  并把五类"未实现能力的草案"整段迁到本文件（见下节），SKILL 只留一句结论 + 替代做法 + 重开条件入口。
- **清理不可证实的数字**：SKILL 里"23 个真题 PDF 约 1.1 秒""某 35 页样本第 17 页 `scale=2` 过"
  "56,341 字节 / NUL 5135 / C0 13666""510 个数据格 4.9% / 密集表 10–15% / `分析6-p7`"
  "9000 字 PDF 0.39 MB vs 4.35 MB""本机原生 992×1403（120dpi）"以及**整张渲染倍率实测表**，
  一律改为可证实表述 + 指到本文件；渲染倍率表已迁入本文件（§9）。README 里"第 N 轮实测 XXX checks"
  类表述清掉，统一改成"数量看最后一行的 `N checks`"。
- **行号引用全部符号化**：SKILL 里 8 处 `index.js:xxxx` / `test.mjs:xxxx`（以及 `rapidocr.js:` /
  `repro.mjs:` / `pdf-render.ps1:` 的行号）已随 R16 漂移；现在**一处不剩**（`grep '\.js:[0-9]' SKILL.md`
  只剩文末那句说明），统一改成 `grep -n '<符号名>'` 的定位方式，并在文末写明"本文档不再用行号定位"。
- **字段与口径补全**：`stats.pdfQuality` 的 `images` / 条件性 `imagesSkipped` / `notice`、
  批量盘点行的 `blocks` / `sheets` / `slides`、`suggestedBatches` 在干净 PDF 上**只有 >12 页**才按
  15 页/批（≤12 页算 1 批）、`DSH_OFFICE_MAX_INLINE_CHARS` 非 0 且 <1000 时回落 120000、
  以及**缓存身份的已知限制**（路径指纹是规范化路径字符串的 SHA-256，**不做短名 / 符号链接 / 网络盘别名归一化**
  ⇒ 同一文件经不同路径形态会被当成不同来源，表现是"缓存没命中、白重识别一次"，**不会**误用别人的缓存）。
- **环境变量总表补全**：`DSH_OFFICE_PDF_GATE_DIR`（新）、`DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` /
  `_MAX_TOTAL_BYTES` / `_ZIP_CRC`、`DSH_OFFICE_ATOMIC_FSYNC` 都进了 SKILL 的「调优环境变量」表
  （那张表自称全集，之前少了这几个）。

### 9 从 SKILL.md 迁出的"渲染倍率实测表"（第十六/十七轮当轮实测，未自动化）

真实扫描件第 5-6 页，本地引擎，0 次视觉调用：

| scale | PNG | 2 页耗时 | 平均置信度 | 升级视觉的页数 |
| --- | --- | --- | --- | --- |
| 不设（本机原生 120dpi） | 902×1257 | 5.8 s | 0.996 | 1 |
| 1 | 723×1006 | 4.6 s | 0.995 | 1 |
| 1.5 | 1084×1509 | 6.4 s | 0.994 | 1 |
| 2 | 1444×2011 | 8.9 s | 0.994 | 1 |

结论：**放大只增加耗时、不减少升级视觉的页数**（2× 时接近 2 倍时间）→ 默认保持不设。
已有自动化守门的只有"A4 像素 = 96dpi×scale"（用例 `P1-3b 各档 PNG 尺寸`：`1`→794×1123、
`1.5`→1191×1685、`2`→1588×2246，±2%）；"不设"时的 992×1403 跟随系统 DPI，**换机器就会变**。

### 10 交付报告（本轮实际执行的命令与结果）

**基线**（改动前，同一环境）：

```
cd work/r18-src                                        # 目标目录的原样副本
$env:DSH_OFFICE_TEST_OUT = Join-Path $env:TEMP 'dsh-office-r18-baseline'
& <node> test.mjs
→ 656 checks / 2 FAILED（exit 1）      # FAIL 就是 R13-2 那两条判别实验
```

**改动后**（最终一次，代码与测试都已定稿）：

```
cd work/r18-src
$env:DSH_OFFICE_TEST_OUT = Join-Path $env:TEMP 'dsh-office-r18-e'
& <node> test.mjs
→ 719 checks / ALL PASS（exit 0，764 行输出）
```

- 新增 **63 项检查**（656 → 719），全部通过；**上一轮遗留的 2 项 FAIL 已消除**（见 §2）。
- **跳过项**（测试自己打印原因，不算通过也不判红；与基线一致，本轮另加一类）：
  Word / PowerPoint COM 冒烟（本机 Office 不可用 → 打印原因）；`DSH_OFFICE_TEST_AES256_PDF` 未设；
  若干"真实文件存在性"用例因样本文档不在本机；多份 `DSH_OFFICE_TEST_*` 夹具槽位未设；
  **本轮新增两类**：① `R18-质量门：非 TEMP 目标目录 create/convert` —— 在 `workspace-write` 沙箱下
  找不到"可写的非 `%TEMP%`、且不在插件目录里"的目录（用 `danger-full-access` 重跑即可覆盖，
  目标目录复核那次就是这样跑的）；② `R18-zip：Expand-Archive 真解包 ZIP64` —— opt-in
  （`DSH_OFFICE_TEST_EXPAND_ARCHIVE=1`），本机解 65536 个文件 **>10 分钟**（见下）。
- **另一路自检**：`node repro.mjs` → 🟢 GREEN（退出码 0，"返回值可无损序列化，harness 一定收"）。

**独立验证（不在 `test.mjs` 里）**：

- `work/r18-probe/probe-zip64.mjs` —— **写出端逐字节对照**：4 组条目（纯 store / ODF 形态 / 大文本 deflate /
  空条目+中文名）用**旧 `zip.js` 快照**与新版 `makeZip` 分别产出，在同一个 2 秒 DOS 时间桶内
  `oldB.equals(newB)` **全为 true**（SHA-256 逐组相同）⇒「<4 GiB 包字节布局与旧版逐字一致」是**实测**，不是推断。
- 同一探针：65536 条目 → 6,269,334 字节、构建 505 ms、`openZip` 回读 65536 条、
  EOCD 计数字段 `0xffff`、定位记录 + EOCD64 存在且计数为真值。
  **`Expand-Archive` 真解包**：900 秒内解出 **54,357 / 65,536** 个文件后被探针的超时杀掉 ——
  **它能正确解包，只是 PowerShell 的 `Expand-Archive` 每文件开销极大（≈60 文件/秒）**，
  所以套件里默认用等价的 `.NET ZipArchive`（`Expand-Archive` 内部就是它）读，
  并把真解包列为 opt-in。
- `work/r18-probe/probe-winrt.mjs` / `probe-gate.mjs` / `probe-bench.mjs` / `probe-bench-old.mjs` /
  `probe-textquality.mjs` —— 分别给 §1 的位置限制、A 的端到端验收、§7 的基准与修复前后对照。

**本轮新增/更新的回归测试清单**（前缀 → 位置）：

| 前缀 | 覆盖 |
| --- | --- |
| `R18-质量门：` | 暂存目录默认在 `tmpdir()`、非 TEMP 目标 create/convert、空白仍拒且不落盘、`DSH_OFFICE_PDF_GATE_DIR` 覆盖、覆盖目录渲染后清空 |
| `R13-2 判别实验…` | 改为与实现同源（暂存到 `%TEMP%` + `runRenderScript`），断言未放宽 |
| `R18-zip：` | 旧版字节布局（3 组）、65536 条目 ZIP64（EOCD64/定位记录/计数）、`openZip` 回读、`.NET ZipArchive` 兼容、`Expand-Archive` opt-in、>4 GiB 偏移分支（哨兵 + extra + EOCD64）、超长名显式错误、`zip64ExtraField` 布局、CRC 三态（等长替换拒绝 / 逃生口 / `crc=0` 放行）、早停 RSS 自校准阈值 |
| `R18-原子写：` | 只读目标（异步 + 同步）、同进程句柄、跨进程 `FileShare.None` 真锁、重试窗口 ≥3 次退避 |
| `R18-schema：` | `output.schema.properties.content`、五形态 content 仍为字符串、17 种非法元素的下标定位、稀疏数组空洞、不误伤（`value` 数字、齐全字段） |
| `R18-大图：` | 2000×2000 噪声 PNG（11.45 MiB）内嵌、`store` 落地、`office_read` 回读、远低于限额、字节账 |
| `R18-基准：` | 8 MiB × 4 种输入不栈溢出/耗时上界、无元音连排口径与旧参考实现一致、点前导口径一致 |

**红线与约束**：零第三方依赖（只加 Node 内置用法）；没有调用任何外部 Office 转换器；
没有碰宿主装配（`cordis.patch.yml` 等）；测试产物全部落 `%TEMP%`（唯一的例外是 §4 的
`nonTempWritableDir()` 候选目录，且**刻意排除插件目录自身**以免污染目标目录的 `test-out/`）。

**唯一的不兼容点**（`office_edit` 参数层收紧）：非字符串的 `find` / `replace` / `text` / `markdown` /
`cell` / `name` / `newName` 由"静默字符串化"改为**参数错误**（原因：运行时与早就声明 string 的 schema 对齐；
迁移：传字符串；影响：只在调用方传了数字/对象时可见，字符串调用逐字兼容）。

**剩余风险**（未验证或验证不充分的部分，如实列出）：

1. **>4 GiB 数据尺寸分支未能端到端**：分配 4 GiB 以上缓冲不可行，该分支只由
   `zip64ExtraField` 的布局单测 + `baseOffset` 接缝（覆盖偏移分支）间接覆盖。**ZIP64 的"数据 > 4 GiB"
   假设仍未被真实产物验证过。**
2. **`Expand-Archive` 真解包只跑到超时**（54,357/65,536 文件，无错误）——"能解开"是实测，
   "最终产出 65,536 个文件"没跑完（套件默认走 `.NET ZipArchive` 等价验证）。
3. **Windows 文件替换行为**仍是经验性的：只读目标与 `FileShare.None` 真锁都实测报 `EPERM`，
   但**杀毒 / 同步盘 / 索引器**制造的瞬时锁只能靠退避重试缓解，本轮没有模拟这一类（探针用的是确定性锁）。
4. **网络盘 / OneDrive 上的 fsync 与 rename 语义未验证**（`DSH_OFFICE_ATOMIC_FSYNC=0` 只是开关，
   没有网络盘实测）。
5. **CRC 误拒面**：选定"严格拒绝 + 逃生口"，因此**确实存在**"包本身可用但 CRC 与数据不符"时被拒的场景；
   这类包必须显式设 `DSH_OFFICE_ZIP_CRC=0`（错误文案已点名该变量）。
6. **`textQuality` 栈溢出的精确阈值未定**（只在 (1 MiB, 8 MiB] 区间内确认），修复方式本身是新写法、
   不再依赖正则，但"其他输入形态下还有没有别的超大匹配"只能靠持续补基准用例。

**同步回目标目录 + 在目标目录再跑一遍**（这一次用 `danger-full-access`，因为目标目录在会话工作区之外）：

```
# ① 同步前：目标目录全量 SHA-256 清单（176 个文件，含 vendor 二进制）
work/r18-run/sha256-before.txt      # 176 行
# ② 备份将被修改的 7 个文件 → work/r18-run/backup-before/
# ③ 逐文件复制 7 个文件（每个都核对 src/dst SHA-256 相同）
# ④ 同步后：work/r18-run/sha256-after.txt
→ 变动文件恰为 7 个（index.js / zip.js / test.mjs / README.md / SKILL.md / DEVELOPMENT.md / CHANGELOG.md）
  新增 0、删除 0、其余 169 个文件哈希逐字不变
# ⑤ 在目标目录再跑一次
cd ${DSH_HOME}/plugins/dsh-office
$env:DSH_OFFICE_TEST_OUT = Join-Path $env:TEMP 'dsh-office-r18-target'
& <node> test.mjs        → 724 checks / ALL PASS（exit 0）
& <node> repro.mjs       → 🟢 GREEN（exit 0）
# ⑥ 跑完再取一次清单：176 个文件哈希与 ④ 完全一致 ⇒ 测试没有污染目标目录
```

目标目录那次比工作副本多 5 项检查，原因是**受限沙箱下被跳过的两类现在真的跑了**：

- **非 `%TEMP%` 目标目录的 PDF 产出（4 项，全部 PASS）**：
  `nt=<home>/.dsh\tmp\dsh-office-r18-nontemp`，
  create → `{"embedded":true,"renderCheck":"pass","firstPageBytes":15103,"images":0}`、
  convert → `firstPageBytes:15010`、空白夹具仍被拒且不落盘、目录里不留临时件/渲染副本。
  **这就是任务 A 的正面验收**（旧版在这条路径上 100% 报"第 1 页渲染失败"）。
- **原生 Office COM 冒烟真的开成功了**：
  `开箱：Word 16.0 能打开产出的 .docx`（`paragraphs=3`）、`Word 能把产出的 .docx 导出 PDF`（186,699 字节）、
  `Word 能打开产出的 .html`（`paragraphs=9`）、`R13-1 开箱：含图 .docx`（`inlineShapes=2`）、
  `R13-7 pptx：PowerPoint 16.0 能打开产出的 .pptx`（`slides=3`），
  以及**负向控制仍然 FAIL**（`notesMaster` 改回共用 `theme1` → `0x80070570`）。
  ⇒ 本轮改动**没有破坏任何真实 Office 打开能力**。

> 文档（README / SKILL / DEVELOPMENT / CHANGELOG）在那次运行之后又改过，但 `index.js` / `zip.js` / `test.mjs`
> 三个**代码文件**的哈希与运行时刻完全一致（`F46947CD77A5…` / `9412D286A36F…` / `3EA85520CB50…`），
> 所以上表那 724 项检查覆盖的就是最终交付的代码。

---

## 第十七轮（R16）：缓存身份 / 原子落盘 / ZIP 限额 / schema 一致性

> 对象：`${DSH_HOME}/plugins/dsh-office/`。本轮是**审计驱动的加固**（不是新功能），
> 每一项都补了回归测试；实际命令、通过与跳过项见本文件末「交付报告」。

### 1. OCR/读取 sidecar 的**缓存身份**（最高优先级）

病灶（两条都是静默的）：

- `.ocr.md` 只校验 `parser:` 版本，**不绑定源内容** —— 同一路径的 PDF 被替换（另存为 / 重新扫描 /
  覆盖下载）之后，只要 parser 没变，旧 OCR 文本照旧被贴回来。
- 集中缓存目录（`DSH_OFFICE_CACHE_DIR`）与临时回退目录按 **basename** 命名 —— `D:\a\第1章.pdf` 与
  `D:\b\第1章.pdf` 共用同一个 `第1章.ocr.md`：轻则互相命中，重则互相覆盖。`.read.md` 与渲染目录
  `renderDirFor()`（页面 PNG）同病 —— 后者更危险：会把**另一份文档的页面图**喂给 OCR。

修法：

- 身份 = **规范化源路径指纹**（`sha256(lowercase(realpath))`，纯字符串运算、零 IO）+ **源内容 SHA-256**。
- 内容哈希**优先复用解析时已经读进内存的 `buf`**（`loadModel` 的 pdf 分支 →
  `registerSourceIdentity`）；没有 buf 时按 1 MiB 分块同步哈希（`sourceIdentityOf`，带
  `路径|size|mtime` memo），**不对大文件重复读盘**，也不只依赖 mtime/size。
- 缓存文件名：同目录 sidecar（`<file>.ocr.md`）**逐字不变**；集中目录 / 临时回退改成
  `<name>-<pathKey8>.ocr.md`，`.read.md` 同理，渲染目录改成 `<safe>-<pathKey8>-<mtime>`。
- manifest 新增**独立一行**身份（covered 行逐字不变，旧调用方 `split('\n')[1]` 的断言不受影响）：
  `<!-- srcpath: <64hex> | srcsha256: <64hex> | srcsize: N -->`。
- 读取时逐项核对，任一不符**整份作废**，原因写进 `stats.ocrCacheStale.reason`：
  `parser` / `parser-missing` / `identity-missing` / `path-mismatch` / `content-mismatch` /
  `unverifiable`。旧命名（无指纹）的历史缓存会被**发现并明确作废**（`staleCacheNote` 说清是哪个文件、
  为什么），绝不静默复用。
- `covered` / `src` / `retry` 的跨批累积语义原样保留（有回归用例守着）。

### 2. 创建 / 转换 / 编辑 / 缓存写盘一律「临时件 → 成功后替换」

`saveBuffer` 旧实现是 `writeFile(file, bytes)`（flag `'w'`，**open 那刻就截断**），而
`office_edit` 的落盘点 100% 覆盖原件 —— 写盘中断就是"改坏原件"。现在统一走：
目标**同目录**唯一临时件（`open(..., 'wx')` 的 O_EXCL 保证并发唯一）→ 完整写入（默认 fsync，best-effort）
→ `rename` 发布。

三条 Windows 实测约束（本轮探针实测）：

- `rename` **能覆盖已存在的目标** → 正常路径直接 rename，不预先删目标；
- 目标上有任何句柄时 rename 会 EPERM/EBUSY → **有限退避重试（60/120/180 ms）**，仍失败就**保留原目标**并抛
  「写盘失败｜四要素」；**绝不 unlink 目标**（那会绕过用户的只读保护、制造"目标短暂消失"的窗口，
  而真正的文件锁下 unlink 同样 EBUSY）；
- 跨卷 rename 是 EXDEV → 临时件必须与目标同目录，**绝不能用 `tmpdir()`**。

`pdfOutputGate` 改成复用同一套两段式（删掉它自己的 stamp 命名），**不叠加两层**；
`renderCheck:'skipped'` 分支经由 `saveBuffer` 自动获得原子性。sidecar 用同步版
`writeFileAtomicSync` 保住"同步返回、失败 undefined"的既有契约。临时件以 `.part` 结尾，顺带解决
"残留临时件被 `office_read paths=[目录]` 当成用户文档盘点"的旧小坑。

### 3. ZIP 解压资源限制与结构校验（zip.js）

`MAX_FILE_BYTES`（200 MiB）只约束压缩包字节数。本轮实测放大比 294×（合法 xlsx XML）~1026×（deflate 零串）：
33 KB 的"docx 形态"文件能让 `office_read` 吃进 +129 MiB，261 KB 的吃 +512 MiB；分散形态的炸弹甚至**不报错**。

- 新增 `ZipView` 有界读取器：中央目录 / local header / 条目偏移 / 压缩长度 / ZIP64 尺寸全部做边界校验，
  越界一律**可读中文错误**（不再 `RangeError: Offset is outside the bounds of the DataView`）。
- 修掉两个真 bug：① 旧第 90 行 `if (localOff === 0xffffffff) { /* localOff = */ … }` 的**赋值被注释掉**，
  完全合法的 ZIP64 包 100% 读不了；② ZIP64 分支多写了 `&& cdCount === 0`，真实 ZIP64 永远进不去。
  另外补上：stored 条目"静默返回 负载+中央目录+EOCD 拼接物"、nameLen 说谎导致名字变垃圾、重名静默覆盖。
- 解压前按**声明尺寸**挡两道上限（单条目 / 整归档累计），解压时用 zlib 的 `maxOutputLength` **早停**
  （实测超限时在流未解压完即中止，RSS 只多 1 MiB），解压后核对实际长度 == 声明值；
  CRC32 默认校验（`crc !== 0` 时才比对），`DSH_OFFICE_ZIP_CRC=0` 关闭。
- 阈值与依据：单条目 **256 MiB**（实测最大合法条目 17.31 MiB 的 14.8×、最大 stored 图片条目 11.45 MiB 的 22×，
  按 908 B/行覆盖约 29.6 万行工作表）；累计 **512 MiB**（合法整包实测最大值的 ~30×）。
  **刻意不做压缩比判据** —— 图片条目压缩比恒为 1.00×，按压缩比拒绝只会误杀正常文档。
  `DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` / `DSH_OFFICE_ZIP_MAX_TOTAL_BYTES` 可覆盖，`0` = 关闭该上限。
- 写侧 `makeZip` 补守卫：条目数 > 65535、条目名 > 65535 字节、数据 / 偏移 > 4 GiB 一律**报错**，
  不再 `Math.min(...)` 静默丢条目、或写 0 哨兵产出废包。
- 同类连带修复：`textQuality()` 的 `s.match(/[A-Za-z]{6,}/g)` 在**单个 8 MiB 连续字母串**上会栈溢出
  （那是合法输入）→ 改成 `matchAll` 惰性迭代。
- 验证：17 个攻击用例（越界 / 说谎 / ZIP64 / 伪 EOCD / CRC）全部转为可读错误**或正常读取**；
  **333/333 份本机真实 Office 文档逐条目长度 + SHA-1 完全一致、0 误拒**。

### 4. 工具参数 schema 与运行时行为一致

- `grid`：实现（`normalizeGrid`）一直支持数字与对象，schema 却限定字符串 → 数字形式**完全不可用**
  （`参数 "grid" 应为字符串`）。现在 schema 用 `oneOf: [{type:'string'},{type:'number'}]`
  （host 的 JSON Schema 子集明确支持 `oneOf`，而 `type: [...]` 数组是**被拒绝**的写法），
  validator 同步放宽；字符串形式逐字兼容。
- 顺手修掉同源缺陷：`jsonSchemaOf` 把 `integer` 降级成 `number`，而 validator 要求整数 ——
  `office_read.offset/limit/pageFrom/pageTo` 的 schema 现在就是 `integer`。
- `office_edit.operations.items`：补 `op` 枚举（16 个实现真正支持的取值）+ `required:['op']` +
  常用字段描述；运行时新增 `itemViolations`，非法元素报 **`operations[N]` 下标定位**的清晰错误
  （`{}`、`['replace_text']`、缺 `newName` 的 `rename_sheet` 过去会被静默放到执行期，
  甚至写出 `Sheet1 → undefined`）。容器相关的必需性仍由执行期精确报错；
  **不写 16 支 oneOf**（schema 体积与歧义都不划算）。
- `width` 描述修正为实现的真实规则：**省略 / 0 / 负数 / NaN 时没有固定默认值**，按
  `原图宽px × 72/96` 换算，高度按纵横比，超过 A4 可用宽 **451.3pt** 时等比缩到 451.3pt 并记账
  `capped=true`（旧文案的"默认 450"在实现里不存在）。
- `as="json"` 描述写明 **`content` 是 JSON 文本字符串**（不是对象，调用方需自行 `JSON.parse`），
  返回结构**未改**（4 处调用方依赖 string）。

### 5. 兼容性影响（本轮唯一的不兼容点）

集中缓存目录 / 临时回退目录下的缓存文件名新增路径指纹后缀，旧文件**不再被当作命中**
（这正是不兼容契约要求的"缺身份字段一律作废"），但会被**发现并明确作废**：
`stats.ocrCacheStale.reason = "identity-missing"`，`staleCacheNote` 给出路径与原因。
迁移方式：重新识别需要的页（旧文件可直接删除）。同目录 sidecar 的命名**未变**。

### 6. 交付报告（本轮实际执行的命令与结果）

**基线**（改动前，同一环境、同一台机器）：

```
cd <工作副本 r16-src>
$env:DSH_OFFICE_TEST_OUT = Join-Path $env:TEMP 'dsh-office-r16-baseline'
& <node> test.mjs            # <node> = <home>/.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe
→ 586 checks / 2 FAILED
```

**改动后**（最终一次，代码与测试都已定稿）：

```
cd <工作副本 r17-work>
$env:DSH_OFFICE_TEST_OUT = Join-Path $env:TEMP 'dsh-office-r16-final4'
& <node> test.mjs
→ 656 checks / 2 FAILED（exit 1）
```

- 新增 **70 项检查**（586 → 656），全部通过。
- **2 项 FAIL 与基线完全相同**，是同一对用例：`R13-2 判别实验（WinRT 打开同骨架 AES-128 夹具）` 与
  `R13-2 判别实验：探针本身有效`。根因是本机环境限制：沙箱内 WinRT 的
  `StorageFile.GetFileFromPathAsync` 只允许 `%TEMP%` 下的文件（实测：TEMP 下的 PDF → `STORAGE OK` +
  `LOAD OK`；工作区与 `~/.dsh` 下的 PDF 一律"拒绝访问，该项目没有位于应用程序可以访问的位置"）。
  该用例的探针 PDF 不在 TEMP，所以 `plainPng=0` → 两条都 FAIL。**与本次改动无关**，改动前后一致。
- **跳过项**（测试会自己打印原因，不计为通过，与基线一致）：Word / PowerPoint COM 冒烟（本机 Office
  不可用时 skip）；`DSH_OFFICE_TEST_AES256_PDF` 真实样本槽位未设 → skip；若干"真实文件存在性"用例
  因样本文档不在本机 → skip；多份 `DSH_OFFICE_TEST_*` 夹具槽位未设 → 走各自的 skip 分支。
- **另一路自检**：`node repro.mjs` → 🟢 GREEN（退出码 0，返回值无损可序列化）。
- **独立验证（不在 test.mjs 里）**：新 `zip.js` 对 17 个攻击用例（越界偏移 / 说谎的长度 / ZIP64 /
  注释内伪 EOCD / 坏 CRC / 重名 / 截断）全部给出**可读中文错误**或**正常读取**；
  对本机 **333 份真实 Office 文档**逐条目长度 + SHA-1 与新读取器比对 **完全一致、0 误拒、0 条目被上限命中**。
- 边界值实测：新读取器能读出 4 MiB 高压缩比条目与 11.45 MiB stored 图片条目；声明尺寸说谎的条目
  在**解压前/解压中**就被挡下（`maxOutputLength` 早停，超限时 RSS 仅 +1 MiB）。

**本轮新增的故障注入 / 调参环境变量**（默认都不设，行为不变）：

| 变量 | 用途 |
| --- | --- |
| `DSH_OFFICE_ATOMIC_FAULT=temp-open｜temp-write｜publish` | 原子写的三个故障点（测试专用的注入钩子，只在 helper 内判断） |
| `DSH_OFFICE_ATOMIC_FSYNC=0` | 关闭原子写的 fsync（网络盘 / OneDrive / 批量场景） |
| `DSH_OFFICE_ZIP_MAX_ENTRY_BYTES` | 单条目解压上限（默认 256 MiB，`0` = 关闭） |
| `DSH_OFFICE_ZIP_MAX_TOTAL_BYTES` | 整归档累计解压上限（默认 512 MiB，`0` = 关闭） |
| `DSH_OFFICE_ZIP_CRC=0` | 关闭 CRC32 校验（默认开启，仅当 CD 里 `crc !== 0` 时才比对） |

## 修复注记（2026-09-24 · 第七轮：docx 插图两处真 bug / 图片解码面 / 公式子集扩展 / AES-128 对照）

| # | 需求 | 结果 |
| --- | --- | --- |
| 1 | **docx 写出端 image 分支**（含图 md/HTML → docx） | `writeDocx` / `docBlocksToXml` 支持 `image` 块：`word/media/imageN.<ext>` + image 关系 + `w:drawing`；内容 SHA-256 去重；尺寸规则显式化并写进 `notice`；`IMAGE_EMBED_TARGETS` 加入 `docx`（`.odt`/`.pptx` 继续走降级兜底 + 记账） |
| 2 | **含图 docx 被 Word 16.0 拒开的真 bug（第六轮遗留）** | 根因 = `docx.js` 的 `WP_NS` 命名空间 **URI 拼错**（`…/drawingWordprocessingDrawing/2006/main`，正确 `…/drawingml/2006/wordprocessingDrawing`）→ Word 解析 `wp:inline` 失败即判包损坏。docx 的插图链路此前只验过 zip 结构/`rId`/包体检，**这三项全都查不出 URI 拼写错误** |
| 3 | **`.pptx` 被 PowerPoint 16.0 拒开（第五轮遗留）** | 根因 = `notesMaster` 的主题关系指向与 `slideMaster` **共用**的 `theme1.xml`；补 `ppt/theme/theme2.xml` 部件 + `[Content_Types].xml` Override + 改 `notesMaster1.xml.rels` 指向，PowerPoint 16.0 能开（含"改回共用必须仍 FAIL"的负向控制） |
| 4 | **AES-256 读取可信度收口** | 新增 **AES-128（R4/V4 + AESV2）对照夹具**（与 AES-256 夹具同一骨架）做**判别实验**；新增真实样本槽位 `DSH_OFFICE_TEST_AES256_PDF`（未设则 skip 并打印原因，主人有样本时一条命令复验） |
| 5 | **公式重算扩子集 + 护栏** | 新增 `SUMIF / AVERAGEIF / COUNTIFS` 与 `LEFT / RIGHT / MID / LEN / TRIM / UPPER / LOWER / CONCAT / CONCATENATE / TEXT / VALUE`；区域里的文本/布尔按**真 Excel 口径忽略**、直接参数仍换算；整列引用 `A:A`、外部工作簿、`_xlfn.`、未加载工作表、循环引用全部显式记账（带地址）；`DSH_OFFICE_RECALC_MAX_CELLS` 大表护栏 |
| 6 | **图片解码面 + 去重 + 对齐** | PNG 索引图（`PLTE`/`tRNS`）、BMP（8/24/32-bit）、GIF 首帧（透明索引 + 交错行序）都能进 PDF；媒体按内容哈希去重；PDF 图片支持 `align: center/right`（缺省仍是 `marginX`，字节不变） |
| 7 | **可选件（只交草案 / 结论）** | PDF 合并·拆分、`.html` 结构化编辑、插图落点（`in_table` / 页眉页脚 / `before`）、`recalc_formula` 写回 —— **全部只交接口草案**（红线要求先报告）；OMML 公式写入结论为**不做**（附量化依据与重开条件） |

诊断留档：`work/office-enhance2/probe/`（单变量变体生成器与对照脚本）、
先报告清单与接口草案：`work/office-enhance2/DRAFTS-r13.md`。

## 修复注记（2026-09-15 · 第一轮：返回边界根治 + 读取降级链）

**问题**：含正文的 `office_read` 对部分 PDF 整条返回 `value is not lossless JSON`，正文全丢；`as="meta"` 正常。
DSH 侧实测更严重——**六条调用矩阵全灭**（整本 / 单页 / `ocr=never` / 本地 OCR / 视觉 OCR / 纯扫描页），
连"单页、文字层干净"这条最窄路径也过不了，说明元凶是**全局返回结构**，不是某页乱码。

**根因**：host 侧在工具体返回值离开 `execute()` 之前做无损 JSON 判定（DSH 走
`@deepseek-ai/dsh-util-values` 的 `snapshotJsonValue`），**字符串内容不参与判定**——元凶是
`stats.ocrCovered: undefined` 这类"有则给、无则 `undefined`"的属性写法（判定器把 `undefined` 属性
视为不可序列化 → 整条作废）。第一轮在 DSH 侧这段代码**存在但从未接线**（`finalizeToolValue` 定义了没人调用）。

**补丁位置**：`index.js`「返回边界」段（`sanitizeTextForReturn` / `losslessJsonProblem` /
`losslessValueOf` / `finalizeToolValue` / `sanitizeThrown`），并**在 `defineToolLite()` 的 `execute` 返回、
`render` 投影、`catch` 三处全部接线**，作为唯一出站终点。返回 JSON 新增 `stats.sanitized`；
正文替换了码点时尾部加脚注"N 个非法码点已替换为 U+FFFD"。
同时把读取做成 fail-safe 降级链（文本层 → 质量门 → 自动本地 OCR → sidecar `<文件>.read.md`），
错误分支统一带四要素（页码 / 格式 / 根因 / 可复制的下一步）。

**自检**：`node repro.mjs` 是返回边界的常驻哨兵（🟢 GREEN 即 host 一定收，退出码 0/1）。
本轮实测：**同步前 🔴 RED（`$.stats.ocrCovered 是 undefined`，且是在干净对照样本上复现）→ 接线后 🟢 GREEN**。

## 修复注记（2026-09-15 · 第二轮：四类静默失败）

**问题**：① `convert` 把 CID 乱码 PDF 的文字层转成含 NUL 的二进制垃圾落盘（实测 56,341 字节 /
NUL 5135 个 + C0 控制符 13666 个），返回 JSON 零告警；② `as=meta` 对乱码文字层"看起来完全健康"
（`pagesWithText=30` / `scannedPages=[]`）；③ 质量门失败的页直接丢掉，把"丢一页"留给人工，
而且**调用方完全看不到**；④ 截断协议要求消费方 `slice(0, nextOffset - offset)` 才能拿到真正文，写错就"又重又漏"。

**补丁位置**：

| 任务 | 位置 | 行为 |
|---|---|---|
| 出站质量门 | `office_convert.execute` + `convertSourceGuard` / `convertRefusalError` | 写盘前判文本层质量；乱码页有完整 `.ocr.md` 则替换后继续（`fallback="ocr"`），否则**拒绝落盘**并抛四要素，绝不产出目标文件 |
| meta 质量画像 | `readPdf` 的 `as=meta` 分支 + `textLayerProfile` / `garbledSpec`（纯 CPU） | 新增 `textLayerUsable` / `garbledPages` / `qualityGate`；命中乱码时 `suggestion` 改为 `{note, copy}`（可复制的 OCR 参数串）。既有字段不改名不改值 |
| 换倍率重试 | `ocrPdfPages` 本地批次 + `ocrRetryScales` / `currentRenderScale` / `scaleTag` / `renderDirFor(file, tag)` | 没过质量门的页按 `DSH_OFFICE_OCR_RETRY_SCALES` 串行换倍率重试，取"过关且置信度最高"的一次 |
| 缓存位置可配 | `cacheDirState` + `ocrCachePath` / `readSidecarPath` | `DSH_OFFICE_CACHE_DIR` 生效；未设行为逐字不变；不可写则回退并把原因写进 `stats.cacheDirNote` |
| 截断协议收敛 | `capWithOffset` / `finalizeToolValue` 护栏 | `content` 只放纯前缀，说明移入 `notice` + `stats.truncateNote`；硬不变式 `offset + content.length === nextOffset` |

干净文件（含 43 页对照样本 `sample-B.pdf`）四条路径 read / meta / convert / repro 的
既有字段与正文逐字不变；新增字段只在命中设计条件时出现。

## 修复注记（2026-09-15 · DSH 专属补充：失败通道可观测性）

DSH 宿主实测（F2/F3/F7）暴露出三条 WorkBuddy 侧不存在或相反的行为，本轮在 DSH 侧补上：

| 补充 | 位置 | 行为 |
|---|---|---|
| 拒绝文本自带 sidecar 现状 | `convertSourceGuard` / `convertRefusalError` | 除四要素外再给 `sidecar=<绝对路径>（covered: X / total: N）`、**仍缺哪几页**、`grep 'covered:'` 提示；`下一步` 的 `pages` 只列**缺页**（不叫已识别过的页重跑） |
| 读取抛错分支同样自带 | 新增 `sidecarCoverage()`，接进 `readErrorHint` | F2 的账：宿主里 sidecar 对"会话只读"的附件目录**写得进去**，所以"调用失败但成果已在盘上"是常态。现在错误文本直接给出路径 + covered + 缺页，不必再靠人去 grep 自救 |
| 批内缺页显式外显 | `stats.ocrFailedPages`（页码 + 一句话原因）+ 正文脚注「批内缺页：…」 | F3 的账：旧版某页没过质量门 → 返回值里看不到、错误里没有。现在数组常驻（空数组 = 本批无缺页），不 grep 也能看到缺哪页、为什么缺 |
| sidecar 落**全文** | `readPdf` / `office_read` 经临时字段 `__fullBody` 传给 `finishRead` | 旧版把被内联护栏截过的**前缀**写进 sidecar，而 notice 承诺"整篇正文"——那是假话。现在 sidecar 恒为全文，另加 `stats.sidecarChars` 可核对 |

**行为边界**：以上全是"纯兜底 / 只增字段 / 只加文本"，不改任何既有参数语义；干净文件的返回值逐字不变。

## 修复注记（2026-09-15 · 第三轮：pdf.js 解析层两处先存缺陷）

**病灶**（第二轮"任务七探路"发现，本轮实测复核并修掉）：

| # | 位置 | 缺陷 | 后果 |
|---|---|---|---|
| A | `pdf.js` `expandObjStms` | 头部索引 `nums` 是从 `text.slice(0, first)` 读出来的（`off`/`end` 都是**区域内相对偏移**），却被当成 `text` 的**绝对**下标 | 每个 ObjStm 内对象整体前错 `First` 字节：第一个对象恰好读到 ObjStm 的**头部索引表本身** → `/Pages` 解析不出 `/Kids` → `pages()` 落到"扫描全部 `/Type/Page`"的兜底（页序变对象号升序）；其余对象整体错位 → 页 `/Resources` 错位 → `fontsOf` 空表 → 2 字节 CID 被拆成两个单字节字符 = **"整本中文乱码"的真因** |
| B | `pdf.js` `mul()` | e/f 两行有三个独立错误：读越界的 `b[6]`、该用 `a[4]/a[5]` 处写成 `b[4]/b[5]`、以及 `0 * undefined = NaN` | **连纯平移都炸**（`mul(ident, [1,0,0,1,5,7])` 的 e/f 都是 NaN）→ 每个 run 的 y 都是 NaN → `Math.round(NaN/2.5)` 是合法 Map key → **每页所有 run 塌成一行** |
| C | `pdf.js` `runPara` | 换行循环把 `fit` 夹在 ≥1；一旦"本行剩余宽度装不下**一个**字符"，判定条件恒为真 | **同步死循环**，事件循环被占死（`Promise.race` 连超时都触发不了，整个宿主进程冻住）。触发条件**与字符种类无关**：私用区最先撞上，汉字/扩展 B/`_`/ASCII 长串一样会挂 |

**A、B 必须同批**：只修 A 时"看着有换行"是乱码里混进行分隔符的假象，单独合入会让病灶样本"看起来更坏"（换行 718 → 0）。

**连带修复**（同一批，否则修完反而暴露新问题）：

- `cm` 的合成顺序：规范是 `CTM_new = M_cm × CTM_old`，旧版写成 `mul(a, ctm)`（反了，带缩放的嵌套 `cm` 会算错）。
  同时补上 **q/Q 的图形状态栈**（旧版 q/Q 只清操作数栈，CTM 整页单调漂移永不复位）。
- `Td` / `T*` / `'` / `"`：旧版 `Td` 把上一个 `Tlm` **整个丢掉**（连续 `Td` 的多行文本全落在同一点）；
  现在按规范 `Tlm_new = T_translate × Tlm_old` 累积，且位移发生在**文本空间**（旋转行也对）。
- `Tz / Tc / Tw / Ts / Tr` 真正读取：旧版按"未知算子"清栈，状态永远读不到。`3 Tr` 是**不可见**文字，
  不生效就会把水印/隐藏层当正文。
- `pushRun` 存**设备空间的有效字号**（`|Trm|` 纵向量长），不再存裸 `Tf`（某样本 209 vs 真实 10.45pt，
  宽度估算差 20 倍）。
- 行内 run 之间补空格改为 **CJK 感知**（两侧有一侧是 CJK 就不补），旧规则只对恰好以 `一` 开头的 run 特判。
- `as="json"` 两处收口：PDF 的 json 分支过去**忽略 `pages`**，且超 `READ_CAP` 时 `capText` 会在
  **字符串中间**切断 → 返回**无法 parse 的半截 JSON**。页序修好之前碰不到，修好后立刻暴露。
  现在 honor `pages`，仍超限就**按页丢尾部**并写进 `stats.truncateNote`（永远是合法 JSON，绝不静默丢）。
- **缓存迁移（`parser:` 版本戳）**：`.ocr.md` 按页码记，而页码来自插件自己的 `pages()`；
  解析器一旦改变"页码 → 哪张纸"的映射，旧缓存就会把**另一张纸**的识别结果贴上这一页（静默错页）。
  现在 sidecar manifest 带 `parser:`，**缺失或不符 = 整份视为未覆盖**，`stats.ocrCacheStale` /
  `stats.ocrCacheNote` / 正文脚注 / convert 拒绝文本都会说明原因。
- **质量门"点前导"假阳性（第六处连带修复）**：目录页整页是 `第一章总论..................1` 这类
  **点前导 + 页码**，逐页 CJK 覆盖率天然极低（① 第 4 页 **12.2%** vs 全书 **73.8%**），会撞上
  "相对 CJK 覆盖率"判据 —— 这是**假阳性**：`ocr="auto"` 白跑一次目录页 OCR，`office_convert`
  更会因"第 4 页文字层不可信"**拒绝整本**（明明 34 页都能读）。
  现在**连续 ≥3 个 ASCII 句点视为排版装饰、不计入可见字符**：该页覆盖率回到 **79.7%**、
  整本 `garbledPages=[]`、`textLayerUsable=true`。
  **不是放宽闸门**：真缺 `ToUnicode` 的文档表现为 **PUA / 替换字符成片**，不会长成"一片点号"；
  私用区 / 替换字符 / 控制符三条判据一字未动，仍各自独立命中（纯点号页也不会被判乱码）。
  ⚠ 这条**故意偏离了"不要动质量门判据"的红线**，属"先报告再动手"；前后对照证据见 `CHANGELOG.md`。

**实测（本轮亲自复核，非采信探路报告）**：

| 样本 | 指标 | 修前 | 修后 |
|---|---|---|---|
| ① `sample-A.pdf`（35 页） | 字符 / 控制符 / CJK / 填空段 / 换行 | 45,692 / 19,085 / **0** / 0 / 718（假） | 23,992 / **0** / **16,141** / **305** / **1,084** |
| ① 的 `as="meta"` | `textLayerUsable` | **false** | **true** |
| ② `sample-B.pdf`（43 页） | 字符 / CJK / 填空段 / 换行 | 31,200 / **19,879** / **700** / **0** | 31,200 / **19,879** / **700** / **1,738** |
| ③ 私用区/汉字/扩展B/`_`/ASCII × 长度 1..200 | `writePdf` | ≥47 字符**永不返回** | 全部 1–2 ms 返回 |

② 的"字符数"必须拆开看：**换行 0 → 1,738** 是修复本身（旧版每页塌成一行），
而**去换行后 29,400 → 29,420（+20）**，增量全是**行分隔 + markdown 标记**（`# ` 行 7 条）；
扣掉标记与换行后只多 **6** 个字符（`# ` 之后的空格与 run 间确有空隙时补的空格）。
**汉字 19,879 与填空 700 段逐字不变**。
（旧版基线 29,400 是**用备份的旧 `pdf.js` 重新实测**的：total 29,442 / 换行 42。）
⚠ 这**故意打破了第二轮"干净文件逐字不变"的红线**，是修复不是回归。
（1,738 是**逐页内**的换行数；用例 `断行：② 换行总数…` 数的是逐页正文再拼接后的换行，
含 42 个页间连接符，所以它报 **1,780** —— 两者口径不同，不是打架。）
⚠ 这**故意打破了第二轮"干净文件逐字不变"的红线**，是修复不是回归。

**① 的正文口径**：上表的 **16,141** 是 `probe.mjs` 直接数解析出的**原始文本层**；
用例 `解析器：① 修后 CJK > 16,000` 数的是 `as="markdown"` 的**落盘正文**（含页标记、脚注），
本轮实测 **16,311**。点前导修复前后该正文的差异已逐行核对：**唯一差异是第 4 页那两行质量门脚注
（80 个汉字）及其附属的 2 个空行**，剔除脚注与空行后正文 **916 行逐字一致**
（`out/ab-diff2.txt` / `out/ab-diff3.txt`）。

**自检**：`node test.mjs`（乱码端到端改用**自造夹具** `buildGarbledFixture()`，零样本依赖；
另加 `解析器：` `ObjStm：` `矩阵：` `死循环：` `断行：` `填空：` `缓存迁移：` 七组新用例）。
本轮实测 **378 checks / ALL PASS**（**当轮基线**；准确数量看最后一行的 `N checks`），`node repro.mjs` 🟢 GREEN（退出码 0）。

## 修复注记（2026-09-16 · 第四轮：读取链路修补——HTML / 批量盘点 / 批间衔接 / 字符级质量启发 / CSV 编码）

> 来源：一次重度实战反馈（25 个中文 PDF + 17 个 HTML 学习笔记的读取会话，含 meta 探测、
> 4 批分页读取、扫描件判定、大文件分批规划）。四项中优先级 1、2 → 3、4 → 5 依次交付。

| # | 需求 | 交付 |
| --- | --- | --- |
| 1 | **HTML 读取与转换**（高优） | 新增 `html.js`（**零第三方依赖**：手写分词器 + 隐式闭标签树构造器 + 文档模型遍历器，不用正则剥标签——正则会丢嵌套列表语义）。`office_read` 支持 `.html`/`.htm`：`as="markdown"` 输出结构化转换（h1-h6→#、嵌套 li→带缩进 -、table→管道表、blockquote/pre/外链/加粗斜体保留、语义顺序不变）；`as="text"` 输出剥离 style/script 后的**纯文本**（表格→制表符行）；`as="meta"`/`as="json"` 走既有非 PDF 分支自动生效。HTML 实体统一解码（named ~200 项 + 数字 + **Windows-1252 数字别名**，`&#150;`→–）。`office_convert` source=.html → .md/.txt/.docx（及其他全部可创建格式）；`office_create from=*.html` 同样可用。兼容性：`.html` 旧版按"未知文本"返回**原始源码**，现改为返回解析后的正文——这是需求本身；其余格式行为逐字不变 |
| 2 | **批量文件 stats 探测**（高优） | `office_read` 新增 `paths` 数组参数（条目可为文件或**目录**，目录展开受支持扩展名并自动排除 `.ocr.md`/`.read.md` 等插件缓存）。一次调用返回 `{format:'scan', total, ok, failed, files:[…], content}`：行字段与 meta.stats 同名——`format`/`pages`/`characters`/`textLayerUsable`/`scannedPages`/`garbledPages`（数）+ `garbledPageSpec`（页码明细，仅乱码时）+ `suggestedBatches`（建议分批数，与 meta.suggestion 同口径：干净 PDF 15 页/批、乱码 20 页/批、文本类按内联上限）+ 行内 `error`。**绝不返回正文**；纯 CPU（PDF 复用解析 memo，无 OCR）。实测 23 个真题 PDF 一次调用约 1.1 秒。兼容性：`path` 从必填变为与 `paths` 二选一（都缺才报错）；`as="markdown"` 等与批量形态同给会显式报错而非静默忽略 |
| 3 | **跨批读取连续性**（中优） | 两条都做：① **页级续读** `pageFrom`/`pageTo`（含端点，等价 `pages="N-M"`，与 `pages` 互斥报错），`stats.pageFrom`/`pageTo`/`nextPage`（下一未读页）回显；② **批间衔接**——凡 `pages` 分批读取，默认在 `stats.prevTail`（上一批末页末行 ≈100 字符）/`stats.nextHead`（下一批首页开头 ≈100 字符）给批间上下文 + notice 摘要；`boundary:true` 时以 `> 〔dsh-office 批间回看/预览〕` 引用行内联进正文首尾（剔除规则固定，拼接可剥离）。**默认不改 `content`**：分批拼接与整本读取逐字 diff 成立（验收用例实测 5 页书 1-2+3-5 拼接 ≡ 整本） |
| 4 | **质量差的页主动降级提示**（中优） | `textQuality()` 新增 `structural` 模式（**只在 PDF 逐页质量门** `textLayerProfile`/readPdf suspect 循环启用；整本口径 finishRead/convert 探针不启用——合法重复表格行不会把整本推进 sidecar 兜底，误报不升）。5 条新启发，全部带行数/占比门槛：重复行率（≥10 行 ≥50% 逐字重复）、单字/双字行率（≥12 行 ≥75%，纯数字行不算）、Latin-1 高带占比（0xC0-0xFF ≥50%，mojibake/CID 错映射）、连续非词典字符（≥8 连排无元音字母串占字母 ≥25%）、非词典区相对异常（希腊/西里尔/制表符 ≥12% 且全书基线 <3%——**俄文原文档不误伤**）。触发页照旧计入 `garbledPages`/`qualityGate.reasons`，`ocr="auto"` 自动送去重识别并给可复制的 `ocr="always"` 参数串。`stats.quality` 增量加 `oddCharRatio` |
| 5 | **CSV/Excel 写出编码与转义收敛**（低优） | 核实：`.csv` 写出**本来就是 UTF-8 with BOM**、RFC 4180 转义（逗号/引号/换行字段自动加引号、内部引号成对双写）已具备——本轮把它参数化并补齐缺口：`office_create`/`office_convert` 新增 `encoding: "utf-8-sig" \| "utf-8"`（仅文本类目标 csv/tsv/md/txt，默认 csv 带 BOM、其余不带，其他目标显式报错）；**`office_edit` 修复**：编辑 CSV/文本不再静默剥掉源文件 BOM（旧行为"编辑一次、Excel 再打开就乱码"）。gb18030 写出因 Node Buffer 不内建该编码、非零依赖不可达，维持不提供（读取侧仍自动识别 GBK/Big5） |

**新文件 `html.js` 的热换说明**：`index.js` 以 `import … from './html.js?v=1'` 引入（与
`cordis.patch.yml` 的 `?v=` 爆破机制同款）。以后单独改 `html.js`，把这个 `?v=` 加一即可随
`index.js` 热换生效，不必等宿主重启。

**自检**：`node test.mjs` 新增 `R1:`-`R5:` 五组用例（HTML 夹具含 style/script/嵌套 ul>ol/表格/中文实体/
隐藏子树；批量盘点含 25 页乱码夹具的纯 CPU 行验证；5 页 PDF 的"分批拼接 ≡ 整本读取"diff；
质量启发的检出/误报双向断言；BOM 字节级断言）。本轮实测 **431 checks / ALL PASS**（**当轮基线**；准确数量看最后一行的 `N checks`），
`node repro.mjs` 🟢 GREEN；真实样本（两份 68-73KB 学习笔记 HTML、23 个真题 PDF 盘点）逐项过。

## 修复注记（2026-09-23 · 第五轮：中文 PDF 嵌字体 + 产出质量门 + Word 拒开根因 + .html 写出）

| # | 缺陷 | 根因（已定位到具体字节/部件） | 修法 |
| --- | --- | --- | --- |
| 1 | 中文 PDF 不嵌字体（STSong-Light + Adobe-GB1，`/FontFile=0` `/ToUnicode=0`）→ WinRT 渲染空白、Edge/Chrome 乱码、`office_read` 提取 `☆①` 变 `?` | 旧路径只用预定义 CMap，从不嵌字形 | 新增 **`pdffont.js`**：零依赖 TTF/TTC 解析（`simsun.ttc` 取 font[0]、cmap 0/4/6/12、hmtx、恒等 GID 子集：不重排 GID，未用字形 `glyf` 清零 + `loca` 重建长格式 + 重算 `checkSumAdjustment`）。`pdf.js` 改为 `Identity-H + FontFile2 + 完整 ToUnicode + /W`，`BaseFont` 带 `XXXXXX+` 前缀；`DSH_OFFICE_PDF_EMBED_CJK=0` 回退旧行为。**子集化顺带剔除 `COLR/CPAL/SVG `**（PDF 用不到，Segoe UI Emoji 的 COLR 单项 7.4 MB）——9000 字样本 **4.35 MB → 0.39 MB** |
| 2 | `office_create/convert` 对自己产出的坏 PDF 一路绿灯 | 没有出站质量门 | 新增 `pdfOutputGate()`：①字节级——正文含 CJK 却无 `/FontFile*` → hard-fail；②渲染级——写同目录临时文件 → `pdf-render.ps1` 渲染首页 → **整页黑像素 < 40 个**（主判据；辅助判据"文本 ≥24 字却 PNG <30KB 且墨迹 <0.1%"自第八轮起**仅对单页文档**生效，多页文档只走主判据 —— 避免"首页少字、后续页内容多"被误判；墨迹测不出时才退回"文本量可观却出图极小"）判空白 → 拒绝落盘、删半成品、抛四要素错误 + 逃生通道；通过才原子改名到目标。hr-only/仅分页标记不算"有可见文本"，不误杀；`DSH_OFFICE_PDF_SKIP_RENDER_CHECK=1` 只跳渲染级 |
| 3 | `office_create` 产出的 docx 被 Word 16.0 拒开（`0x800A1401`），同内容 odt 能开 | **`[Content_Types].xml` 里 `docProps/app.xml` 的 Override 类型写错**：用了 `…wordprocessingml.extended-properties+xml`，包级正确值是 `…openxmlformats-officedocument.extended-properties+xml`。单变量实测：只改这一处 → Word 打开；只把 app.xml/core.xml 内容换成 Word 自产的 → 照样失败（内容无辜，类型是根因）；`styles.xml` 的 `pPr` 顺序**不是**原因 | 改 `docx.js` 的 Content-Types 常量；`pptx.js` 同类错误一并改对 |
| 4 | 原生 Office 开箱没有任何回归防线 | — | `test.mjs` 新增 Word COM 冒烟：`Documents.Open` + `ExportAsFixedFormat($out,17)`（Word 不可用按套件惯例 skip）；另修 `html.js` 读取端 `<pre>` 只取直接文本子节点、导致 `<pre><code>…</code></pre>` 读出空代码块的问题 |
| 5 | 目标格式缺 `.html` | 只有读取端 `htmlToDocument` | `html.js` 新增写出端 `documentToHtml()`：语义化 HTML5（h1-h6/`<table>`+`<thead><th>`/嵌套 `ul`/`ol`/`blockquote`/`pre><code`/`hr`/行内强调），UTF-8 + `<meta charset>`，中文不转实体；接进 `WRITERS`/`KIND_TARGET`/工具描述 |

**自检**：`node test.mjs` **ALL PASS / 467 checks**（退出码 0；**467 是当轮基线**，之后逐轮增长 ——
准确数量看最后一行的 `N checks`，别写死），`node repro.mjs` 🟢 GREEN（退出码 0）。
本轮另用真实 Word 16.0 亲验：修复后的 docx `OPEN_OK`（无"发现不可读内容"）+ 导出 PDF 成功；
产出的 `.html` 也能被 Word 打开。

**✅ 已修（第七轮定位并修复：notesMaster 独立主题部件）** —— 原「尚未解决：`.pptx` 被 PowerPoint 16.0
拒开（`0x80070570`）」这条已作废。`office_create` 产出的 `.pptx` 当时的症状是 `Presentations.Open` 报
`0x80070570`「The file or directory is corrupted and unreadable」，而当时的 `test.mjs` 只用插件自带解析器
回读，所以**套件全绿也盖不住这条**。
**根因**：`ppt/notesMasters/_rels/notesMaster1.xml.rels` 把备注母版的主题指向与 `slideMaster` **共用**的
`../theme/theme1.xml`；PowerPoint 要求 notesMaster 有**自己独立的**主题部件（它自己保存时分配 `theme2.xml`）。
**补丁** = `pptx.js` 新增 `THEME_NOTES` 常量 + `ppt/theme/theme2.xml` 部件 + `[Content_Types].xml` 加 Override +
`NOTES_MASTER_RELS` 改指 `theme2.xml`。
**验证（正负向都在）**：正向 = PowerPoint COM `Presentations.Open` **真开成功**（`slides=3`）；
负向控制 = 把 notesMaster 主题改回共用 `theme1` **必须仍 FAIL**（实测 `0x80070570`）
—— 即"改对才开、改错必不开"，不是靠放宽条件蒙过去的。
**历史诊断链（一句话注记，保留修订轨迹）**：zip 层（自产 pptx 用我们的 `makeZip` 重打包仍能开）、
OPC 不变量（无缺失部件 / 无悬空关系 / 全部 XML 可解析 / 所有 `r:id` 都能解析 / 无重复 Id 与重复 Override）、
**所有非必需部件**（逐个删掉 app.xml / tableStyles / viewProps / presProps / notesMaster / core.xml 后仍失败）、
**叶子部件本身**（把我们的 theme / presProps / viewProps / tableStyles / app.xml / core.xml，乃至
`ppt/slides/slide1.xml` + 其 rels 换进 PowerPoint 自产的包里，全都开得开）**全部排除**；与
`presentation.xml` 的元素顺序、空白、`sldId`、`notesSz`、占位符 `idx` 也**全部无关**（已单变量排除）
⇒ 元凶就在 `slideMasters` / `slideLayouts` / `notesMasters` 的**组合关系**里，具体是 notesMaster 的主题引用。
**`.pptx` 现在可以直接交付**：不必再先 `office_convert` 成 `.docx` / `.md`，也不必借 PowerPoint 模板另存
（这两条绕行指引已随本条一并删除）。诊断全文与第五轮那两处佐证见 `SKILL.md`「已知边界」。

## 修复注记（2026-09-23 · 第六轮：插图链路三端 / AES-256 读取 / 公式重算 / 稿纸网格）

> 对象：`${DSH_HOME}/plugins/dsh-office\`（DSH 侧）。开发方式沿用红线 6：工作副本
> `<work>/deepseek1\work\office-enhance\plugin\`（`vendor/` 为指向 live 的 junction，
> 既有基线 467 checks（**当轮基线**；后续轮次已增长，准确数量看最后一行的 `N checks`）全绿），跑绿后一次性回写 live。

| # | 需求 | 交付 |
| --- | --- | --- |
| 1 | **插图链路补全（三端）** | ①输入端 `model.js` 的 `markdownToDocument` 把独立成行的 `![alt](path)` 解析成 `image` 块；②`office_edit` 的 docx 新增 `insert_image` / `append_image`（zip 级 `word/media/imageN.*` + image 关系 + `w:drawing`，关系 id 接着既有 `rId` 往后编）；③`pdf.js` 写出端支持 image 块：JPEG `DCTDecode` 原样内嵌、PNG 解成原始采样后 `FlateDecode`（带 alpha 挂 `/SMask`） |
| 2 | **AES-256（R5/R6）读取** | `pdfcrypt.js` 新增 `hash2B`（ISO 32000-2 算法 2.B）、`checkPerms`、`collectV5Keys` 与 `createDecryptor` 的 V5 分支 |
| 3 | **xlsx 公式重算（轻量）** | 新增 `formula.js`（纯 JS 子集求值器）；入口 `office_read recalc=true`（默认关闭），结果进正文与 `stats.recalc` |
| 4 | **可选项** | 4a PDF 合并/拆分只交接口草案、不实现；4b `office_edit` 支持 `.html`（核实后本来就通，补用例与文档口径）；4c `office_create grid` → docx `<w:docGrid>`；4d OMML 默认不做 |

**本批的"病灶"是 DSH 侧的能力缺失（不是既有 bug）**：第五轮之后 DSH 与 WB 分叉，WB 第十二轮已实现
上述四项，DSH 侧没有。移植策略是**只搬目标能力**，不整体覆盖：

- 独立模块直接对齐：`image.js` / `formula.js` 新增；`pdfcrypt.js` / `docx.js` 与 WB 侧**逐字节一致**
  （diff 已核对：全部差异都是本批目标改动，没有 DSH 侧独有内容被覆盖）。
- `pdf.js` **只加图片链路**（WB 同批还把字体管线换成了 `ttf.js`；DSH 侧继续用第五轮的 `pdffont.js`
  + `FontChain`，不把未审计的字体重写带进来）。
- `html.js` **只把 image 分支改成真 `<figure><img src>`**（保留 DSH 第五轮 `documentToHtml` 的其余行为）。
- `index.js` 逐处接线（导入、`editDocx` 插图、`recalc`、`grid`、`.html` 编辑、notice/stats 传递）。

**DSH 侧补的一条硬兜底（WB 侧没有，属本批新增能力）**：`docx` / `odt` / `pptx` / `md` 等写出端
**还没有 image 块分支**，而输入侧现在会产出 image 块 —— 原样交出去会**静默丢图**（`blockToXml` 落进
default 分支、什么都不输出）。所以 `index.js` 新增 `degradeImageBlocks()`：这些目标上把 image 块降级为
**字面文本** `![alt](path)`，并写进 `stats.imageFallback` + `notice`；只有 `pdf` / `html` 保留 image 块。

**实测对照（本机，工作副本；用例前缀 `R12-`）**：

| 用例 | 结果 |
| --- | --- |
| `R12-1` md 图片行 → image 块 | `heading,image,paragraph`（旧版是 `heading,paragraph`，且只剩 `!` + 链接） |
| `R12-1` docx 插图 2 张 | `word/media/image1.png` / `image2.png`，`rId` 唯一递增、`w:drawing` 落在锚点段落之后、`[Content_Types].xml` 补了 png Default、包结构体检通过 |
| `R12-1` 含图 PDF | `/XObject` 资源 + 2 个 ` Do ` 绘制算子 + `/SMask`（透明不丢）；读不到的图片进 `imagesSkipped`（来源+原因） |
| `R12-1` md → docx 含图 | `stats.imageFallback` 1 条 + `notice` 明说"写出端不内嵌"，正文保留字面 `![…](…)` |
| `R12-2` AES-256 | R5/R6 空口令透明解密读出正文、原文件大小/mtime 不变；真口令 / 坏 `/Perms` / 截断三条失败路径都给出明确结论（绝不猜） |
| `R12-3` 公式重算 | `SUM=240`、`ROUND(AVERAGE,2)=1.07`、`COUNTIF(>75)=2`、`IF`、`VLOOKUP` 命中 80 / 未命中 `#N/A`、跨表 `240`、`1/0 → #DIV/0!`、`CONCATENATE → unsupported`（保留原值、绝不猜） |
| `R12-4` 稿纸网格 | `20x25` → `w:linePitch="558"` / `w:charSpace="211"`；非 docx 与非法取值都显式报错 |

**行为边界（本批）**：

| 变更 | 是否改变既有返回 | 说明 |
| --- | --- | --- |
| `markdownToDocument` 的 image 块 | 是（块类型） | 独立成行的 `![…](…)` 从"段落"变成"image 块"；不支持内嵌的写出端按上面的兜底降级并记账 |
| `writePdf` 图片 XObject | 只增 | 没有 image 块时资源字典与产物字节**与旧版逐字一致**（`/XObject` 只在真有图片时出现） |
| `office_edit` 插图操作 | 新能力 | docx 操作集扩大 |
| `office_read recalc` | 只增（默认关闭） | 不设时返回值与旧版逐字一致 |
| `office_create grid` | 只增 | 不设时产物字节不变 |
| `office_edit` 支持 `.html` | 是（原来直接报"暂不支持"） | 文件级文本替换，不改结构、不重排版 |
| AES-256 R5/R6 | 是（原来直接报不支持） | 只影响这类加密文件 |

**红线偏离：无。** 本批没有动质量门判据、没有打破"干净文件逐字不变"（既有 467 条用例全部保持通过 ——
**467 是当轮基线**，准确数量看最后一行的 `N checks`）、
没有碰 `~/.dsh/cordis.patch.yml` 与宿主装配、没有碰 WorkBuddy 目录（只读取其实现作移植源）、
没有新增任何第三方依赖（`zlib` / `crypto` 都是 Node 内置）。

唯一与任务约定**冲突**的是需求 1 的验收项"含图 HTML → `office_convert` → docx 的 `word/media/` 真有图片"：
`docx.js` 的 `writeDocx`/`docBlocksToXml` image 分支被本任务的范围隔离明确排除（另一会话负责），
本批改为**显式降级 + 记账**，并在用例 `R12-1 降级：含图 HTML → docx 时 word/media 仍为空…` 里如实断言。

## 从 SKILL.md 迁出的历史注记（第十七轮 R16 整理）

> 这些段落原先混在技能文档的「已知边界 / 排查」里。它们讲的是**历史上修过什么、为什么这么修**，
> 调用者不需要读；当前行为见 [SKILL.md](SKILL.md)。

- **写 PDF 时不要塞成片私用区码点 —— 这条已是"曾经会死循环、现已修好"的历史注记**：
  `writePdf` 的换行循环（`runPara`）曾把 `fit` 夹在 ≥1，一旦"本行剩余宽度装不下一个字符"就
  **同步死转**，把整个宿主事件循环占死（`Promise.race` 连超时都触发不了）。触发条件**与字符种类无关**：
  私用区因为按 1em 记宽最先撞上（≥47 个就挂），汉字、CJK 扩展 B、`_`.repeat(200)、ASCII 长串同样会挂。
  第三轮已修（改成"估算只作初值 + 真实宽度收敛"，并保证每次循环至少推进 1 个字符）。
  现在成片私用区可以正常写，`test.mjs` 里有 1..200 长度 × 5 种字符的扫描回归。

- **✅ 已修（第七轮定位并修复：notesMaster 独立主题部件）—— 原「⚠ 未解决：`.pptx` 被 PowerPoint 16.0 拒开」
  这条已作废**：病灶是 `Presentations.Open` 报 `0x80070570`「The file or directory is corrupted and
  unreadable」，而 `test.mjs` 只用插件自带解析器回读，所以**套件全绿也盖不住这条**。
  **根因**：`ppt/notesMasters/_rels/notesMaster1.xml.rels` 把备注母版的主题指向与 `slideMaster` **共用**的
  `../theme/theme1.xml`；PowerPoint 要求 notesMaster 有**自己独立的**主题部件（它自己保存时分配 `theme2.xml`）。
  **补丁** = `pptx.js` 新增 `THEME_NOTES` 常量 + `ppt/theme/theme2.xml` 部件 + `[Content_Types].xml` 加 Override +
  `NOTES_MASTER_RELS` 改指 `theme2.xml`。
  **验证**：当日 `node test.mjs` **563 checks ALL PASS**；PowerPoint COM 冒烟**真开成功**（`Presentations.Open`，
  `slides=3`）；**负向控制**（把 notesMaster 改回共用 `theme1`）**必须仍 FAIL**（实测 `0x80070570`）
  —— 即"改对才开、改错必不开"，不是靠放宽条件蒙过去的。
  **历史诊断链（一句话注记，保留修订轨迹）**：zip 层（自产 pptx 用我们的 `makeZip` 重打包仍能开）、
  OPC 不变量（无缺失部件 / 无悬空关系 / 全部 XML 可解析 / 所有 `r:id` 都能解析 / 无重复 Id 与重复 Override）、
  **所有非必需部件**（逐个删掉 app.xml / tableStyles / viewProps / presProps / notesMaster / core.xml 后仍失败）、
  **叶子部件本身**（把我们的 theme / presProps / viewProps / tableStyles / app.xml / core.xml，乃至
  `ppt/slides/slide1.xml` + 其 rels 换进 PowerPoint 自产的包里，全都开得开）**全部排除**；与
  `presentation.xml` 的元素顺序、空白、`sldId`、`notesSz`、占位符 `idx` 也**全部无关**（已单变量排除）。
  ⇒ 元凶就在 `slideMasters` / `slideLayouts` / `notesMasters` 的**组合关系**里，具体是 notesMaster 的主题引用。
  **`.pptx` 现在可以直接交付**：不必再先 `office_convert` 成 `.docx` / `.md`，也不必借 PowerPoint 模板另存
  （这两条绕行指引已随本条一并删除）。
  （第五轮另修好两处——`docProps/app.xml` 的 Override 类型、`aRuns` 把 `<a:solidFill>` 写成 `<a:rPr>` 兄弟的
  DrawingML 违规——两处都正确，但**不足以**让 PowerPoint 接受；真正的解锁点是本轮的 notesMaster 独立主题。）

### 挂载方式的历史（2026-09-26 前）

- **改了插件源码要热生效**：`~/.dsh/cordis.patch.yml` 现在用**裸包名**挂载（`dsh-office` /
  `dsh-office/skill`，由各 profile 的 `node_modules/dsh-office` 链接解析），所以**没有 `?v=N` 缓存爆破**了——
  编辑任何插件源码（含 `html.js` 之类兄弟模块）后，都要**重启 profile**（关掉再起 dsh）才生效。
  （2026-09-23 之前是 `file:///…/index.js?v=N`：`file://` specifier 含 `:`，harness 的 `barePackageName()`
  会直接放弃解析显示元数据，插件列表里只能显示一长串原始 URL。改成裸包名后，列表显示的是包内
  `locale/en.json` + `locale/zh.json` 的 `meta.title` / `meta.description`（中文界面为
  "dsh-office（办公文档）"）。注意 `moduleShortName()` 会剥掉 `dsh-` 前缀，所以友好名必须来自
  locale 字典——字典存在时 title 是多语言对象，UI 直接按语言取，不再走剥前缀那条分支。）

## 挂载方式与热生效（当前事实）

- `~/.dsh/cordis.patch.yml` 用**绝对路径**挂载两行：`id: tool-office` →
  `${DSH_HOME}/plugins/dsh-office/index.js`，`id: skill-office` → `…/skill.js`。
- **没有 `?v=N` 缓存爆破**：改任何插件源码（含 `html.js`）后必须**重启 profile** 才生效。
- 历史上曾用 `file:///…/index.js?v=N`，后改裸包名 `dsh-office`（靠各 profile 的
  `node_modules/dsh-office` 链接解析），2026-09-26 起改回绝对路径 —— 裸包名在 desktop profile 下解析失败。
- 裸包名时代列表会显示包内 `locale/*.json` 的 `meta.title`；绝对路径下 specifier 含 `:`，
  harness 的 `barePackageName()` 放弃解析，列表显示原始路径。

## 未实现能力：草案、结论与重开条件（第十八轮 R18 从 SKILL.md 迁出）

> 这些段落原先在 SKILL.md 的「已知边界」里占了大半屏。它们讲的是**"打算怎么做、为什么现在不做、
> 什么时候才该重开"**，调用者不需要读；SKILL.md 只留一句"不做 + 替代做法 + 本文件入口"。

- **OMML 公式写入：不做**（第七轮 R13 的结论，附量化依据与重开条件）。
  量化：考公/申论目标语料实测（`<corpus>/<corpus>` 下 **486 个 md / 13,048,494 字符**）
  里**矩阵 0 处、根式 `√` 2 处、Unicode 上下标 0 处、积分/求和号 0 处**；`d/d` 形式的 1511 处抽样看，
  绝大多数是编号/年龄分档（`04/05/06`、`16/14/12`）与可线性表达的比例（浓度 `(1-1/3)×(1-1/4)`）。
  现有 Unicode 数学符号（`√ ∑ ∫ ≤ ≥ ≠ ≈ ± × ÷`）+ 文本分数 + `^`/`_` 已够用。成本侧：OMML 是独立命名空间
  （`m:`），docx 写出、`office_read` 反解析、html/pdf 映射**三端都要动**，且红线不允许"看起来像公式但提取不出来"的产物。
  **什么时候该重开**：① 出现必须**可编辑分式/根式/矩阵**的真实排版需求（如数学讲义）；② 且先做**读端**
  （`office_read` 能反解析 OMML）；③ 且任何降级（写成纯文本）都能同时进 `stats` + `notice`。
- **PDF 合并 / 拆分：只交接口草案，未实现**（第七轮 R13；第十八轮 R18 复核后仍维持）。草案形态 = **新增独立工具 `office_pdf`**
  （`op: "merge" | "split"`，不动现有四个工具的签名）。实现阶段必须先重写 xref 表/流（含 `/Prev`、混合 xref、
  增量更新段）、整体重排对象号并跟着改 `/Root` `/Info` `/Names` `/AcroForm` `/Outlines` `/Dest` `/Annots`
  `/StructTreeRoot`、展开或重写 `/ObjStm`、把页面继承属性（`/MediaBox` `/Resources` `/Rotate` `/CropBox`）
  显式落到页对象、移除 `/Linearized`；加密输入默认**拒绝**并给四要素（不同文件密钥无法直接合并，重加密不做）。
  明确不做：表单域合并、数字签名、附件、对象流压缩重排。**不许"拼字节"糊过去**；产出走"临时文件 → 校验 → 原子改名"。
- **`.html` 结构化编辑：只交接口草案，未实现**（第七轮 R13）。`office_edit` 对 `.html`/`.htm` 仍是
  **文件级文本替换**（`replace_text` / `append_text` / `prepend_text`，不改结构、不重排版）。草案新增
  `append_markdown`（追加进 `<body>`）与 `set_meta`（改 `<title>`）两个**只增 op**，走
  `htmlToDocument` → 改模型 → `documentToHtml`；代价是**重建骨架**：原 `<head>`、内联 `style`/`class`、
  未知标签都会丢 —— 按红线必须先报告，且丢失项要逐条进 `stats` + `notice`。
- **插图落点扩展：只交草案**（第七轮 R13）。现状只有文末 / `after=<文本>` / `at=start|end`；
  草案补 `before=<文本>`（低风险，与 `after` 同构）、`in_table=<行,列>`（中）、页眉页脚
  `in_header`/`in_footer`（**高风险**：要新增 `word/headerN.xml` 部件 + `[Content_Types].xml` Override +
  `document.xml.rels` 关系 + `sectPr` 里的 `w:headerReference`，建议单独一批并配 Word 冒烟）。
- **公式重算的写回（`office_edit recalc_formula`）：只交草案，未实现**（第七轮 R13）。
  草案：保留 `<f>` 只换 `<v>`；`t` 按值类型写（文本走共享字符串表，**绝不**写不存在的 sst 索引）；`s` 样式原样保留；
  非 xlsx 目标显式报错；未支持函数保留原值 + 记账。这是唯一"会改文件内容"的 op，放开前要在真 Excel 里冒烟。

## 从 README.md 迁出的实现细节与"当轮实测、未自动化"数字（第十八轮 R18 整理）

> README.md 现在只装快速入门 / 格式能力矩阵 / 常见用法。原先混在它「扫描件 PDF 怎么读」「已知边界」里的
> **实现细节与一次性实测数字**迁到这里 —— 它们不是调用方需要背的契约，也不该在 README 里冒充"当前保证"。

**栅格化 + 识别的两级实现**：

1. **栅格化**：调用 Windows 内置 PDF 渲染器（WinRT `Windows.Data.Pdf`，脚本 `pdf-render.ps1`）
   把页面转成 PNG，中间文件放在系统临时目录 `dsh-office-ocr\`（渲染缓存目录，清理规则见 SKILL）。
2. **识别**：整批 PNG 交给本地引擎 `rapidocr.js`（RapidOCR-json / PP-OCRv4 简中，ONNX Runtime，离线，
   约 0.5 s/页，不产生任何模型调用），再按**质量门**逐页判定：无结果、平均置信 < 0.88、
   或"框多且置信偏软 / 短句占比过高"的复杂版面，才升级给视觉桥（`modlens_read_image`）复核。
   引擎自报 `No text found`（code 101）的页判为空白页，**不**浪费视觉调用。
3. **没过质量门的页自动换渲染倍率重试**：见 SKILL「读取的正确顺序」第 6 条。

**长页面自动分带（仅视觉路径）**：视觉模型的单次输出长度是硬上限（智谱 `glm-4v-flash` 的
`max_tokens` 范围为 [1,1024]，无法调更高）。因此交给视觉桥的文字密集页若一次识别被截断，
插件会自动把该页**横向**切成 2/4/8 段分别识别再拼接（切带在本地纯 JS 完成，段间留 **24 像素重叠**，
避免跨行截断），从而拿到完整文本而不是半页。本地引擎无输出长度限制，整页一次识别。

**当轮实测、未自动化的数字（保留存档，不要再当成当前保证）**：

| 数字 | 出处 | 现在怎么读 |
| --- | --- | --- |
| 盘点 **23 个真题 PDF ≈ 1.1 秒** | 第四轮批量盘点实测 | 单机单次量级参考；套件不断言绝对耗时 |
| 病灶样本**第 17 页**原生不过、`scale=2` 能过 | 第八轮换倍率重试实测 | **具体页号是样本相关的**；套件只断言"备选倍率能过门"的机制 |
| 旧版乱码转换落盘 **56,341 字节 / NUL 5135 个 / C0 13,666 个** | 第二轮 | 已被"出站质量门直接拒绝落盘"取代，数字仅存证 |
| 抽样 **510 个数据格 4.9% 出错**、密集表 **10–15%**、反例 `分析6-p7` | WB 侧第七轮 §7.1 移植 | 定性结论仍然有效（过门 ≠ 可直接引用），数字未自动化 |
| **9000 字中文 PDF 嵌入后 0.39 MB**（未子集化 4.35 MB） | 第五轮字体子集实测 | 套件只断言"样本 <2MB"这条护栏；绝对体积随字体/内容变 |
| Segoe UI Emoji 的 `COLR` 单项 **7.4 MB** | 第五轮子集化剔除实测 | 解释"为什么要剔 COLR"，不是阈值 |
| 渲染倍率逐档耗时 / 置信度 / 升级页数表 | 第十六/十七轮 | 已迁到本文件「第十八轮 R18」§9 |
| 本机原生分辨率 **992×1403（120dpi）** | 同上 | **换机器就会变**；可自动化的是 96dpi×scale 那三档 |

---

> **署名徽章说明（第十八轮 R18 记录 / 第十九轮 R19 复核）**：本文档按交付约定本应在文末挂
> `dsh-badge` 技能规范的 "powered by dsh" 徽章（121×20，不改颜色 / Logo / 项目链接）。
> **R19 复核：`~/.dsh/skills/dsh-badge` 仍不存在**（`Test-Path` → False；`~/.dsh/skills/` 下只有
> `convert-documents-to-markdown` / `photo-to-comic` / `quark-video-reader` / `token-efficient-workflow`
> 四个技能）。拿不到规范图片 ⇒ **继续按约定跳过，不自造徽章**；拿到该技能目录后补挂即可
> （徽章只影响观感，不影响任何功能或契约）。

