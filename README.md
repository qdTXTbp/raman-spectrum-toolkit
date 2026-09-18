<div align="center">

<img src="assets/icon.png" width="128" alt="Raman Spectrum Toolkit">

# 拉曼光谱工具 · Raman Spectrum Toolkit

**JASCO `.jws` 光谱转换 · 拉曼峰分析 · 矿物鉴定**
**Convert · Analyze · Identify — JASCO `.jws` / CSV / SPC / JCAMP-DX spectra**

<b>中文</b> ｜ <a href="#english">English</a>

</div>

---

<a id="中文"></a>

## 这是什么

一个 Windows 桌面工具，把光谱仪导出的原始文件（JASCO `.jws`、`.csv`、`.spc`、`.jdx`、`.txt`、`.xlsx`）转成能直接看的图和能直接用的数据，
并内置拉曼光谱的常用处理、峰分析、**未知矿物鉴定**与 **RRUFF / ROD 参考谱库**对接。

界面、命令行、分析报告与说明书全部**中英双语**。

## 一眼看懂的功能

| 类别 | 能做什么 |
| --- | --- |
| **格式转换** | `.jws`/CSV/SPC/JCAMP-DX/TXT → CSV、**带图的 Excel**、PNG 曲线图、峰列表、JCAMP-DX |
| **峰分析** | 自动找峰并标注峰位（**峰位带波长虚线**）、**标错的峰连同自动峰一起右键删除、可一键恢复**、手动补标、高斯/洛伦兹/伪 Voigt 峰拟合、峰位检索；峰位数值可一键隐藏 |
| **未知谱鉴定** | 不知道样品是什么？拿它的峰去整个参考库比对，按可信度给出候选矿物 + 峰位对照；**批量鉴定**可一次把一个文件夹的谱逐条给出最佳候选 |
| **参考谱库** | 在线检索 ROD、按矿物批量抓取 RRUFF 数据包（拉曼 / 红外 / XRD / 化学成分），导出为本地库；**下载支持断点续传 + 32 路并发 + 动态领活 + 自动重试，可填代理，后台进行、随时可取消，大包按实测速率标出预计耗时** |
| **配对比较** | 实测谱 ↔ 标准谱手动/自动配对，输出峰位匹配 F1、相关系数、谱角与对照报告图；**批量配对**可拿一组标准谱一次筛一个文件夹，**对不上的排在最前面**（综合分、F1、相关系数等全列出来，判读只是提示，结论自己下） |
| **预处理** | 尖峰（宇宙射线）去除、基线校正、平滑、导数、归一化、拉曼位移校准 |
| **统计分析** | 层次聚类 + PCA、二维成像热图、平均/相减、谱段替换、交互式 A−k·B 找平 |
| **多谱对照** | **瀑布图**（纵向错开，看有哪些峰）+ **多数据图叠加**（按 stacked spectra 排布：各条**上下错开、谱线分开**，每条一色带图例；**峰位跨谱合并，一个峰只画一条虚线、只标一个平均波数**；**先预览、左键加峰 / 右键删峰，确认后再导出**；横坐标取各条**交集**，可导出 PNG） |
| **图例摆放** | **几条谱线就列几条**（不再只列前 12 条），放不下自动分列；图例画在**绘图区外面**（默认右侧留白），谱线再密也不遮挡；位置可选右侧留白 / 上方 / 下方 / 图内自由位置 / 不显示，也能在预览窗口里**按住图例直接拖**，位置记住、导出沿用 |
| **批量与报告** | 整目录批处理 + 汇总表、自包含 HTML 分析报告（图片内嵌，可打印成 PDF） |

<table>
<tr>
<td width="50%"><img src="docs/example_spectrum.png" alt="自动标峰的拉曼谱"></td>
<td width="50%"><img src="docs/example_pairing.png" alt="配对对比报告"></td>
</tr>
<tr>
<td align="center"><sub>自动识别并标注峰位，每个峰用虚线引到横坐标轴（锆石参考谱，785 nm）</sub></td>
<td align="center"><sub>未知谱 ↔ 库中标准谱配对报告（峰位匹配 F1 排序）</sub></td>
</tr>
</table>

<div align="center">

<img src="docs/example_overlay.png" alt="多数据图叠加">

<sub><b>多数据图叠加</b> — 4 条光谱按 stacked spectra 上下错开、谱线分开，每条一色；峰位跨谱合并后<b>一个峰只画一条虚线、只标一个平均波数</b>。打开先出预览窗口，<b>左键加峰、右键删峰</b>，确认后再导出 PNG；<b>图例画在绘图区右侧的留白里</b>，谱线再密也不遮挡，还能直接拖到别处</sub>

</div>

## 快速开始

### 方式一：绿色版（推荐，免安装）

1. 到 **[Releases](../../releases)** 下载最新的 `RamanSpectrumToolkit-v2.x.zip`
2. 解压到任意目录（含中文路径也可以）
3. 双击 `RamanSpectrumToolkit.exe`

不需要 Python，不写注册表，不联网也能用（联网只用于下载参考谱）。
所有工具自己下载/生成的数据都写在解压目录的 `工具数据/` 里，不会污染样品文件夹。

### 方式二：从源码运行

```bash
pip install pillow openpyxl        # PNG 出图与 Excel 导出需要；只用 CSV 可跳过
python jws2csv.py                  # 打开图形界面
```

也可以双击 `启动拉曼光谱工具.bat`（自动寻找本机 Python）。

## 命令行速查

```bash
python jws2csv.py a.jws 某文件夹              # 批量转换
python jws2csv.py --all a.jws                # CSV + Excel + PNG + 峰列表 + 峰拟合
python jws2csv.py --identify unknown.csv     # 未知光谱全库鉴定
python jws2csv.py --pair unknown.csv         # 与本地参考谱库配对
python jws2csv.py --pair-batch 待测文件夹 --pair-ref 锆石标准谱   # 批量配对，对不上的排最前
python jws2csv.py --identify-batch 待测文件夹 --identify-top 5    # 批量鉴定，逐条给最佳候选
python jws2csv.py --rruff-fetch Zircon       # 一键下载 + 检索 + 导出锆石参考谱
python jws2csv.py --mineral-info Zircon      # 矿物信息卡（特征峰归属 + RRUFF 样品记录）
python jws2csv.py --cluster 文件夹           # 聚类分析 + 主成分
python jws2csv.py --map 5,5 --map-metric main_peak   # 二维成像热图
python jws2csv.py --waterfall 文件夹         # 瀑布图（纵向错开）
python jws2csv.py --overlay 文件夹           # 多数据图叠加（每条一色，批处理无预览）
python jws2csv.py --report 文件夹            # 自包含 HTML 分析报告
python jws2csv.py --lang en|zh               # 切换界面/输出语言
python jws2csv.py --manual                   # 打印完整说明书
```

完整参数见程序内【帮助 → 使用说明】或仓库里的 `使用说明.txt` / `User_Guide.txt`。

## 未知谱鉴定是怎么判的

先用参考库的**特征峰索引**做峰位粗筛（毫秒级），再对前列候选读原始谱精算：

* **综合分** = 0.5 × 峰位匹配 F1 + 0.5 × 强峰命中率
* **可信候选**要求三条同时成立：综合分 ≥ 60、未知谱最强 3 个峰至少命中 2 个、且领先第二名 ≥ 15 分
* 同时给出**偶然概率**（二项分布），分数接近时明确提醒“可能是库里没有对应矿物”
* 未命名样品单独提示，不会用没有矿物名的记录冒充结论

## 批量配对 / 批量鉴定

单条谱的鉴定与配对都只能一次处理一条。**批量**版本在外面套一层循环，
一次吃一整个文件夹，汇总成一条谱一行的表：

* **批量配对**：拿一组参考谱（比如一批锆石标准谱）去筛一个文件夹的实测谱。
  汇总表**按综合分升序排**——对不上的排在最前面，一眼就能看到可疑的。
* **批量鉴定**：一堆不认识的谱，逐条在全库里找最像的矿物，给出最佳候选与结论文本。
* **汇总表列的是辅助数据**：综合分（与上面「未知谱鉴定」同一口径，所以两处数字对得上）、
  F1、强峰命中、命中 / 实测 / 参考峰数、平均偏差、相关系数、谱角、偶然概率、领先第二名。
* **「参考判读」一列只是提示**（匹配良好 / 部分匹配 / 匹配很差），
  **不下最终结论**——到底是不是同一种物相，请自己核对峰位与谱型。
  表里某一行可疑，可以再单独对它跑一次配对，看峰位对照表差在哪几个峰。
* 输出：汇总表 `CSV` + 汇总报告 `HTML`（浏览器里可直接打印为 PDF）。
* 注意：只识别到 **1 个峰**的谱，F1 一律记 0（命中 < 2 不算识别），
  综合分因此封顶 50、一直停在「部分匹配」档。这是故意的——单个峰定不了案。

## 下载数据库为什么慢，以及现在怎么处理

RRUFF 官网在国外，**主要瓶颈是这条跨境链路**（实测 ping 丢包 25%、RTT 236 ms），
不是工具的问题。在你的网络下实测：

| 同时开的连接数 | 实测吞吐 | 227 MB 的包要等 |
|---|---|---|
| 1 路 | ~0.03 MB/s | 约 2 小时 |
| 8 路 | ~0.32 MB/s | 约 12 分钟 |
| **32 路（默认）** | **~0.89 MB/s** | **约 4 分钟** |
| 64 路 | ~1.28 MB/s | 约 3 分钟 |
| 96 路 | ~1.45 MB/s | 约 2.6 分钟 |

**吞吐几乎正比于同时在跑的连接数**（服务器全程没有拒绝），所以并发数是最好用的那个旋钮。
默认给 32 路，可在【下载】对话框的“并发连接数”里调（1~128），
命令行 `--dl-conns 96` 也行。
（上表取自同一次测量；这条链路随时段波动很大，绝对值能差 2~3 倍，
但“连接数越多越快”这个趋势是稳定的。另外为了不把开销都花在握手上，
每个连接至少要分到 256 KB 的活，所以很小的包不会硬开几十路。）

既然带宽改不了，就把**能拿到的全拿到**：

* **后台下载**：进度条 + 已下/总数 + 实时速率 + 剩余时间。以前下载期间窗口是**假死**的
  （旧代码在循环里只调 `update_idletasks()`，不处理鼠标和关闭事件），
  十几分钟点不动、关不掉，只能强杀进程——这就是你说的"崩溃"。
* **断点续传**：服务器支持 `Accept-Ranges`，取消 / 掉线 / 强杀**都不会白费**，
  下次接着下。以前一断就从 0 重来。
* **动态领活（work-stealing）**：不是"把文件平均分给 N 条连接"，
  而是切成小段放进队列、**谁下完谁再领一段**。
  实测同一个 33 MB 的包，32 等分里**最快那段 4.3 s、最慢那段 79 s**——
  静态均分等于花 79 s 等一条卡住的连接（95% 的时间在干等）。
  真实链路同一时段交错 A/B（12 MB 的包、都是 32 路）：
  **静态均分平均 39.1 s → 动态领活平均 23.2 s，快 1.69 倍**。
* **代理**：丢包严重的链路上，走代理 / VPN 往往比堆并发更快。
  在【下载】对话框填代理地址（留空则沿用 Windows 系统代理），或命令行 `--proxy`。
* **自动重试**：断线按 1.5×n 秒退避重试，并从断点继续；服务器返回 429/503
  会识别出来并提示"把并发调小些"。
* **完整性校验**：长度对不上**不算下载成功**。以前 `Content-Length` 只用来显示进度，
  截断的 zip 会被当成"下载成功"并标记"已下载"，直到建索引时才报错——
  看着就像又崩了一次。
* **预计耗时**：数据包列表里按**你上次实测到的速率**估算（越用越准），
  下 ≥50 MB 的大包前会弹窗告知要等多久。

如果实在慢，先下 `fair_oriented`（271 KB）练手，或者用浏览器在能直连的网络上下好 zip，
再用【导入本地 zip…】。

## 中文 / 英文

界面、日志、命令行输出、分析报告、图注与说明书都可切换：

* 菜单【设置 → 语言】
* 命令行 `--lang en|zh`
* 首次启动跟随 Windows 显示语言，之后记住你的选择

磁盘上的数据目录名（`工具数据 / 参考谱库 / RRUFF数据包 / 分析结果`）保持中文，保证两种语言下数据互通。

## 数据来源与致谢

* **RRUFF**（[rruff.info](https://rruff.info)）—— 参考谱与数据包来自 RRUFF 项目，请遵守其使用条款；
  本工具只是下载与检索客户端，**与 RRUFF 项目无隶属关系**。
* **ROD, Raman Open Database**（[rod.ens-lyon.fr](https://rod.ens-lyon.fr)）—— 在线拉曼参考谱检索。
* `.jws` 二进制结构参考开源项目 `jasco_jws_reader` / `jasco-jws-converter` 的 DataInfo 说明；
  本工具的数值解析已与参考实现逐点比对一致。

## 许可

[MIT](LICENSE)

---

<a id="english"></a>

## English

**Raman Spectrum Toolkit** is a Windows desktop tool that turns raw spectrometer exports
(JASCO `.jws`, `.csv`, `.spc`, JCAMP-DX `.jdx`, `.txt`, `.xlsx`) into charts and usable data,
and bundles the everyday Raman workflow: peak detection and fitting, mineral peak assignment,
**unknown-spectrum identification** against RRUFF / ROD reference libraries, pairing,
clustering, 2D imaging, batch conversion and HTML reports.

The GUI, CLI, reports and manuals are **fully bilingual (Chinese / English)**.

### What it does

| Area | Capability |
| --- | --- |
| **Conversion** | `.jws` / CSV / SPC / JCAMP-DX / TXT → CSV, **Excel with embedded chart**, PNG plot, peak table, JCAMP-DX |
| **Peak analysis** | automatic peak detection with position labels (**each peak gets a dashed line down to the x axis**), **right-click to delete a wrong peak — automatic ones included — and restore them all with one click**, manual annotation, Gaussian / Lorentzian / pseudo-Voigt fitting, peak-position search; peak values can be hidden |
| **Unknown spectra** | identify a spectrum whose mineral you do not know by matching its peaks against a whole reference library, with confidence ranking and a peak-by-peak comparison; **batch identification** runs a whole folder and gives the best candidate per spectrum |
| **Reference libraries** | search ROD online, bulk-fetch RRUFF packages (Raman / IR / XRD / chemistry), export them into a local library; **downloads resume after a drop, use 32 parallel connections with work-stealing scheduling and automatic retries, accept a proxy, run in the background with a working Cancel button, and show an estimated time based on your own measured speed** |
| **Pairing** | measured ↔ reference pairing (manual or automatic) with peak-match F1, correlation, spectral angle, and a comparison report figure; **batch pairing** screens a whole folder against a reference set and **puts the mismatches first** (score, F1, correlation and more — the reading is advisory, you decide) |
| **Preprocessing** | spike (cosmic ray) removal, baseline correction, smoothing, derivative, normalization, Raman shift calibration |
| **Statistics** | hierarchical clustering + PCA, 2D imaging heat map, average / subtract, range replacement, interactive A−k·B flattening |
| **Multi-spectrum comparison** | **waterfall** (offset stacks, to see *which* peaks are there) + **multi-dataset overlay** (stacked-spectra layout: curves **offset and separated**, one colour each with a legend; **peaks merged across datasets — one dashed line and one averaged value per peak**; **preview first, left-click to add / right-click to remove peaks, export only after you confirm**; x axis is the **intersection** of all ranges) |
| **Legends** | **one entry per curve** (no longer capped at the first 12), wrapping into extra columns when needed; the legend is drawn **outside the plot area** (right-hand gutter by default) so dense curves never cover it; pick right gutter / above / below / free position inside / hidden, or **drag the legend** in the preview window — the position is remembered and reused for exports |
| **Batch & reports** | whole-folder batch conversion with a summary table, self-contained HTML report (images embedded, printable to PDF) |

### Quick start

**Portable build (no Python needed)**

1. Download the latest `RamanSpectrumToolkit-v2.x.zip` from **[Releases](../../releases)**
2. Unpack anywhere and run `RamanSpectrumToolkit.exe`

Everything the tool downloads or produces stays inside `工具数据/` next to the executable,
so your sample folders are never touched.

**From source**

```bash
pip install pillow openpyxl
python jws2csv.py
```

### Command line

```bash
python jws2csv.py --identify unknown.csv    # identify an unknown spectrum
python jws2csv.py --pair unknown.csv        # pair against your local library
python jws2csv.py --pair-batch FOLDER --pair-ref ZIRCON_REFS   # batch pairing, mismatches first
python jws2csv.py --identify-batch FOLDER --identify-top 5     # batch identification, best candidate each
python jws2csv.py --rruff-fetch Zircon      # download + index + export Zircon references
python jws2csv.py --waterfall folder        # waterfall chart (offset stacks)
python jws2csv.py --overlay folder          # multi-dataset overlay, one colour each (batch, no preview)
python jws2csv.py --report folder           # self-contained HTML analysis report
python jws2csv.py --lang en|zh              # switch UI / output language
python jws2csv.py --manual                  # print the full user guide
```

### How identification works

Reference peaks are pre-indexed for a millisecond-level pre-filter; the top candidates are then
re-scored against their original spectra. The **combined score** is
`0.5 × peak-match F1 + 0.5 × strong-peak hit rate`. A result is marked **reliable** only when
the score ≥ 60, at least 2 of the 3 strongest peaks are matched, and it leads the runner-up by ≥ 15
points; a binomial **chance probability** is reported and close scores trigger an explicit warning
that the mineral may simply not be in the library. Unnamed RRUFF records are never presented as an
identification.

### Batch pairing / batch identification

Identification and pairing each handle ONE spectrum at a time. The **batch** versions wrap a loop
around them and take a whole folder, producing a one-row-per-spectrum table:

* **Batch pairing** — screen a folder of measured spectra against a reference set (say a batch of
  zircon references). The table is **sorted by ascending score, so the mismatches come first**.
* **Batch identification** — for a pile of unrecognised spectra, each one is searched against the
  whole library and gets a best candidate plus the conclusion text.
* **The table holds advisory metrics**: the combined score (the same scoring as above, so the two
  places agree), F1, strong-peak hits, matched / measured / reference peak counts, mean deviation,
  correlation, spectral angle, chance probability and the lead over the runner-up.
* **The "Advisory reading" column is a hint only** (good match / partial match / poor match).
  It does **not** decide for you — check the peaks and the profile yourself. If a row looks off,
  run the single-spectrum pairing on it to see which peaks disagree.
* Output: a summary `CSV` plus a summary report `HTML` (printable to PDF from a browser).
* Note: a spectrum with only **one** detected peak always gets F1 = 0 (fewer than 2 matches does
  not count), so its score caps at 50 and it stays in the "partial match" band. That is deliberate —
  one peak cannot settle anything.

### Why package downloads are slow, and what was done about it

rruff.net is hosted outside China, so **the bottleneck is this cross-border link** (measured: 25%
packet loss, 236 ms RTT) rather than a bug in the tool. Measured on this network:

| Simultaneous connections | Measured throughput | Time for the 227 MB package |
|---|---|---|
| 1 | ~0.03 MB/s | ~2 hours |
| 8 | ~0.32 MB/s | ~12 minutes |
| **32 (default)** | **~0.89 MB/s** | **~4 minutes** |
| 64 | ~1.28 MB/s | ~3 minutes |
| 96 | ~1.45 MB/s | ~2.6 minutes |

**Throughput is roughly proportional to the number of simultaneous connections** (the server never
refused), so concurrency is the lever that matters. The default is 32; change it in the download
dialog's "connections" field (1–128) or with `--dl-conns 96` on the command line.
(The table is from a single measurement session; this link varies a lot by time of day, with
absolute numbers swinging 2–3x, but the "more connections is faster" trend is stable. Each
connection is also guaranteed at least 256 KB of work so that small packages do not spend
everything on handshakes.)

Bandwidth cannot be fixed, so everything obtainable from the code side is now taken:

* **Background download** with a progress bar, downloaded/total, live speed and time left.
  The window used to be **frozen** during a download (the old loop only called
  `update_idletasks()`, which does not process mouse or close events), so for ten-odd minutes it
  could not be clicked or closed and the only option was to kill the process — that is the "crash".
* **Resume**: the server supports `Accept-Ranges`, so a cancel, a dropped connection or a
  force-kill **no longer wastes anything**; the next run continues where it stopped. Previously it
  restarted from zero.
* **Work-stealing scheduling**: instead of splitting the file into N equal shards (one per
  connection), the file is cut into work units in a queue and **a connection that finishes grabs the
  next unit**. Measured on the same 33 MB package, the fastest of 32 equal shards needed 4.3 s while
  the slowest needed 79 s — equal shards mean 79 s of waiting for one stalled connection (95% of the
  wall clock spent idle). Interleaved A/B on the real link (12 MB package, 32 connections both
  ways): **39.1 s average for equal shards vs 23.2 s with work stealing — 1.69x faster**.
* **Proxy**: on a lossy link a proxy or VPN is often faster than piling on connections. Enter one in
  the download dialog (leave blank to reuse the Windows system proxy) or use `--proxy`.
* **Automatic retries** with 1.5 x n second backoff, resuming inside the part; a 429/503 from the
  server is detected and reported with a "lower the concurrency" hint.
* **Integrity check**: a byte-count mismatch is **not** treated as success. Previously
  `Content-Length` was only used for the progress display, so a truncated zip was recorded as a
  successful download and marked "downloaded" — the failure surfaced only later when the index was
  built, which looked like another crash.
* **Estimated time** in the package list, based on **your own last measured speed** (it improves over
  time); packages of 50 MB or more ask for confirmation first.

If it really is too slow, start with `fair_oriented` (271 KB), or download the zip with a browser on
a well-connected network and use "Import local zip...".

### Credits

Reference spectra and data packages come from the **RRUFF** project ([rruff.info](https://rruff.info))
and the **Raman Open Database** ([rod.ens-lyon.fr](https://rod.ens-lyon.fr)).
This tool is an independent client and is **not affiliated with those projects**.

### License

[MIT](LICENSE)
