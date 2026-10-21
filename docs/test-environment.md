# 测试环境规格（钉死版，2026-10-21 实测）

全套测试（114 项，含 §7 的 BUG-B1 回归 6 项）要可复现，必须同时钉死四件事：**Python 依赖、LibreOffice、
官方模板的转换链、字体**。一键执行：`bash scripts/setup-test-env.sh`（幂等）。
本文件同时记录 2026-10-21 的 30 项真模板失败逐项归因结论——归因表全文在
修复工作笔记中，本文件只留结论与环境侧处置。

## 1. Python 依赖（实测版本）

| 项 | 实测版本 | 说明 |
|---|---|---|
| Python | 3.12.3 | |
| python-docx | 1.2.0 | requirements.txt 只写 `>=1.1` 未锁版；样式名内部式解析（BabelFish）1.1/1.2 行为一致，不是本次失败的变量，但复现时请用 1.2.0 |
| pymupdf | 2.0 | 逐页文本提取顺序影响个别渲染断言，见 §5 |
| pytest | 9.1.1 | |

## 2. LibreOffice

- 系统版 **24.2.7.2**（`/usr/bin/soffice`）为基准。
- 已对照实测：官方历史版 **7.6.7.2**（TDF archive 解压直跑）对固定稿的渲染
分页与 24.2.7.2 **逐页一致**（18 页、封面溢出位置、页码序列、文本块顺序全同）
——渲染差异不是 LO 版本造成的，不必降版。

## 3. 官方模板转换链（本次 30 项失败中 25 项的根源）

规定动作（与 README 一致＋新增第 3 步）：

```bash
# 1) 下载官方.doc（北大研究生院 2024 版硕士模板），基准件 sha256:
# bcdf2bbd7cac76a1cb281397580004eb1f23f0f8a56d5c1dc782fd3d2f5c9f36
# 2) LibreOffice 转换
soffice --headless --convert-to docx --outdir local-assets/pku local-assets/pku/master-2024.doc
# 3) 样式名归一化
python3 scripts/normalize-template-styles.py local-assets/pku/master-2024.docx
```

**为什么必须有第 3 步（实测定案）**：LibreOffice 24.2 转换产物 `word/styles.xml`
里内置样式名是 UI 形式（`Heading 1`、`TOC 1`、`Caption`），MS Word 产物是内部
形式（`heading 1`、`toc 1`、`caption`）。python-docx 按内部名解析
（`BabelFish.ui2internal`），于是 LO 转换稿上：

- `check` 按名找不到 `toc 1/2/3` → 误报 3 条 `toc_style_missing`；
- `fix` 的 `_ensure_style` 捕获 KeyError 后**新建同名幽灵样式**
（产物里同时出现原 `TOC 1`(id `TOC1`) 与新 `toc 1`(id `toc1`)），
全部样式级修复写进无人使用的幽灵样式 → 修复后静默残留 **86 条**可修复问题
（size 16、bold 16、font_east_asia 12、align/line_spacing/space_before/
space_after 各 10、font_latin 2，集中在用 Heading 1/2/3 的约 10 个段落）。

单变量对照实验：只改样式名（styleId 不动）后，`check` 总数回到文档记载的
199 条且 `toc_style_missing` 消失，`fix` 后残留可修复问题 **86 → 0**，
残留恰为测试预期的内容层建议类（ref_mixed_system / ref_type_tag / empty_filler）。

产品侧的根治（按 styleId 解析、禁止静默建幽灵样式）已立 bug 报告，属复核轮
裁决事项；**环境侧先用归一化钉死转换链**，不改产品代码、不改测试断言。

## 4. 字体

模板写死的 宋体/黑体/仿宋 是专有字体，Linux 无官方免费分发；不许从非官方
站点抓字体包。钉死方案 = Ubuntu 官方源开源字体＋fontconfig 替代映射
（`config/fontconfig/99-docx-format-guard.conf`，setup 脚本装到
`~/.config/fontconfig/conf.d/`）：

| 模板字体名 | 替代 | 来源包 |
|---|---|---|
| 宋体 / SimSun / 仿宋 / FangSong / 楷体 | Noto Serif CJK SC | fonts-noto-cjk |
| 黑体 / SimHei | Noto Sans CJK SC | fonts-noto-cjk |
| Times New Roman | Liberation Serif（度量等价） | fonts-liberation |
| Arial | Liberation Sans（度量等价） | fonts-liberation |
| Calibri | Carlito（度量等价） | fonts-crosextra-carlito |
| Cambria | Caladea（度量等价） | fonts-crosextra-caladea |

对照组（已装、已实测，不作默认）：fonts-arphic-uming（AR PL UMing CN）、
fonts-wqy-zenhei（文泉驿正黑）。

## 5. 已知红榜：钉死环境后仍失败的 5 项（归因完毕，非环境问题）

| 测试 | 归因 | 证据要点 |
|---|---|---|
| test_pku_wps::test_cover_single_page | 缺专有字体 | 固定稿封面溢出 2 行（复选框行＋日期行落到第 2 页）。四种开源组合实测：Noto 系行高偏大→溢出；UMing＋ZenHei 组合封面可合页但前置部分被压缩（罗马页 I–V 变 I–III）且正文多出空白页。LO 7.6/24.2 结果一致。真分页介于两组开源度量之间，无开源字体可复现 |
| test_pku_wps::test_page_summary_no_hidden_blank | 缺专有字体 | 同上：溢出页混入 nearly-empty 集合、页码序列随度量漂移 |
| test_pku_wps::test_nearly_empty_pages_are_template_placeholders | 缺专有字体 | 残留页即封面溢出页（第 2 页、2 行） |
| test_pku_wps::test_parity_note_explains_wps_virtual_page | 缺专有字体 | 逐页摘要中目录页页眉/罗马页码位置随度量漂移：Noto 栈下无 header=='目录' 的页（StopIteration）；UMing 栈下该页存在（第 5 页/III）但期望的“止于 V”不成立 |
| test_pku_features::test_page_sections_headers | 渲染链（PDF 文本块顺序） | 断言假设摘要页文本第 2 行是页脚罗马数字 `I`；实测 LO（7.6 与 24.2 一致）写入顺序为 页眉→正文→页脚，`I` 在第 8 行。文档与 PDF 页脚内容本身正确，是测试辅助函数 `lab()` 对提取顺序的假设不成立 |

**WPS 说明**：30 项失败中“缺 WPS”占 **0 项**。test_pku_wps 按其文件头注释
设计为用 LibreOffice 渲染（WPS for Linux 无可用的无界面转 PDF 命令行，
作者环境亦然），失败与 WPS 缺席无关。

## 6. 基准数字（本机实测）

| 阶段 | 全套（108） | 三个 PKU 文件（54) |
|---|---|---|
| 原始环境（LO 直转、无归一化、无映射） | 78 过 / 30 败 | 24 过 / 30 败 |
| 钉死环境（归一化＋映射，§3+§4） | **103 过 / 5 败** | 49 过 / 5 败 |

由红转绿的 25 项全部来自转换链归一化（含 test_pku_wps::test_heading_keep_next：
其失败是幽灵样式污染了测试对 `styles["Heading 2"]` 的修改落点，同根同源）。
字体映射不改变通过数——它钉死的是渲染确定性（无映射时 fc 默认回退甚至会选到
Noto Sans CJK **HK** 字形），并消除了“换个回退字体结果又不一样”的不确定性。

## 7. 复核轮（C 组）补记：BUG-B1 已在产品侧根治，归一化降为环境卫生

2026-10-21 复核轮实测更新（§1–§6 的归因与数字仍成立，本节是其后进展）：

- **BUG-B1 已修**：`guard/core.py` 新增名式不敏感的样式解析（`find_style`：
  原生名 → 样式名归一 → styleId 归一，永不新建）与 `require_style`
  （解析不到显式抛 `StyleResolutionError`）；`_ensure_style` 先解析、
  确不存在才新建，且新建必须记入 fix 的 changes 清单。`features.py` 的
  check_toc/fix_toc/rebuild_toc_result、脚注与表样式、标题 numPr 清理
  全部改走同一解析。回归测试 `tests/test_style_resolution.py`（6 项，
  对修复前代码全红、修复后全绿）。
- **未归一化原始转换稿上的产品验收（实测）**：check 总数 199、
  `toc_style_missing` 0（修复前误报 3）；fix 后可修复残留 **0**（修复前 86）；
  产物无重复 styleId（修复前 Heading1/2/3 各有一份重复元素、另有
  toc 1/2/3 与 footnote text 幽灵样式）。
- **最新全套数字**：钉死环境（归一化模板）**109 过 / 5 败**（114 项，
  5 败即 §5 已知红榜）；未归一化原始模板 **107 过 / 7 败**——多出的
  2 败在测试代码自身：test_pku_features::test_toc 的注入函数与
  test_pku_wps::test_heading_keep_next 直接用 python-docx 原生
  `styles["Heading 1"]`/`styles["toc 1"]` 改文档，python-docx 原生查找
  在 UI 名文档上同样落空（与 §5 的 `lab()` 同属测试侧脆弱点，待测试轮
  改走 `core.find_style` 后可消）。产品代码路径在原始模板上已无失败。
- 因此 `scripts/normalize-template-styles.py` 不再是正确性前提，
  保留为环境卫生（让测试代码与外部工具按名查找也能工作）。
