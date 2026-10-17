<p align="right"><strong>简体中文</strong> | <a href="https://github.com/labulaka-ai/docx-format-guard/blob/main/README.en.md">English</a></p>

# docx-format-guard · 学位论文 Word 格式守护

![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg) ![Rules](https://img.shields.io/badge/rules-CC%20BY%204.0-lightgrey.svg) ![Agent Skill](https://img.shields.io/badge/Agent%20Skill-SKILL.md-green) ![Docs](https://img.shields.io/badge/docs-中文%20%7C%20English-orange)

> 作者：**Labulaka**（GitHub [@labulaka-ai](https://github.com/labulaka-ai)）· 公众号「**时空AI深水区**」

[为什么做这个](#为什么我做这个) · [功能](#功能) · [北大规则包](#北大规则包) · [安装与使用](#安装与使用) · [作为 agent skill](#作为-agent-skill-使用) · [可信度](#可信度测试与实测) · [已知限制](#已知限制)

我是时空大数据方向的博士在读。这个工具解决的是写学位论文时最磨人的一件事：**内容改完了，Word 格式乱了**。它按学校的规则包检查 .docx，把漂移的格式修回去，同时**不碰导师留下的修订和批注**；也能列出、回复这些批注，并用 LibreOffice 排版逐页核对版面。

它不是一个“一键美化”的模板，而是一套带出处的规则：每条规则都写明来自学校指南的哪一条；修不了或不该由工具改的（比如参考文献著录制混用），只给建议并说明原因，不悄悄改你的内容。

---

## 为什么我做这个

论文从初稿到定稿，要经过导师改稿、合并稿、从别处粘贴、在 Word / WPS / LibreOffice 之间来回打开。格式问题非常稳定地出现在这几处：

- **改稿后格式漂移**：粘贴进来的段落带着 14pt 楷体、Arial；冒出一个 “Normal (Web)” 之类的外来样式；行距、首行缩进、页边距被悄悄改掉；
- **改格式时毁掉修订和批注**：常见的“全选→清除格式”或重套模板，会把导师的修订痕迹、批注线程一起冲掉；
- **前置部分和页码最难弄**：封面、摘要、目录、原创性声明各自要求不同的页眉页码，分节一错，奇数页页眉、罗马/阿拉伯页码、目录页码全跟着错；
- **不同软件版面不一样**：缺宋体/黑体/仿宋的环境里，封面挤成两页、签名栏被单独挤到下一页。

这个工具就是针对这四类问题做的：先检查、再修复，修完逐项复查，并用排版结果核对版面。

## 功能

| 命令 | 作用 |
|---|---|
| `check` | 按规则包检查格式，每条问题给出期望值、实际值、原因和指南条款；`--json` 输出机器可读结果 |
| `fix` | 修复可修复的问题，输出到新文件（不覆盖原稿）；保留修订（w:ins / w:del / 格式修订）与批注；修完自动复查，列出剩余问题 |
| `diff` | 比较改稿前后两份 docx，按文字对齐段落，只报告**格式**变化（样式、中西文字体、字号、加粗、行距、段前段后、缩进、对齐） |
| `review` | 列出全部修订与批注（作者、时间、所在段落、批注线程关系） |
| `reply` | 回复某条批注（在 Word 中显示为同一线程），或在某处修订上加一条批注 |
| `pages` | 用 LibreOffice 排版，逐页给出页眉、页脚（页码）、正文首尾几个字，标出空白页 / 几乎空的页，以及相邻页码同为奇数时双面打印的“虚拟空白页” |

问题分两类：**可修复**（`fix` 自动处理）和**只给建议**（内容层面的问题，例如参考文献著录制混用、关键词个数、模板占位空段，工具不代改，只说明原因和建议）。

## 北大规则包

`profiles/pku-master.yaml`：北京大学硕士学位论文。依据：

- 《北京大学研究生学位论文写作指南》：<https://grs.pku.edu.cn/docs/2019-09/20190911111739293567.pdf>（材料下载页标注 2018-03-01）
- 《硕士研究生学位论文格式模板(2024)》：<https://grs.pku.edu.cn/xwgz11/xwsy11/ssxw111/clxz08/50118yjsy346393.htm>（2024-02-29 发布）

覆盖：页面设置（A4、页边距、装订线、对称页边距）、奇偶页页眉（奇数页为章名，用 STYLEREF 域自动取）、页眉横线、分节页码（封面和声明无页码，摘要至目录罗马数字，正文阿拉伯数字）、封面、版权声明和原创性声明（固定文本）、中英文摘要与关键词、目录（三级、前置部分不进目录、页码回写）、标题自动编号（第一章 / 1.1 / 1.1.1）、正文与标题字体字号行距、脚注（①②按页编排）、表格与图表题注、公式编号、GB/T 7714 参考文献。每条规则带 `clause` 字段，写明指南条款。

官方模板**不随本仓库分发**。需要用模板的功能（固定文本比对、模板自检测试）请自行下载到 `local-assets/`（已 .gitignore）。

`profiles/example-master.yaml` 是一个通用数值的示例规则包，不对应任何学校，用来说明规则包怎么写。

## 安装与使用

```bash
git clone https://github.com/labulaka-ai/docx-format-guard.git
cd docx-format-guard
python -m venv .venv && .venv/bin/pip install -r requirements.txt
# 可选：LibreOffice（pages、目录页码、渲染回归测试需要）
sudo apt-get install libreoffice-writer   # macOS: brew install --cask libreoffice
```

常用命令：

```bash
# 检查
.venv/bin/python -m guard.cli check thesis.docx --profile profiles/pku-master.yaml
# 修复（输出新文件；有 LibreOffice 时同时计算目录页码，--no-toc-pages 关闭）
.venv/bin/python -m guard.cli fix thesis.docx -o thesis-fixed.docx --profile profiles/pku-master.yaml
# 导师改稿后，看哪些段落的格式被改了
.venv/bin/python -m guard.cli diff thesis-v1.docx thesis-v2.docx
# 列出修订与批注，回复第 0 条批注
.venv/bin/python -m guard.cli review thesis.docx
.venv/bin/python -m guard.cli reply thesis.docx -o thesis-replied.docx --comment 0 --text "已补充引用。" --author 张三
# 逐页核对版面
.venv/bin/python -m guard.cli pages thesis-fixed.docx
```

使用北大官方模板（可选，本地）：

```bash
mkdir -p local-assets/pku
# 从上面的模板页面下载 .doc，保存为 local-assets/pku/master-2024.doc，然后转成 .docx
soffice --headless --convert-to docx --outdir local-assets/pku local-assets/pku/master-2024.doc
```

`demo/` 里是用自建样例生成的演示文件：`broken.docx`（注入漂移）→ `fixed.docx`（修复后）；`reviewed.docx`（带导师修订和批注）→ `r1.docx`（回复批注）→ `r2.docx`（在修订上加批注）。

## 作为 agent skill 使用

仓库根目录的 `SKILL.md` 是给 AI agent（Claude Code 及兼容的 skill 运行时）的入口：什么时候触发、按什么顺序调用 `review → check → fix → check → pages`、哪些事不许做（不覆盖原稿、不删修订批注、不代改内容层面的问题、不把 LibreOffice 的页码说成 Word 的页码）。

把本仓库克隆到 agent 的 skills 目录（例如 `~/.claude/skills/docx-format-guard/`），并按上面的步骤装好依赖即可。触发示例：

- 「导师改完稿，帮我把论文格式按北大要求修回去，别动他的批注」
- 「看看这两版之间哪些段落格式变了」
- 「把导师的批注列出来，第 3 条帮我回复‘已修改’」
- 「检查一下目录页码和页眉对不对，有没有空白页」

## 可信度：测试与实测

**测试体系**（本地全套 `pytest` 实际结果：**100 passed**，0 failed，0 skipped）：

- **注入漂移**：在合规稿上注入 9 类真实漂移（直接格式覆盖、中西文字体错配、主题字体、外来粘贴样式、行距/页边距、页眉、题注、标题、参考文献），逐类断言能检出、修复后逐段与 golden 一致；全部叠加再加修订批注也能修回；
- **修订与批注**：fix 处理修订插入内部的文字并保留修订标记；`review` 覆盖插入、删除、移动、格式修订、段落标记修订等；`reply` 写出的回复挂在父批注下；
- **Word 结构校验**：自写的 OOXML 结构检查（未引入 XSD），fix / reply 的输出在测试中都要通过；
- **LibreOffice 渲染回归**：修复稿渲染为 PDF 与 golden 逐页比对；LibreOffice 读入再存出后修订批注仍在；
- **北大官方模板自检**：当前规则包在官方模板（.doc 经 LibreOffice 转 .docx）上报出 199 处问题，**全部归因**，测试保证新出现未归因的问题即失败。其中“模板与规范不一致”154 处，其余为模板结构、模板自述（示例文献未核准、目录为空占位）、模板占位内容、规则口径待确认（2 处）、环境兼容等；
- **`diff`**：注入漂移必报、自身对比为空、修复后与 golden 对比为空。

CI（GitHub Actions）装 LibreOffice 后运行；北大官方模板不入库，依赖模板的 54 个用例在 CI 中自动跳过，其余 46 个照常运行。

**WPS 实测**（WPS for Linux，缺宋体/黑体/仿宋，用 Noto CJK 等替代）：封面在一页内、无空白页、原创性声明不进目录、目录页码正确、标题编号与奇偶页页眉正确、声明签名栏不再单独成页；可见页与 LibreOffice 排版逐页一致。

**奇偶页“虚拟页”说明**：WPS 状态栏比可见页多 1 页，原因是目录止于页码 V、第一章从 1 起，两页页码同为奇数。在“奇偶页不同 + 对称页边距”（指南 1.11 要求）下，双面打印时两页之间应空一面，WPS 把它计入总页数但不显示。这是合规设置下的正常行为，工具不为“页数对上”而去掉这些设置；`pages` 命令会提示这种情况。

**Word**：批注回复在 Word 中显示为同一线程（Word 截图人工确认）。

## 已知限制

- **目录页码按 LibreOffice 排版计算**，Word 排版可能差一页；打开后可在 Word 里“更新域”；
- **替代字体环境**：为避免缺字体时版面溢出，封面、版权声明、原创性声明按 Word 下的行高固定；Word 在文档网格下的行高算法未逐项实测；
- 规则包目前只有**北大硕士**一个真实学校；博士、其他学校需另写规则包；
- 官方模板原件是 .doc，经 LibreOffice 转换后测试，个别偏差（如奇数页页眉为空、无对称页边距）可能是转换造成的；
- 参考文献按 GB/T 7714 解析，属启发式；顺序编码制与著者-出版年制混用、无法解析的条目只给建议；
- 页眉页脚和脚注中的修订暂不处理；回复批注后不会把父批注标为“已解决”；
- `diff` 按文字对齐段落，只比较段落格式，不比较表格、页眉页脚和分节；
- 模板占位空段、空章节导致的“几乎空的页”只提示，不删除用户段落。

## 与 Anthropic docx skill 的关系

无关。本项目是独立实现（python-docx + lxml + LibreOffice + PyMuPDF），没有使用 anthropics/skills 中 docx skill 的任何代码、提示词、模板或 XSD。

## 文件结构

```
docx-format-guard/
├── SKILL.md                 # agent 入口：触发场景、调用顺序、执行规则
├── README.md / README.en.md
├── LICENSE                  # Apache-2.0（代码）
├── requirements.txt
├── guard/
│   ├── cli.py               # check / fix / diff / review / reply / pages
│   ├── core.py              # 有效格式计算、角色识别、检查与修复主流程
│   ├── features.py          # 页面、页眉页码分节、封面、摘要、目录、脚注、表格、公式、编号、版面兼容
│   ├── references.py        # GB/T 7714 参考文献
│   ├── revisions.py         # 修订与批注：列出、回复、在修订上加批注
│   ├── diff.py              # 改稿前后格式差异
│   ├── layout.py            # LibreOffice 逐页摘要、签名栏兜底、奇偶页提示
│   └── toc_update.py        # LibreOffice 排版 → 目录页码回写
├── profiles/
│   ├── pku-master.yaml      # 北京大学硕士规则包
│   ├── example-master.yaml  # 通用示例
│   ├── LICENSE.md           # 规则包 CC BY 4.0 说明
│   └── LICENSE-CC-BY-4.0.txt
├── tests/                   # pytest；北大用例需本地模板，缺失时跳过
├── demo/                    # 自建样例的演示文件
└── .github/workflows/test.yml
```

## License

- 代码：Apache-2.0，见 `LICENSE`。
- 规则包（`profiles/*.yaml`）：CC BY 4.0，见 `profiles/LICENSE.md`。规则包中引用的指南条款原文版权归原发布单位，引用仅用于标注出处。
- 学校官方模板与指南不随本仓库分发。本项目与北京大学无隶属关系，规则以学校最新发布的官方文件为准。

---

**作者**：Labulaka（[@labulaka-ai](https://github.com/labulaka-ai)），时空大数据方向博士在读
**公众号**：「时空AI深水区」——时空数据 × AI 的论文解读与一手实测
