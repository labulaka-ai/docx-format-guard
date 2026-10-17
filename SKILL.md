---
name: docx-format-guard
description: >-
  学位论文 Word（.docx）格式守护：按学校规则包检查并修复改稿后的格式漂移，
  保留导师的修订与批注；列出、回复批注；比较两版之间的格式变化；用 LibreOffice 逐页核对版面。
  触发场景：「论文格式乱了 / 按学校要求修格式」「导师改完稿，别动他的批注」「这两版哪些段落格式变了」
  「把批注列出来 / 回复第 N 条批注」「目录页码、页眉、空白页对不对」。内置北京大学硕士规则包。
  不用于：写或改论文内容（那是 paper-writing 的事）、LaTeX 论文、.doc 旧格式（先转 .docx）。
---

# docx-format-guard（学位论文 Word 格式守护）

核心原则：**只改格式，不改内容；不覆盖原稿；不丢修订和批注；改不了的说清楚为什么。**

## 准备

- 依赖：`pip install -r requirements.txt`；`pages` 与目录页码需要 LibreOffice（`soffice` 在 PATH 中）。
- 规则包：`profiles/pku-master.yaml`（北京大学硕士）。用户的学校没有规则包时，如实告知，可参照 `profiles/example-master.yaml` 起草，但起草的规则必须注明来源，且在用户确认前不得称为“学校要求”。
- 所有命令在本仓库根目录运行：`python -m guard.cli <命令> ...`

## 总流程

1. **先看批注和修订**：`review thesis.docx`。告诉用户有多少处修订、多少条批注，谁写的。
2. **检查**：`check thesis.docx --profile <规则包>`（需要时加 `--json`）。按问题类别汇总，不要逐条倾倒；区分“可修复”和“只给建议”。
3. **修复到新文件**：`fix thesis.docx -o thesis-fixed.docx --profile <规则包>`。永远输出到新文件。
4. **复查**：fix 结束时会打印剩余问题；把剩余项连同原因交给用户（通常是参考文献著录制混用、关键词个数、模板占位空段等内容层面的问题）。
5. **核对版面**（有 LibreOffice 时）：`pages thesis-fixed.docx`。报告空白页、几乎空的页，以及“相邻页码同为奇数”的提示。
6. **改稿前后对比**（用户给了两版时）：`diff v1.docx v2.docx`，按段落列出格式变化。
7. **回复批注**（用户明确要求、并给出回复内容时）：`reply in.docx -o out.docx --comment ID --text "..." --author <用户名>`；对修订加批注用 `--revision ID`。

## Operating Rules

1. 不覆盖原稿：`fix` / `reply` 一律 `-o` 到新文件。
2. 不删修订、不删批注、不接受或拒绝修订；这些是用户和导师的决定。
3. 只给建议的问题不代改：说明原因和建议，由用户改内容。
4. 不编造规则：每条规则以规则包里的 `clause` 为准；规则包没写的，不说成“学校要求”。
5. 页码口径：目录页码按 LibreOffice 排版计算，Word / WPS 可能差一页；不要把它说成 Word 的页码。WPS 状态栏比可见页多 1 页时，先看 `pages` 的奇偶页提示，不要为了页数对上去掉“奇偶页不同 / 对称页边距”等合规设置。
6. 回复批注的作者名用用户提供的名字；未提供时询问，不要默认填写。
7. 学校官方模板只放在本地 `local-assets/`，不提交、不外传。
