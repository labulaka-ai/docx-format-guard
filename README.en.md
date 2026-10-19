<p align="right"><a href="https://github.com/labulaka-ai/docx-format-guard/blob/main/README.md">简体中文</a> | <strong>English</strong></p>

# docx-format-guard · Word Formatting Guard for Theses

![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg) ![Rules](https://img.shields.io/badge/rules-CC%20BY%204.0-lightgrey.svg) ![Agent Skill](https://img.shields.io/badge/Agent%20Skill-SKILL.md-green) ![Docs](https://img.shields.io/badge/docs-中文%20%7C%20English-orange)

> Author: **Labulaka** (GitHub [@labulaka-ai](https://github.com/labulaka-ai)) · WeChat Official Account: 「时空AI深水区」 (Spatiotemporal AI Deep Dive)

[Why I built this](#why-i-built-this) · [Features](#features) · [PKU rule pack](#peking-university-rule-pack) · [Installation and use](#installation-and-use) · [As an agent skill](#as-an-agent-skill) · [Trust](#how-far-to-trust-it-tests-and-real-world-checks) · [Known limitations](#known-limitations)

I'm a PhD candidate working on spatiotemporal big data. This tool targets the most tedious part of writing a thesis: **the content is done, but the Word formatting has fallen apart**. It checks a .docx against a school's rule pack and puts drifted formatting back, **without touching your advisor's tracked changes and comments**. It can also list and reply to those comments, and lay the document out with LibreOffice to check it page by page.

It is not a one-click "beautify" template. It is a set of rules with sources: every rule names the clause of the school's guide it comes from. Anything the tool cannot or should not change (mixing two citation systems, for example) is reported as a suggestion with a reason. Your content is never changed silently.

---

## Why I built this

Between the first draft and the final version, a thesis goes through advisor edits, merged drafts, text pasted from elsewhere, and trips between Word, WPS, and LibreOffice. The formatting failures are remarkably consistent:

- **Drift after revisions**: pasted paragraphs arrive in 14pt KaiTi or Arial; a foreign style such as "Normal (Web)" appears; line spacing, first-line indents, and margins change quietly;
- **Reformatting destroys review marks**: the usual "select all → clear formatting" or re-applying a template wipes out the advisor's tracked changes and comment threads with it;
- **Front matter and page numbers are the hardest part**: cover, abstracts, table of contents, and declaration each need different headers and page numbering. One wrong section break and the odd-page headers, Roman/Arabic numbering, and TOC page numbers all go wrong;
- **Different software, different layout**: without SimSun/SimHei/FangSong installed, the cover spills onto two pages and the signature block gets pushed onto a page of its own.

This tool exists for those four problems: check first, then fix, re-check every item after fixing, and verify the layout from an actual rendering.

## Features

| Command | What it does |
|---|---|
| `ingest` | Derives a **draft rule pack** from a school's template .docx: page setup, styles (fonts, size, line spacing, indents, alignment), header and page-number formats are read from the template; rules that exist only in the written guide (region semantics, reference-list rules, footnote numbering) are reported as not derived. Drafts must be reviewed against the guide before they are trusted |

| `check` | Checks formatting against a rule pack; each issue reports expected value, actual value, cause, and the guide clause; `--json` for machine-readable output |
| `fix` | Fixes what can be fixed and writes a new file (the original is never overwritten); keeps tracked changes (w:ins / w:del / formatting changes) and comments; re-checks afterwards and lists what is left |
| `diff` | Compares two versions of a .docx, aligns paragraphs by text, and reports **formatting** changes only (style, CJK/Latin font, size, bold, line spacing, spacing before/after, indents, alignment) |
| `review` | Lists every tracked change and comment (author, time, paragraph, reply threads) |
| `reply` | Replies to a comment (shown as one thread in Word), or attaches a comment to a tracked change |
| `pages` | Lays the document out with LibreOffice and prints, per page, the header, footer (page number), and the first and last words of the body; flags blank and nearly empty pages, and the "virtual blank page" a duplex printout needs when two consecutive page numbers are both odd |

Issues come in two kinds: **fixable** (handled by `fix`) and **suggestion only** (content-level problems such as mixed citation systems, keyword counts, or template placeholder blank lines; the tool explains them but does not rewrite your content).

## Peking University rule pack

`profiles/pku-master.yaml`: Peking University master's thesis. Sources:

- *Peking University Graduate Thesis Writing Guide*: <https://grs.pku.edu.cn/docs/2019-09/20190911111739293567.pdf> (download page dated 2018-03-01)
- *Master's Thesis Format Template (2024)*: <https://grs.pku.edu.cn/xwgz11/xwsy11/ssxw111/clxz08/50118yjsy346393.htm> (published 2024-02-29)

Coverage: page setup (A4, margins, gutter, mirror margins); odd/even headers (odd pages carry the chapter title via a STYLEREF field); the header rule; sectioned page numbering (none on cover and declaration, Roman numerals from the abstract to the TOC, Arabic from chapter 1); cover; copyright notice and originality declaration (fixed text); Chinese and English abstracts with keywords; TOC (three levels, front matter excluded, page numbers written back); automatic heading numbering (第一章 / 1.1 / 1.1.1); body and heading fonts, sizes, and spacing; footnotes (①② restarting each page); tables and figure/table captions; equation numbering; GB/T 7714 references. Every rule carries a `clause` field naming the guide clause.

The official template is **not distributed** with this repository. Features that need it (fixed-text comparison, template self-check tests) expect you to download it into `local-assets/` (git-ignored).

`profiles/example-master.yaml` is a generic example rule pack that does not correspond to any school; it shows how a rule pack is written.

## Installation and use

```bash
git clone https://github.com/labulaka-ai/docx-format-guard.git
cd docx-format-guard
python -m venv .venv && .venv/bin/pip install -r requirements.txt
# Optional: LibreOffice (needed for pages, TOC page numbers, and rendering regression tests)
sudo apt-get install libreoffice-writer   # macOS: brew install --cask libreoffice
```

Common commands:

```bash
# Check
# No rule pack for your school yet? Derive a draft from its official template, then review it against the written guide
.venv/bin/python -m guard.cli ingest school-template.docx -o profiles/school-draft.yaml

.venv/bin/python -m guard.cli check thesis.docx --profile profiles/pku-master.yaml
# Fix (writes a new file; with LibreOffice present, TOC page numbers are computed too; --no-toc-pages turns that off)
.venv/bin/python -m guard.cli fix thesis.docx -o thesis-fixed.docx --profile profiles/pku-master.yaml
# After your advisor's edits: which paragraphs changed format?
.venv/bin/python -m guard.cli diff thesis-v1.docx thesis-v2.docx
# List tracked changes and comments; reply to comment 0
.venv/bin/python -m guard.cli review thesis.docx
.venv/bin/python -m guard.cli reply thesis.docx -o thesis-replied.docx --comment 0 --text "Citation added." --author "Zhang San"
# Check the layout page by page
.venv/bin/python -m guard.cli pages thesis-fixed.docx
```

Using the official PKU template (optional, local only):

```bash
mkdir -p local-assets/pku
# Download the .doc from the template page above as local-assets/pku/master-2024.doc, then convert it
soffice --headless --convert-to docx --outdir local-assets/pku local-assets/pku/master-2024.doc
```

`demo/` holds demonstration files generated from a self-built sample: `broken.docx` (injected drift) → `fixed.docx` (after fixing); `reviewed.docx` (advisor's tracked changes and comments) → `r1.docx` (comment replied) → `r2.docx` (comment attached to a tracked change).

## As an agent skill

`SKILL.md` at the repository root is the entry point for AI agents (Claude Code and compatible skill runtimes). It covers when to trigger, the call order `review → check → fix → check → pages`, and what is off limits: never overwrite the original, never drop tracked changes or comments, never rewrite content-level issues, never present LibreOffice page numbers as Word's.

Clone this repository into your agent's skills directory (for Claude Code, e.g. `~/.claude/skills/docx-format-guard/`) and install the dependencies as above. Example triggers:

- "My advisor finished editing — restore the PKU formatting, but don't touch his comments"
- "Which paragraphs changed format between these two versions?"
- "List my advisor's comments and reply 'done' to number 3"
- "Check the TOC page numbers and headers — any blank pages?"

## How far to trust it: tests and real-world checks

**Test suite** (the full local `pytest` run: **100 passed**, 0 failed, 0 skipped):

- **Injected drift**: 9 kinds of real-world drift (direct formatting overrides, CJK/Latin font mismatch, theme fonts, foreign pasted styles, line spacing/margins, header, captions, headings, references) are injected into a compliant document; each must be detected, and after fixing every paragraph must match the golden copy. All nine together, plus tracked changes and comments, must also come back clean;
- **Tracked changes and comments**: `fix` reformats text inside tracked insertions and keeps the revision marks; `review` covers insertions, deletions, moves, formatting changes, and paragraph-mark revisions; `reply` attaches replies under the parent comment;
- **Word structure validation**: a self-written OOXML structure check (no XSD bundled); the output of `fix` and `reply` must pass it in the tests;
- **LibreOffice rendering regression**: the fixed document is rendered to PDF and compared page by page with the golden copy; tracked changes and comments survive a LibreOffice open-and-save round trip;
- **PKU official template self-check**: on the official template (.doc converted to .docx with LibreOffice), the current rule pack reports 199 issues, **all attributed**; the test fails if a new unattributed issue appears. 154 are "template deviates from the guide"; the rest are template structure, the template's own notes (unverified sample references, an empty placeholder TOC), template placeholder content, rule interpretations still to be confirmed (2), and environment compatibility;
- **`diff`**: injected drift is always reported, a document compared with itself yields nothing, and a fixed document compared with the golden copy yields nothing.

CI (GitHub Actions) installs LibreOffice and runs the suite. The official PKU template is not in the repository, so the 54 tests that need it are skipped in CI; the other 46 run normally.

**Real-world check in WPS** (WPS for Linux, without SimSun/SimHei/FangSong, substituted by Noto CJK and others): the cover fits on one page, there are no blank pages, the originality declaration stays out of the TOC, TOC page numbers are correct, heading numbers and odd/even headers are correct, and the declaration's signature block no longer lands on a page of its own. The visible pages match the LibreOffice layout page for page.

**About the odd/even "virtual page"**: the WPS status bar shows one page more than is visible. The TOC ends on page V and chapter 1 starts on page 1, so two consecutive page numbers are both odd. With "different odd and even pages" plus mirror margins (both required by guide clause 1.11), a duplex printout needs a blank side between them; WPS counts it in the total but does not display it. This is correct behaviour under compliant settings, and the tool does not remove those settings just to make the page count match. `pages` reports the situation.

**Word**: comment replies show up as a single thread in Word (checked by hand from a Word screenshot).

## Known limitations

- **TOC page numbers come from the LibreOffice layout**; Word's layout may differ by a page. Use "Update field" in Word after opening;
- **Substitute fonts**: to keep the layout from overflowing when fonts are missing, line heights on the cover, copyright notice, and originality declaration are pinned to their Word values. Word's line-height rules on the document grid have not been verified case by case;
- The only complete rule pack shipped so far is **PKU master's** (with guide clause citations). For other schools, `ingest` derives a draft pack from the official template for review against the written guide — but a template can deviate from its own school's guide (the PKU template itself accounts for 154 measured deviations), so derived packs are drafts, never guarantees;
- The official template is a .doc, tested after LibreOffice conversion; a few deviations (empty odd-page headers, no mirror margins) may come from the conversion;
- References are parsed against GB/T 7714 heuristically; mixed numeric/author-year systems and unparseable entries only get suggestions;
- Tracked changes inside headers, footers, and footnotes are not yet handled; replying to a comment does not mark the parent as resolved;
- `diff` aligns paragraphs by text and compares paragraph formatting only, not tables, headers/footers, or sections;
- "Nearly empty" pages caused by template placeholder blank lines or empty chapters are reported, never fixed by deleting your paragraphs.


## Repository layout

```
docx-format-guard/
├── SKILL.md                 # Agent entry point: triggers, call order, operating rules
├── README.md / README.en.md
├── LICENSE                  # Apache-2.0 (code)
├── requirements.txt
├── guard/
│   ├── cli.py               # check / fix / diff / review / reply / pages / ingest

│   ├── ingest.py            # school template .docx -> draft rule pack + derivation report
│   ├── core.py              # Effective formatting, role detection, check and fix pipeline
│   ├── features.py          # Page setup, headers/page-number sections, cover, abstracts, TOC, footnotes, tables, equations, numbering, layout compatibility
│   ├── references.py        # GB/T 7714 references
│   ├── revisions.py         # Tracked changes and comments: list, reply, comment on a change
│   ├── diff.py              # Formatting diff between two versions
│   ├── layout.py            # LibreOffice per-page summary, signature-block fallback, odd/even note
│   └── toc_update.py        # LibreOffice layout → TOC page numbers written back
├── profiles/
│   ├── pku-master.yaml      # Peking University master's rule pack
│   ├── example-master.yaml  # Generic example
│   ├── LICENSE.md           # CC BY 4.0 notice for rule packs
│   └── LICENSE-CC-BY-4.0.txt
├── tests/                   # pytest; PKU tests need the local template and are skipped without it
├── demo/                    # Demonstration files from a self-built sample
└── .github/workflows/test.yml
```

## License

- Code: Apache-2.0, see `LICENSE`.
- Rule packs (`profiles/*.yaml`): CC BY 4.0, see `profiles/LICENSE.md`. Guide clauses quoted in the rule packs remain the copyright of their publishers and are quoted only to cite the source.
- Official school templates and guides are not distributed with this repository. This project is not affiliated with Peking University; the school's latest official documents always take precedence.

---

**Author**: Labulaka ([@labulaka-ai](https://github.com/labulaka-ai)), PhD candidate in spatiotemporal big data
**WeChat Official Account**: 「时空AI深水区」 — paper explainers and first-hand tests at the intersection of spatiotemporal data and AI
