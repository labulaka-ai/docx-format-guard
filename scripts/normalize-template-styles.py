#!/usr/bin/env python3
"""把 docx 里 styles.xml 的内置样式名从「UI 式」归一化为「内部式」。

背景（2026-10-21 实测定案）：LibreOffice 24.2 把官方 .doc 转成 .docx 时，
内置样式的 w:name 写成 UI 形式（"Heading 1"、"TOC 1"、"Caption"），
而 MS Word 写的是内部形式（"heading 1"、"toc 1"、"caption"）。
python-docx 按内部名解析样式（BabelFish.ui2internal），于是 LO 转换稿上
按名查找全部落空：guard 的 check 误报 toc_style_missing，fix 则在
_ensure_style 里新建同名“幽灵样式”、把修复写进无人使用的样式，
静默残留 86 条可修复问题（详见 docs/test-environment.md）。

本脚本只改 word/styles.xml 中 <w:style> 的 <w:name w:val="..."/>：
styleId 与段落引用（w:pStyle 按 styleId）不动，幂等可重复执行。

用法：python3 scripts/normalize-template-styles.py <in.docx> [out.docx]
"""
import re
import shutil
import sys
import zipfile

RENAMES = {
    **{f"Heading {i}": f"heading {i}" for i in range(1, 10)},
    "Caption": "caption",
    "Header": "header",
    "Footer": "footer",
    "TOC 1": "toc 1",
    "TOC 2": "toc 2",
    "TOC 3": "toc 3",
}


def normalize(src: str, dst: str) -> int:
    zin = zipfile.ZipFile(src)
    changed = 0
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "word/styles.xml":
                text = data.decode("utf-8")
                for ui, internal in RENAMES.items():
                    new, n = re.subn(
                        rf'(<w:name w:val="){re.escape(ui)}("/>)',
                        rf"\g<1>{internal}\g<2>",
                        text,
                    )
                    text, changed = new, changed + n
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return changed


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    if dst == src:
        tmp = src + ".norm-tmp"
        n = normalize(src, tmp)
        shutil.move(tmp, src)
    else:
        n = normalize(src, dst)
    print(f"normalize-template-styles: {src} -> {dst}: {n} style name(s) rewritten")
