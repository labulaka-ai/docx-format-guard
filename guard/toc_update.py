"""用 LibreOffice 排版得到各标题所在页，把页码写回 TOC 域的缓存结果（不依赖 UNO，只用 soffice 转 PDF + PyMuPDF）。
页码只对 LibreOffice 排版准确；Word 排版可能差 1 页，故同时设置 updateFields，Word 打开时会提示更新域。"""
import os, re, subprocess, tempfile, shutil
from docx import Document
from . import core, features

def _render(path, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, path], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")

def heading_pages(pdf, entries):
    import pymupdf
    d = pymupdf.open(pdf); lines = []
    for i, pg in enumerate(d):
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                lines.append((i, re.sub(r"\s+", "", "".join(s["text"] for s in l["spans"]))))
    res = {}; pos = 0
    toc_head = next((j for j, (_, t) in enumerate(lines) if t == "目录"), None)
    if toc_head is not None:                 # 目录页里的条目先出现一次，跳过
        pos = toc_head + 1
        last = max((j for j, (_, t) in enumerate(lines) if j > toc_head and re.search(r"\.{8,}", t)), default=toc_head)
        pos = last + 1
    def hit(j, key):
        a = lines[j][1]
        if a == key or (a.startswith(key) and len(a) - len(key) <= 3): return True     # 标题末尾可能带脚注号
        if j + 1 < len(lines) and lines[j + 1][0] == lines[j][0] and len(a) >= 4 and key.startswith(a):
            b = a + lines[j + 1][1]                                                       # 长标题折行
            return b == key or (b.startswith(key) and len(b) - len(key) <= 3)
        return False
    for lv, txt in entries:
        key = re.sub(r"\s+", "", txt)
        for j in range(pos, len(lines)):
            if hit(j, key): res[key] = lines[j][0]; pos = j + 1; break
    return res

def update_toc_cached(path, profile):
    entries = features.expected_toc(Document(path), profile)
    tmp = tempfile.mkdtemp()
    try:
        pdf = _render(path, tmp)
        phys = heading_pages(pdf, entries)
        if not phys: return {}
        start = min(phys.values())                       # 正文第一页 = 阿拉伯页码 1
        pages = {k: v - start + 1 for k, v in phys.items()}
        doc = Document(path); features.rebuild_toc_result(doc, profile, pages); doc.save(path)
        return pages
    finally: shutil.rmtree(tmp, ignore_errors=True)
