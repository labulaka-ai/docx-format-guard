"""排版后逐页摘要：用 LibreOffice 转 PDF，每页给出页眉、页脚（页码）、正文前几个字、正文行数，标出空白页/几乎空的页。"""
import os, re, shutil, subprocess, tempfile

PAGE_NO = re.compile(r"^([IVXLCDM]+|[ivxlcdm]+|\d+)$")

def render_pdf(docx, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, docx], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
    return os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")

def page_summary(pdf, header_bottom_pt=80):
    import pymupdf
    res = []
    for i, pg in enumerate(pymupdf.open(pdf)):
        H = pg.rect.height; hd, ft, bd = [], [], []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if not t: continue
                y0, y1 = l["bbox"][1], l["bbox"][3]
                if y1 <= header_bottom_pt: hd.append((y0, t))
                elif y0 > H * 0.88 and PAGE_NO.match(t): ft.append((y0, t))
                else: bd.append((y0, t))
        bd.sort()
        res.append(dict(page=i + 1, header=" ".join(t for _, t in sorted(hd)), footer=" ".join(t for _, t in ft),
                        lines=len(bd), first=bd[0][1][:16] if bd else "", last=bd[-1][1][:16] if bd else "",
                        blank=not bd, nearly_empty=0 < len(bd) <= 2))
    return res

def summarize(docx):
    tmp = tempfile.mkdtemp()
    try: return page_summary(render_pdf(docx, tmp))
    finally: shutil.rmtree(tmp, ignore_errors=True)

def _sig_orphans(pdf, rx, first_text):
    """返回在 PDF 中“整页正文只有签名栏/日期”的签名段序号（按出现顺序计数）。"""
    rows = []; import pymupdf
    started = False; k = -1; orphan = set()
    for pg in pymupdf.open(pdf):
        H = pg.rect.height; body = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if not t or l["bbox"][3] <= 80 or (l["bbox"][1] > H * 0.88 and PAGE_NO.match(t)): continue
                body.append(t)
        if not started and any(t.startswith(first_text) for t in body): started = True
        if not started: continue
        idx = []
        for t in body:
            if rx.match(t): k += 1; idx.append(k)
        if body and len(idx) == len(body): orphan.update(idx)
    return orphan

def fix_sig_orphans(path, profile, changes):
    """排版兜底：声明块超过一页时“与下段同页”链会在末尾断开，签名栏仍可能单独成页。
    渲染后若发现这种页，就让该块最后一段正文“段前分页”，把它和签名栏一起带到下一页。"""
    from docx import Document
    from . import core, features
    kt = profile.get("keep_together") or {}
    if not kt: return
    rx = re.compile(kt["pattern"]); tmp = tempfile.mkdtemp()
    try:
        for _ in range(3):
            doc = Document(path); regs = {}; core.roles(doc, profile, regs)
            paras = [p for p in doc.paragraphs if regs.get(p._p) in set(kt["regions"])]
            if not paras: return
            first = next((p.text.strip()[:12] for p in paras if p.text.strip()), "")
            orph = _sig_orphans(render_pdf(path, tmp), rx, first)
            if not orph: return
            sig = [i for i, p in enumerate(paras) if rx.match(p.text.strip())]
            done = False
            for n in sorted(orph):
                if n >= len(sig): continue
                i = sig[n]
                bs = re.compile(kt["block_start"]) if kt.get("block_start") else None
                cand = []
                for j in range(i - 1, -1, -1):
                    s = paras[j].text.strip()
                    if not s or rx.match(s): continue
                    if bs and bs.match(s): break                   # 不越过块标题（否则等于整块分页，前面已证明放不下）
                    cand.append(j)
                    if len(cand) >= kt.get("orphan_min_paras", 3): break
                t = cand[-1] if cand else None
                if t is None or paras[t].paragraph_format.page_break_before: continue
                paras[t].paragraph_format.page_break_before = True; done = True
                changes.append(f"原创性声明：声明块超过一页，签名栏会单独成页 → “{paras[t].text.strip()[:12]}…”段前分页，与签名栏同页（排版相关兜底）")
            if not done: return
            doc.save(path)
    finally: shutil.rmtree(tmp, ignore_errors=True)

_ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
def _num(s):
    if s.isdigit(): return int(s)
    s = s.upper(); v = 0
    if not s or any(c not in _ROMAN for c in s): return None
    for a, b in zip(s, s[1:] + " "):
        v += -_ROMAN[a] if b != " " and _ROMAN[a] < _ROMAN[b] else _ROMAN[a]
    return v

def parity_notes(rows):
    """双面打印 + 奇偶页不同时，页码的奇偶决定左右页。相邻两页页码同为奇数（如目录止于 V、第一章从 1 起），
    阅读器/打印程序会在中间补一个“虚拟偶数页”（WPS 状态栏计数 +1，Word 打印时可能出空白页）。"""
    out = []
    for a, b in zip(rows, rows[1:]):
        x, y = _num(a["footer"]), _num(b["footer"])
        if x is not None and y is not None and x % 2 == 1 and y % 2 == 1:
            out.append(dict(after_page=a["page"], prev=a["footer"], next=b["footer"],
                            note=f"第{a['page']}页（页码 {a['footer']}）与第{b['page']}页（页码 {b['footer']}）同为奇数页码：双面打印时两页之间需一张空白偶数页，WPS 会计入总页数但不显示"))
    return out
