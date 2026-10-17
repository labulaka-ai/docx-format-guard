"""构造"格式被改乱"的 docx 测试集。每个 corruption 对应一种真实漂移场景。"""
import copy
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PARAS = [
    ("第一章 绪论", "Heading 1"),
    ("1.1 研究背景", "Heading 2"),
    ("深度学习（Deep Learning）在遥感影像分析中取得了显著进展，本文基于 Transformer 模型展开研究。", "Normal"),
    ("实验在 3 个公开数据集上进行，平均精度提升 2.5%。", "Normal"),
    ("图1-1 系统总体框架", "Caption"),
    ("1.2 本文结构", "Heading 2"),
    ("本文共分五章，各章内容安排如下。", "Normal"),
]

def base_doc():
    d = Document()
    for text, st in PARAS:
        p = d.add_paragraph(style=st); p.add_run(text)
    d.sections[0].header.paragraphs[0].add_run("示例大学硕士学位论文")
    return d

def _rfonts(run, **kw):
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for k, v in kw.items(): rf.set(qn(f"w:{k}"), v)

def body_paras(d): return [p for p in d.paragraphs if p.style.name == "Normal" and p.text.strip()]

def c_direct_override(d):
    """直接格式覆盖：从别处粘贴后 run 带了 14pt 楷体/Arial。"""
    for p in body_paras(d):
        for r in p.runs:
            r.font.size = Pt(14); _rfonts(r, eastAsia="楷体", ascii="Arial", hAnsi="Arial")
def c_font_mismatch(d):
    """中英文字体错配：中文槽位被设成 Times New Roman，西文被设成宋体。"""
    for p in body_paras(d):
        for r in p.runs: _rfonts(r, eastAsia="Times New Roman", ascii="宋体", hAnsi="宋体")
def c_theme_font(d):
    """主题字体引用：粘贴自其他模板，字体指向 minorEastAsia 主题，且无显式 eastAsia。"""
    for p in body_paras(d):
        for r in p.runs:
            rpr = r._r.get_or_add_rPr(); rf = OxmlElement("w:rFonts")
            rf.set(qn("w:asciiTheme"), "minorHAnsi"); rf.set(qn("w:eastAsiaTheme"), "minorEastAsia")
            rf.set(qn("w:eastAsia"), "等线"); rpr.insert(0, rf)
def c_pasted_style(d):
    """样式混乱：粘贴带入 'Pasted Normal'（微软雅黑 11pt 单倍行距），标题被降成正文+手动加粗。"""
    st = d.styles.add_style("Pasted Normal", WD_STYLE_TYPE.PARAGRAPH)
    st.font.size = Pt(11); st.font.name = "Calibri"; st.element.get_or_add_rPr()
    rf = st.element.rPr.find(qn("w:rFonts")); rf.set(qn("w:eastAsia"), "微软雅黑")
    st.paragraph_format.line_spacing = 1.0
    for p in body_paras(d): p.style = st
    h = [p for p in d.paragraphs if p.style.name == "Heading 2"][0]
    h.style = d.styles["Normal"]
    for r in h.runs: r.bold = True; r.font.size = Pt(14)
def c_spacing_margin(d):
    """行距/缩进/页边距被改：1.5 倍行距、无首行缩进、页边距 2cm。"""
    for p in body_paras(d):
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.first_line_indent = Pt(0)
    for s in d.sections:
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2.0)
def c_header(d):
    """页眉被改成别人的论文题目。"""
    hp = d.sections[0].header.paragraphs[0]
    for r in hp.runs: r.text = ""
    hp.runs[0].text = "某某大学本科毕业设计"
def c_caption(d):
    """题注被改成正文样式、左对齐、12pt。"""
    for p in d.paragraphs:
        if p.text.startswith("图"):
            p.style = d.styles["Normal"]; p.alignment = 0
def add_revision(d):
    """导师修订：在第一段正文插入一处 tracked insertion。"""
    p = body_paras(d)[0]
    ins = OxmlElement("w:ins"); ins.set(qn("w:id"), "901"); ins.set(qn("w:author"), "导师"); ins.set(qn("w:date"), "2026-10-17T00:00:00Z")
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr"); sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "30"); rpr.append(sz); r.append(rpr)
    t = OxmlElement("w:t"); t.text = "（导师补充）"; r.append(t); ins.append(r); p._p.append(ins)

CORRUPTIONS = {
    "direct_override": (c_direct_override, {"size", "font_east_asia", "font_latin"}),
    "font_mismatch": (c_font_mismatch, {"font_east_asia", "font_latin"}),
    "theme_font": (c_theme_font, {"font_east_asia"}),
    "pasted_style": (c_pasted_style, {"style", "line_spacing"}),
    "spacing_margin": (c_spacing_margin, {"line_spacing", "indent", "margin"}),
    "header": (c_header, {"header"}),
    "caption": (c_caption, {"style", "align"}),
}

def build(path, corruptions=(), revision=False):
    d = base_doc()
    for c in corruptions: CORRUPTIONS[c][0](d)
    if revision: add_revision(d)
    d.save(path); return path

# ---------- 带修订与批注的样本（模拟导师用 Word 审阅后的稿子） ----------
def _el(tag, **attrs):
    e = OxmlElement(f"w:{tag}")
    for k, v in attrs.items(): e.set(qn(f"w:{k}"), v)
    return e
def _run(text, deleted=False, sz=None):
    r = _el("r")
    if sz: rpr = _el("rPr"); rpr.append(_el("sz", val=str(sz))); r.append(rpr)
    t = _el("delText" if deleted else "t"); t.text = text; t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); r.append(t)
    return r
REV_META = dict(author="导师", date="2026-10-16T08:00:00Z")

def build_reviewed(path):
    d = base_doc()
    bp = body_paras(d)
    ins = _el("ins", id="101", **REV_META); ins.append(_run("（导师补充：需注明数据来源）", sz=30)); bp[0]._p.append(ins)
    dl = _el("del", id="102", **REV_META); dl.append(_run("平均精度提升 2.5%。", deleted=True))
    r0 = bp[1].runs[0]; r0.text = "实验在 3 个公开数据集上进行，"; r0._r.addnext(dl)
    r = bp[2].runs[0]; r.bold = True                         # 字符格式修订：被设为加粗
    ch = _el("rPrChange", id="103", **REV_META); ch.append(_el("rPr")); r._r.rPr.append(ch)
    ppr = bp[2]._p.get_or_add_pPr(); pch = _el("pPrChange", id="104", **REV_META)
    old = _el("pPr"); old.append(_el("jc", val="left")); pch.append(old); ppr.append(pch)
    mark = _el("rPr"); mark.append(_el("ins", id="105", **REV_META)); pch.addprevious(mark)   # schema 顺序：rPr 在 pPrChange 之前
    d.add_comment(bp[0].runs, text="这段缺少引用。", author="导师", initials="DS")
    d.save(path); return path
