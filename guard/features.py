"""北大规则包所需的结构性检查与修复：页面、页眉页脚页码、封面、固定文本、摘要、目录、脚注、
表格与题注、公式、标题自动编号、参考文献（GB/T 7714-2015）。
约定：每个 check_* 产出 issue（fixable=False 的为"只能给修改建议"）；fix_* 就地修改 doc。"""
import copy, re
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from . import core
from .references import check_references, fix_references

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def _iss(kind, where, expect, actual, cause, clause=None, fixable=True, suggestion=None):
    return dict(kind=kind, where=where, expect=expect, actual=actual, cause=cause, clause=clause,
                fixable=fixable, suggestion=suggestion)

def _adder(out, fixable=True):
    def add(kind, where, expect, actual, cause, clause=None):
        out.append(_iss(kind, where, expect, actual, cause, clause, fixable))
    return add

# ---------- 公共：段落直接格式写入（用于封面、页眉页脚等不走样式的位置） ----------
def apply_direct(p, r):
    pf = p.paragraph_format
    if "align" in r: pf.alignment = core.ALIGN[r["align"]]
    if "line_exact_pt" in r: pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = Pt(r["line_exact_pt"])
    elif r.get("line_single"): pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if "space_before_pt" in r: pf.space_before = Pt(r["space_before_pt"])
    if "space_after_pt" in r: pf.space_after = Pt(r["space_after_pt"])
    for run in core.all_runs(p):
        run.font.size = Pt(r["size_pt"]); core._set_fonts(run._r, r["east_asia"], r["latin"])
        if "bold" in r: run.bold = r["bold"]

def P(doc, p_el): return Paragraph(p_el, doc._body)

# ---------- 区段（section）与区域（region）映射 ----------
def section_map(doc):
    """返回 [(段落元素, 节序号)]，以及节总数。"""
    out, si = [], 0
    for el in doc.element.body:
        if el.tag == qn("w:p"):
            out.append((el, si))
            ppr = el.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None: si += 1
        elif el.tag == qn("w:tbl"):
            for p in el.iter(qn("w:p")): out.append((p, si))
    return out

def region_sections(doc, profile):
    regs = {}; core.roles(doc, profile, regs)
    sm = dict(section_map(doc)); res = {}
    for el, name in regs.items():
        if name is None: continue
        res.setdefault(name, set()).add(sm.get(el))
    # 封面：第一个分区之前的段落所在节
    first = {si for el, si in section_map(doc) if regs.get(el) is None and si == 0}
    if first: res.setdefault("cover", set()).update(first)
    return res, regs

# ---------- 1. 页面设置 ----------
def check_page(doc, profile, out):
    pg = profile.get("page", {}); cl = pg.get("clause")
    for i, s in enumerate(doc.sections):
        if pg.get("paper") == "A4" and (abs(s.page_width.cm - 21.0) > 0.05 or abs(s.page_height.cm - 29.7) > 0.05):
            out.append(_iss("paper", f"节{i+1}", "A4 21.0×29.7cm", f"{s.page_width.cm:.2f}×{s.page_height.cm:.2f}", "纸张大小", cl))
        if "gutter_cm" in pg and abs(s.gutter.cm - pg["gutter_cm"]) > 0.02:
            out.append(_iss("gutter", f"节{i+1}", pg["gutter_cm"], round(s.gutter.cm, 2), "装订线", cl))
    st = doc.settings.element
    if pg.get("mirror_margins") and st.find(qn("w:mirrorMargins")) is None:
        out.append(_iss("mirror_margins", "全文", True, False, "左右对称页边距（对称页边距）", cl))
    if pg.get("gutter_pos") == "left" and st.find(qn("w:gutterAtTop")) is not None:
        out.append(_iss("gutter_pos", "全文", "左", "上", "装订线位置", cl))

def fix_page(doc, profile, changes):
    pg = profile.get("page", {})
    for s in doc.sections:
        if pg.get("paper") == "A4": s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        if "gutter_cm" in pg: s.gutter = Cm(pg["gutter_cm"])
    st = doc.settings.element
    if pg.get("mirror_margins") and st.find(qn("w:mirrorMargins")) is None:
        _settings_insert(st, OxmlElement("w:mirrorMargins"))
    g = st.find(qn("w:gutterAtTop"))
    if pg.get("gutter_pos") == "left" and g is not None: st.remove(g)
    changes.append("页面：A4、装订线、对称页边距")

SETTINGS_ORDER = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries",
    "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState",
    "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
    "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
    "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone",
    "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders",
    "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
    "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
    "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields"]

def _settings_insert(st, el):
    """按 CT_Settings 的 schema 顺序插入子元素（Word 对顺序敏感）。"""
    name = el.tag.split("}")[1]; idx = SETTINGS_ORDER.index(name) if name in SETTINGS_ORDER else len(SETTINGS_ORDER)
    for child in st:
        cn = child.tag.split("}")[1]
        if cn in SETTINGS_ORDER and SETTINGS_ORDER.index(cn) > idx:
            child.addprevious(el); return
    st.append(el)

# ---------- 2. 页眉：下横线、奇数页章名域 ----------
def _has_field(el, name):
    instr = " ".join((t.text or "") for t in el.iter(qn("w:instrText")))
    instr += " ".join(f.get(qn("w:instr")) or "" for f in el.iter(qn("w:fldSimple")))
    return re.search(rf"\b{name}\b", instr, re.I) is not None, instr

def _bottom_border(doc, hp):
    for ppr in [hp._p.pPr] + [st.element.pPr for st in core._style_chain(hp.style)]:
        b = None if ppr is None else ppr.find(qn("w:pBdr"))
        if b is not None and b.find(qn("w:bottom")) is not None:
            e = b.find(qn("w:bottom")); return e.get(qn("w:val")), int(e.get(qn("w:sz")) or 0)
    return None, 0

def _header_text_paras(hdr): return [p for p in hdr.paragraphs]

def check_header_extra(doc, profile, out):
    h = profile.get("header") or {}; cl = h.get("clause")
    if not h: return
    firsts = sec_first_region(doc, profile)
    for i, s in enumerate(doc.sections):
        if i < h.get("from_section", 0) or (firsts[i] in (profile.get("page_numbers") or {}).get("none_regions", []) and firsts[i] not in (h.get("odd_styleref") or {})): continue
        for nm, hdr in (("奇数页", s.header), ("偶数页", s.even_page_header)):
            paras = [p for p in hdr.paragraphs] or []
            target = paras[0] if paras else None
            if target is None: continue
            if "bottom_border_pt" in h:
                val, sz = _bottom_border(doc, target)
                if val != "single" or abs(sz / 8 - h["bottom_border_pt"]) > 0.01:
                    out.append(_iss("header_border", f"节{i+1}{nm}", f"单线 {h['bottom_border_pt']}磅", f"{val} {sz/8}磅", "页眉下横线", cl))
        want = (h.get("odd_styleref") or {}).get(firsts[i])
        if want:
            ok, instr = _has_field(s.header._element, "STYLEREF")
            if not ok or want.lower() not in instr.lower():
                out.append(_iss("header_odd_styleref", f"节{i+1}", f'STYLEREF "{want}"', instr.strip() or "（无域）", "奇数页页眉应显示相应内容/各章名称", cl))

def _styleref_runs(p, style_name, cached):
    def r_(child): r = OxmlElement("w:r"); r.append(child); return r
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); it.text = f' STYLEREF  "{style_name}" '
    sp = OxmlElement("w:fldChar"); sp.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = cached
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    for c in (b, it, sp, t, e): p._p.append(r_(c))

def fix_header_extra(doc, profile, changes):
    h = profile.get("header") or {}
    if not h: return
    firsts = sec_first_region(doc, profile)
    if "bottom_border_pt" in h:                     # 写在 Header 样式上，并清掉段落级边框
        st = core._ensure_style(doc, "Header", changes); ppr = st.element.get_or_add_pPr()
        old = ppr.find(qn("w:pBdr"))
        if old is not None: ppr.remove(old)
        bdr = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
        bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), str(int(round(h["bottom_border_pt"] * 8)))); bt.set(qn("w:space"), "1"); bt.set(qn("w:color"), "auto")
        bdr.append(bt)
        # pBdr 在 pPr 中的位置：在 shd/tabs/spacing/ind/jc 之前，keepNext 等之后
        anchor = next((c for c in ppr if c.tag.split("}")[1] in ("shd", "tabs", "suppressAutoHyphens", "spacing", "ind", "jc", "outlineLvl", "rPr")), None)
        if anchor is not None: anchor.addprevious(bdr)
        else: ppr.append(bdr)
    for i, s in enumerate(doc.sections):
        if i < h.get("from_section", 0):
            # 封面、版权声明不应有页眉：模板的 Header 样式自带下框线，空页眉在封面顶部也会画出一条线
            for hdr in (s.header, s.even_page_header, s.first_page_header):
                if hdr.is_linked_to_previous: continue
                for p in hdr.paragraphs:
                    if p.text.strip(): continue
                    ppr = p._p.get_or_add_pPr(); old = ppr.find(qn("w:pBdr"))
                    if old is not None: ppr.remove(old)
                    bdr = parse_xml(f'<w:pBdr xmlns:w="{W}"><w:bottom w:val="nil"/></w:pBdr>')
                    anchor = next((c for c in ppr if c.tag.split("}")[1] in ("shd", "tabs", "suppressAutoHyphens", "snapToGrid", "spacing", "ind", "jc", "outlineLvl", "rPr")), None)
                    anchor.addprevious(bdr) if anchor is not None else ppr.append(bdr)
            continue
        for hdr in (s.header, s.even_page_header):
            if hdr.is_linked_to_previous: continue
            hst = core.find_style(doc, "Header")
            for p in hdr.paragraphs:
                if hst is not None and p.style.name != hst.name: p.style = hst
                p.paragraph_format.first_line_indent = Pt(0)
                ppr = p._p.pPr
                if ppr is not None and ppr.find(qn("w:pBdr")) is not None: ppr.remove(ppr.find(qn("w:pBdr")))
        want = (h.get("odd_styleref") or {}).get(firsts[i])
        if want:
            ok, instr = _has_field(s.header._element, "STYLEREF")
            if not ok or want.lower() not in instr.lower():
                s.header.is_linked_to_previous = False
                hp = s.header.paragraphs[0]
                for extra in s.header.paragraphs[1:]: extra._p.getparent().remove(extra._p)
                for c in list(hp._p):
                    if c.tag != qn("w:pPr"): hp._p.remove(c)
                cached = next((P(doc, el).text for el, si in section_map(doc) if si == i and P(doc, el).style.name == want), "")
                _styleref_runs(hp, want, cached.strip())
                hp.alignment = core.ALIGN[h["align"]]
                for run in hp.runs: run.font.size = Pt(h["size_pt"]); core._set_fonts(run._r, h["east_asia"], profile["styles"]["body"]["latin"])
    changes.append("页眉：0.75 磅下横线、奇数页 STYLEREF 章名域")

# ---------- 3. 页码分节 ----------
def _footers(s): return [("奇数页", s.footer), ("偶数页", s.even_page_footer), ("首页", s.first_page_footer)]

def sec_first_region(doc, profile):
    """每一节第一个有区域归属的段落所在区域。"""
    regs = {}; core.roles(doc, profile, regs); res = {}
    for el, si in section_map(doc):
        if si not in res and regs.get(el): res[si] = regs[el]
    return [res.get(i, "cover" if i == 0 else None) for i in range(len(doc.sections))]

def _sect_role(doc, profile):
    pn = profile.get("page_numbers") or {}; roles = {}
    for i, rg in enumerate(sec_first_region(doc, profile)):
        roles[i] = "roman" if rg in pn.get("roman_regions", []) else "arabic" if rg in pn.get("arabic_regions", []) else "none"
    return roles, None

def _needs_split(doc, profile):
    """own_section 中的区域若不是从节首开始，返回 [(区域名, 首段元素)]。"""
    pn = profile.get("page_numbers") or {}
    regs = {}; core.roles(doc, profile, regs); sm = section_map(doc); res = []
    for name in pn.get("own_section", []):
        k = next((j for j, (el, si) in enumerate(sm) if regs.get(el) == name and el.getparent() is doc.element.body), None)
        if k is None: continue
        el, si = sm[k]
        if any(s2 == si for _, s2 in sm[:k] if _.getparent() is doc.element.body): res.append((name, el))
    return res

def _sectpr_after(doc, el):
    n = el
    while n is not None:
        if n.tag == qn("w:p"):
            ppr = n.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None: return ppr.find(qn("w:sectPr"))
        n = n.getnext()
    return doc.element.body.find(qn("w:sectPr"))

def _split_before(doc, first_para_el):
    """在 first_para_el 前分节。原节的 sectPr 复制到前一段（前半节保持原页眉页脚与起始页码）；
    后半节去掉页眉页脚引用（继承前一节）与 start（页码续排），由后续步骤按区域重设。"""
    cur = _sectpr_after(doc, first_para_el)
    prev = first_para_el.getprevious()
    while prev is not None and prev.tag != qn("w:p"): prev = prev.getprevious()
    # 避免空白页：前一段为空、再前一段以分页符结尾时，去掉该分页符（分节符本身就换页）
    pp = prev
    while pp is not None and pp.tag == qn("w:p") and not _ptext(pp):
        pp = pp.getprevious()
    if pp is not None:
        brs = [br for br in pp.iter(qn("w:br")) if br.get(qn("w:type")) == "page"]
        if brs: brs[-1].getparent().remove(brs[-1])
    ppr = prev.find(qn("w:pPr"))
    if ppr is None: ppr = OxmlElement("w:pPr"); prev.insert(0, ppr)
    ppr.append(copy.deepcopy(cur))
    for ref in cur.findall(qn("w:headerReference")) + cur.findall(qn("w:footerReference")): cur.remove(ref)
    pg = cur.find(qn("w:pgNumType"))
    if pg is not None and pg.get(qn("w:start")) is not None: del pg.attrib[qn("w:start")]
    tp = cur.find(qn("w:titlePg"))
    if tp is not None: cur.remove(tp)

def check_page_numbers(doc, profile, out):
    pn = profile.get("page_numbers")
    if not pn: return
    cl = pn.get("clause"); roles, _ = _sect_role(doc, profile)
    for name, el in _needs_split(doc, profile):
        out.append(_iss("section_split", name, "单独成节", "与前文同节", "页码/页眉需与前文不同", cl))
    for i, s in enumerate(doc.sections):
        role = roles[i]; pg = s._sectPr.find(qn("w:pgNumType"))
        fmt = None if pg is None else pg.get(qn("w:fmt")); start = None if pg is None else pg.get(qn("w:start"))
        if role == "none":
            if any(_has_field(f._element, "PAGE")[0] for _, f in _footers(s)):
                out.append(_iss("pagenum_present", f"节{i+1}", "无页码", "有页码", "封面与原创声明不要页码", cl))
            continue
        want = pn["roman_fmt"] if role == "roman" else pn["arabic_fmt"]
        if (fmt or "decimal") != want: out.append(_iss("pagenum_fmt", f"节{i+1}", want, fmt, "页码格式", cl))
        first_of_kind = not any(roles[j] == role for j in range(i))
        if first_of_kind and start != "1": out.append(_iss("pagenum_start", f"节{i+1}", 1, start, "页码起始", cl))
        if not first_of_kind and start is not None: out.append(_iss("pagenum_restart", f"节{i+1}", "续排", start, "页码不应重新起始", cl))
        for nm, f in _footers(s)[:2]:
            ok = _has_field(f._element, "PAGE")[0]
            if not ok: out.append(_iss("pagenum_missing", f"节{i+1}{nm}", "PAGE 域", "无", "页脚页码", cl)); continue
            for p in f.paragraphs:
                if not _has_field(p._p, "PAGE")[0]: continue
                if core.eff_pf(p, "alignment") != WD_ALIGN_PARAGRAPH.CENTER:
                    out.append(_iss("pagenum_align", f"节{i+1}{nm}", "居中", str(core.eff_pf(p, "alignment")), "页码位置", cl))
                for run in p.runs:
                    if run._r.find(qn("w:t")) is None: continue
                    sz = core.eff_size(doc, p, run); la = core.norm_font(core.eff_font(doc, p, run, "ascii"))
                    if sz is None or abs(sz - pn["size_pt"]) > 0.25 or la != pn["latin"]:
                        out.append(_iss("pagenum_font", f"节{i+1}{nm}", f"{pn['latin']} {pn['size_pt']}", f"{la} {sz}", "页码字体字号", cl)); break
                deco = _page_decor(p._p)
                if deco: out.append(_iss("pagenum_decor", f"节{i+1}{nm}", "数字两侧无修饰", deco, "页码修饰线", cl))

def _page_decor(p_el):
    """域结果以外的可见文字（如 “- 1 -” 中的短横线）。"""
    depth = 0; in_result = False; txt = ""
    for r in p_el.iter(qn("w:r")):
        fc = r.find(qn("w:fldChar"))
        if fc is not None:
            t = fc.get(qn("w:fldCharType"))
            if t == "begin": depth += 1
            elif t == "separate": in_result = True
            elif t == "end": depth -= 1; in_result = False
            continue
        if depth == 0:
            txt += "".join(t.text or "" for t in r.findall(qn("w:t")))
    return txt.strip()

def _page_field_par(doc, pn):
    p = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
                  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    para = P(doc, p)
    for run in para.runs: run.font.size = Pt(pn["size_pt"]); core._set_fonts(run._r, "宋体", pn["latin"])
    return p

def _split_before(doc, first_para_el):
    """在 first_para_el 之前插入分节：前一段带上当前节的 sectPr 副本（保持原页眉页脚），后面成为新节。"""
    prev = first_para_el.getprevious()
    while prev is not None and prev.tag != qn("w:p"): prev = prev.getprevious()
    body_sect = doc.element.body.find(qn("w:sectPr"))
    newsp = copy.deepcopy(body_sect)
    ppr = prev.find(qn("w:pPr"))
    if ppr is None: ppr = OxmlElement("w:pPr"); prev.insert(0, ppr)
    ppr.append(newsp)
    # 原文档级 sectPr 现在描述"新节"：页眉继承（去掉 headerReference）、页脚另建（无页码）
    for ref in body_sect.findall(qn("w:headerReference")): body_sect.remove(ref)
    for ref in body_sect.findall(qn("w:footerReference")): body_sect.remove(ref)
    pg = body_sect.find(qn("w:pgNumType"))
    if pg is not None: body_sect.remove(pg)

def fix_page_numbers(doc, profile, changes):
    pn = profile.get("page_numbers")
    if not pn: return
    for name, el in _needs_split(doc, profile):
        _split_before(doc, el); changes.append(f"分节：{name} 单独成节")
    roles, _ = _sect_role(doc, profile)
    for i, s in enumerate(doc.sections):
        role = roles[i]; sp = s._sectPr; pg = sp.find(qn("w:pgNumType"))
        if role == "none":
            for nm, f in _footers(s):
                if f.is_linked_to_previous:
                    if i == 0: continue
                    f.is_linked_to_previous = False
                for p in list(f.paragraphs):
                    if _has_field(p._p, "PAGE")[0]:
                        for c in list(p._p):
                            if c.tag != qn("w:pPr"): p._p.remove(c)
            continue
        if pg is None:
            pg = OxmlElement("w:pgNumType")
            anchor = next((c for c in sp if c.tag.split("}")[1] in ("cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange")), None)
            if anchor is not None: anchor.addprevious(pg)
            else: sp.append(pg)
        pg.set(qn("w:fmt"), pn["roman_fmt"] if role == "roman" else pn["arabic_fmt"])
        if not any(roles[j] == role for j in range(i)): pg.set(qn("w:start"), "1")
        elif pg.get(qn("w:start")) is not None: del pg.attrib[qn("w:start")]
        for nm, f in _footers(s)[:2]:
            if not _has_field(f._element, "PAGE")[0]:
                f.is_linked_to_previous = False
                el = f._element
                for p in list(el.findall(qn("w:p"))): el.remove(p)
                el.append(_page_field_par(doc, pn))
            if f.is_linked_to_previous: continue
            for p in f.paragraphs:
                if not _has_field(p._p, "PAGE")[0]: continue
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Pt(0)
                depth = 0
                for r in list(p._p.iter(qn("w:r"))):
                    fc = r.find(qn("w:fldChar"))
                    if fc is not None:
                        depth += {"begin": 1, "end": -1}.get(fc.get(qn("w:fldCharType")), 0); continue
                    if depth == 0 and r.find(qn("w:t")) is not None: r.getparent().remove(r)    # 去修饰
                for run in p.runs:
                    run.font.size = Pt(pn["size_pt"]); core._set_fonts(run._r, "宋体", pn["latin"])
    changes.append("页码：封面/声明无页码，摘要至目录罗马数字从 I 起，正文起阿拉伯数字从 1 起，Times New Roman 五号居中")

# ---------- 3b. 空白页（连续分页符） ----------
def _visible(el):
    if el.tag == qn("w:tbl"): return True
    if list(el.iter(qn("w:drawing"))) or list(el.iter(qn("w:pict"))) or list(el.iter(core.MATH + "oMath")): return True
    return bool("".join(t.text or "" for t in el.iter(qn("w:t"))).strip())

def redundant_breaks(doc):
    """返回会造成空白页的分页符：上一次换页之后还没有任何可见内容就又分页。
    换页来源：w:br type=page、分节符（段落内 sectPr，均为下一页分节）、pageBreakBefore。"""
    fresh = True; bad = []
    for el in doc.element.body:
        if el.tag == qn("w:tbl"): fresh = False; continue
        if el.tag != qn("w:p"): continue
        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None and ppr.find(qn("w:pageBreakBefore")).get(qn("w:val")) not in ("0", "false"):
            fresh = True
        for node in el.iter(qn("w:t"), qn("w:br"), qn("w:drawing"), qn("w:pict"), core.MATH + "oMath"):
            if node.tag == qn("w:br") and node.get(qn("w:type")) == "page":
                if fresh: bad.append(node)
                fresh = True
            elif node.tag == qn("w:t"):
                if (node.text or "").strip(): fresh = False
            else: fresh = False
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None: fresh = True
    return bad

def check_blank_pages(doc, profile, out):
    n = len(redundant_breaks(doc))
    if n: out.append(_iss("blank_page", "全文", "无空白页", f"{n} 处连续分页符", "每部分从新的一页开始（指南“学位论文…以上各部分独立为一部分，每部分从新的一页开始”），未要求从奇数页开始；连续分页会产生空白页",
                          (profile.get("page") or {}).get("clause")))

def fix_blank_pages(doc, profile, changes):
    bad = redundant_breaks(doc)
    for br in bad:
        r = br.getparent(); r.remove(br)
        if r.tag == qn("w:r") and not [c for c in r if c.tag != qn("w:rPr")]: r.getparent().remove(r)
    if bad: changes.append(f"空白页：删除 {len(bad)} 处多余分页符")

# ---------- 3c. 目录范围外的大纲级别（声明标题等不进目录） ----------
def _outline_level(doc, p):
    for ppr in [p._p.pPr] + [st.element.pPr for st in core._style_chain(p.style)]:
        ol = None if ppr is None else ppr.find(qn("w:outlineLvl"))
        if ol is not None: return int(ol.get(qn("w:val")))
    m = re.match(r"heading (\d)$", (p.style.name or "").lower())
    return int(m.group(1)) - 1 if m else 9

def toc_leaks(doc, profile):
    tc = profile.get("toc") or {}; ex = set(tc.get("exclude_regions", []))
    regs = {}; core.roles(doc, profile, regs)
    return [p for p in doc.paragraphs if regs.get(p._p) in ex and p.text.strip() and _outline_level(doc, p) < 9]

def check_toc_leaks(doc, profile, out):
    for p in toc_leaks(doc, profile):
        out.append(_iss("toc_outline_leak", f"「{p.text.strip()[:16]}」", "无大纲级别（不进目录）", f"样式 {p.style.name}，大纲级别 {_outline_level(doc, p)+1}",
                        "目录从第一章开始，目录之前的内容及目录本身不列入；原创性声明不属于正文（Word/WPS 更新目录时会把它列入）", (profile.get("toc") or {}).get("clause")))

def fix_toc_leaks(doc, profile, changes):
    from docx.enum.style import WD_STYLE_TYPE
    for p in toc_leaks(doc, profile):
        src = p.style; name = f"NoTOC {src.name}"   # 纯 ASCII 名：LibreOffice 的 STYLEREF 不认带括号的中文样式名
        try: st = doc.styles[name]
        except KeyError:
            el = copy.deepcopy(src.element); sid = re.sub(r"\W", "", src.style_id) + "NoToc"
            el.set(qn("w:styleId"), sid); el.find(qn("w:name")).set(qn("w:val"), name)
            for t in ("w:link", "w:next"):
                x = el.find(qn(t))
                if x is not None: el.remove(x)
            for x in list(el.iter(qn("w:qFormat"))): x.getparent().remove(x)
            src.element.addnext(el); st = doc.styles[name]
            ppr = st.element.get_or_add_pPr()
            for x in ppr.findall(qn("w:numPr")) + ppr.findall(qn("w:outlineLvl")): ppr.remove(x)
            ol = OxmlElement("w:outlineLvl"); ol.set(qn("w:val"), "9")
            anchor = next((c for c in ppr if c.tag.split("}")[1] in ("divId", "cnfStyle", "rPr", "sectPr", "pPrChange")), None)
            anchor.addprevious(ol) if anchor is not None else ppr.append(ol)
        p.style = st
        if p._p.pPr is not None:
            for x in p._p.pPr.findall(qn("w:outlineLvl")) + p._p.pPr.findall(qn("w:numPr")): p._p.pPr.remove(x)
        changes.append(f"目录：「{p.text.strip()[:16]}」改用 {name}（格式相同、无大纲级别）")

# ---------- 4. 封面 ----------
CN_DIGIT = "〇一二三四五六七八九"
def cn_year(y): return "".join(CN_DIGIT[int(c)] for c in y)
def cn_month(m):
    m = int(m); return ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"][m]

def _cover_tables(doc):
    out = []
    for tbl in doc.tables:
        first = tbl.rows[0].cells[0].text.strip()
        out.append((first, tbl))
    return out

def _cover_targets(doc, cv):
    """(段落, 规则名) 列表：题目单元格、作者导师信息单元格、日期段落。只看第 1 节（封面）。"""
    sm = dict(section_map(doc)); res = []
    for first, tbl in _cover_tables(doc):
        if sm.get(tbl._tbl.find(".//" + qn("w:p"))) != 0: continue
        if re.match(cv["title_table"], first):
            for row in tbl.rows:
                for ci, c in enumerate(row.cells):
                    if ci == 0: continue                       # 标签列（"题目："）规范未规定
                    for p in c.paragraphs:
                        if p.text.strip(): res.append((p, "title"))
        elif re.match(cv["info_table"], first):
            for row in tbl.rows:                       # 只管填写内容列；标签列（“姓    名：”）保持官方模板原样
                for ci, c in enumerate(row.cells):
                    if ci == 0: continue
                    for p in c.paragraphs:
                        if p.text.strip(): res.append((p, "info"))
    for p in doc.paragraphs:
        if sm.get(p._p) == 0 and re.search(cv["date_pattern"], p.text): res.append((p, "date"))
    return res

def check_cover(doc, profile, out):
    cv = profile.get("cover")
    if not cv: return
    for p, kind in _cover_targets(doc, cv):
        r = cv[kind]; add = _adder(out)
        core.check_para(doc, p, r, f"封面·{kind}「{p.text.strip()[:10]}」", add)
        if kind == "date" and re.search(r"\d", p.text):
            out.append(_iss("cover_date_digits", "封面·日期", "汉字年月（如 二〇一四年六月）", p.text.strip(), "日期不用阿拉伯数字", r.get("clause")))

def _grid_pitch(sect):
    g = sect._sectPr.find(qn("w:docGrid"))
    if g is None or g.get(qn("w:type")) not in ("lines", "linesAndChars"): return None
    return int(g.get(qn("w:linePitch") or 0)) / 20 or None

def _mark_size(doc, p):
    """段落高度所用字号：有文字取最大 run 字号，空段取段落标记字号。"""
    runs = [r for r in core.all_runs(p) if r.text.strip()]
    if runs: return max(core.eff_size(doc, p, r) or 12 for r in runs)
    rpr = p._p.pPr.find(qn("w:rPr")) if p._p.pPr is not None else None
    v = core._rpr_val(rpr, "sz")
    if v: return int(v) / 2
    for st in core._style_chain(p.style):
        v = core._rpr_val(st.element.rPr, "sz")
        if v: return int(v) / 2
    return core.eff_size(doc, p, type("R", (), {"_r": type("X", (), {"rPr": None})()})())

def _pin_conf(profile):
    lp = profile.get("layout_pin")
    if lp: return lp
    cv = profile.get("cover") or {}
    return {"regions": ["cover"], "clause": cv.get("layout_clause")} if cv.get("pin_line_heights") else None

def pinned_heights(doc, profile):
    """固定版式区域（封面、声明等）每段在 Word（宋体/黑体/仿宋，em 高 = 1.0 字号）下的行高：
    字号 × 多倍行距，向上取整到文档网格行距；“最小值”行距取 max(最小值, 一格)。
    返回 [(段落, 固定行距pt)]；含图片的段落、已是固定行距的段落不处理。"""
    lp = _pin_conf(profile)
    if not lp: return []
    want = set(lp["regions"]); regs = {}; core.roles(doc, profile, regs)
    import math
    res = []
    for el, si in section_map(doc):
        top = el
        while top.getparent() is not doc.element.body: top = top.getparent()
        reg = regs.get(el) or regs.get(top)
        if reg is None and si == 0: reg = "cover"
        if top.tag == qn("w:tbl") and si == 0: reg = "cover"
        if reg not in want: continue
        p = P(doc, el); pitch = _grid_pitch(doc.sections[si])
        if list(el.iter(qn("w:drawing"))) or list(el.iter(qn("w:pict"))): continue
        rule = core.eff_pf(p, "line_spacing_rule"); ls = core.eff_pf(p, "line_spacing")
        if rule == WD_LINE_SPACING.EXACTLY: continue
        s = _mark_size(doc, p)
        snap = p._p.pPr is None or p._p.pPr.find(qn("w:snapToGrid")) is None or p._p.pPr.find(qn("w:snapToGrid")).get(qn("w:val")) not in ("false", "0")
        grid = (lambda h: math.ceil(h / pitch - 1e-6) * pitch) if (pitch and snap) else (lambda h: h)
        if rule == WD_LINE_SPACING.AT_LEAST: h = max(ls.pt, grid(s))
        else: h = grid(s * (ls if isinstance(ls, float) else 1.0))
        res.append((p, round(h, 2)))
    return res

def cover_pinned_heights(doc, profile): return pinned_heights(doc, profile)

def check_cover_layout(doc, profile, out):
    lp = _pin_conf(profile)
    if not lp: return
    left = pinned_heights(doc, profile)
    if left:
        out.append(_iss("cover_line_auto", "/".join(lp["regions"]), "固定行距（与 Word+宋体 排版一致）", f"{len(left)} 段为多倍/单倍/最小值行距",
                        "缺少宋体/黑体/仿宋时，替代字体行高大于文档网格，每行被放大为两格，封面/声明溢出", lp.get("clause")))

def pin_cover(doc, profile, changes):
    items = pinned_heights(doc, profile)
    for p, h in items:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY; p.paragraph_format.line_spacing = Pt(h)
    if items: changes.append(f"封面/声明：{len(items)} 段行距按 Word 排版结果固定（与字体无关）")

# ---------- 签名栏与上文同页 ----------
def _sig_targets(doc, profile):
    """签名栏段落及其上方需要 keepNext 的段落：从上一个有文字的段落（含）到签名栏前一段。"""
    kt = profile.get("keep_together") or {}
    if not kt: return []
    rx = re.compile(kt["pattern"]); bs = re.compile(kt["block_start"]) if kt.get("block_start") else None
    regs = {}; core.roles(doc, profile, regs)
    paras = [p for p in doc.paragraphs if regs.get(p._p) in set(kt["regions"])]
    res = []; seen = set()
    def add(p, k):
        if (p._p, k) not in seen: seen.add((p._p, k)); res.append((p, k))
    for i, p in enumerate(paras):
        if not rx.match(p.text.strip()): continue
        # 往上找本块起点（如“学位论文使用授权说明”）；找不到就只连到上一段有文字的段落
        j0 = next((j for j in range(i - 1, -1, -1) if bs and bs.match(paras[j].text.strip())), None)
        if j0 is None or any(rx.match(paras[k].text.strip()) for k in range(j0 + 1, i) if not rx.match(paras[i].text.strip()) or False):
            j0 = i - 1
            while j0 > 0 and not paras[j0].text.strip(): j0 -= 1
        if j0 is not None and j0 >= 0:
            for k in range(j0, i):
                add(paras[k], "keep_next")
                if paras[k].text.strip(): add(paras[k], "keep_lines")
        if i + 1 < len(paras) and rx.match(paras[i + 1].text.strip()): add(p, "keep_next")
    return res

def check_keep_together(doc, profile, out):
    kt = profile.get("keep_together") or {}
    miss = [p for p, k in _sig_targets(doc, profile) if not getattr(p.paragraph_format, "keep_with_next" if k == "keep_next" else "keep_together")]
    if miss: out.append(_iss("sig_keep", "原创性声明", "签名栏与上文同页（与下段同页、段中不分页）", f"{len(miss)} 处未设置",
                             "签名栏被单独挤到下一页", kt.get("clause")))

def fix_keep_together(doc, profile, changes):
    t = _sig_targets(doc, profile)
    for p, k in t:
        if k == "keep_next": p.paragraph_format.keep_with_next = True
        else: p.paragraph_format.keep_together = True
    if t: changes.append(f"原创性声明：签名栏及其上方 {len(t)} 处设置“与下段同页/段中不分页”")

# ---------- 模板占位空段（只给建议） ----------
def check_empty_filler(doc, profile, out):
    regs = {}; core.roles(doc, profile, regs); run = 0; start = None
    for p in doc.paragraphs + [None]:
        empty = p is not None and not p.text.strip() and regs.get(p._p) in ("abstract_zh", "abstract_en", "body") \
            and not list(p._p.iter(qn("w:drawing"))) and not list(p._p.iter(qn("w:pict"))) and p._p.find(f".//{core.MATH}oMath") is None
        if empty:
            if run == 0: start = p
            run += 1
        else:
            if run >= 5:
                out.append(_iss("empty_filler", f"{regs.get(start._p)}·连续空段", "无大段空行", f"{run} 个连续空段落", "模板占位空行会把后文（如关键词）推到下一页，形成几乎空的页",
                                fixable=False, suggestion="属于模板示例的占位空行，写入真实内容时删除；工具不删除用户段落"))
            run = 0

def fix_cover(doc, profile, changes):
    cv = profile.get("cover")
    if not cv: return
    for p, kind in _cover_targets(doc, cv):
        if kind == "date":
            txt = p.text
            m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", txt)
            if m:
                new = txt.replace(m.group(0), f"{cn_year(m.group(1))}年{cn_month(m.group(2))}月")
                for run in p.runs[1:]: run._r.getparent().remove(run._r)
                if p.runs: p.runs[0].text = new
        apply_direct(p, cv[kind])
    changes.append("封面：题目一号黑体、信息三号仿宋、日期三号宋体汉字")

# ---------- 5. 固定格式与内容：版权声明、原创性声明 ----------
def _region_paras(doc, profile, name):
    regs = {}; core.roles(doc, profile, regs)
    return [p for p in doc.paragraphs if regs.get(p._p) == name]

def _norm(t): return re.sub(r"\s+", "", t)

def _fixed_sig(doc, paras):
    sig = []
    for p in paras:
        if not p.text.strip(): continue
        runs = [x for x in core.all_runs(p) if x.text.strip()]
        f = runs[0] if runs else None
        sig.append((_norm(p.text), core.eff_pf(p, "alignment"),
                    round(core.eff_size(doc, p, f), 1) if f else None, core.eff_font(doc, p, f, "eastAsia") if f else None))
    return sig

def _template_doc(profile, key):
    path = (profile.get("fixed_text") or {}).get(key, {}).get("template")
    import os
    if not path: return None
    path = os.path.join(profile.get("_base", "."), path)
    return Document(path) if os.path.exists(path) else None

def check_fixed_text(doc, profile, out):
    for key, cfg in (profile.get("fixed_text") or {}).items():
        tdoc = _template_doc(profile, key)
        paras = _region_paras(doc, profile, cfg["region"])
        if not paras:
            out.append(_iss("fixed_missing", key, "存在", "缺失", "固定内容缺失", cfg.get("clause"), fixable=tdoc is not None)); continue
        if tdoc is None:
            out.append(_iss("fixed_unverified", key, "与官方模板一致", "未提供官方模板，无法比对", "固定内容", cfg.get("clause"), fixable=False,
                            suggestion="在 profile 的 fixed_text.*.template 指向本地下载的官方模板")); continue
        want = _fixed_sig(tdoc, _region_paras(tdoc, profile, cfg["region"]))
        got = _fixed_sig(doc, paras)
        if [w[0] for w in want] != [g[0] for g in got]:
            out.append(_iss("fixed_text", key, "与官方模板文字一致", "文字被改动", "固定内容（全校统一）", cfg.get("clause")))
        elif want != got:
            out.append(_iss("fixed_format", key, "与官方模板格式一致", "格式被改动", "固定格式（全校统一）", cfg.get("clause")))

def fix_fixed_text(doc, profile, changes):
    for key, cfg in (profile.get("fixed_text") or {}).items():
        tdoc = _template_doc(profile, key)
        if tdoc is None: continue
        paras = _region_paras(doc, profile, cfg["region"])
        if not paras: continue
        if _fixed_sig(tdoc, _region_paras(tdoc, profile, cfg["region"])) == _fixed_sig(doc, paras): continue
        src = _region_paras(tdoc, profile, cfg["region"])
        anchor = paras[0]._p
        keep_sect = [p._p.pPr.find(qn("w:sectPr")) for p in paras if p._p.pPr is not None and p._p.pPr.find(qn("w:sectPr")) is not None]
        for p in src:
            el = copy.deepcopy(p._p)
            ppr = el.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None: ppr.remove(ppr.find(qn("w:sectPr")))
            for b in list(el.iter(qn("w:bookmarkStart"))) + list(el.iter(qn("w:bookmarkEnd"))): b.getparent().remove(b)
            anchor.addprevious(el)
        last_new = anchor.getprevious()
        for p in paras: p._p.getparent().remove(p._p)
        for sp in keep_sect:                      # 原有分节符保留在区域末尾
            ppr = last_new.find(qn("w:pPr"))
            if ppr is None: ppr = OxmlElement("w:pPr"); last_new.insert(0, ppr)
            ppr.append(sp)
        changes.append(f"{key}：按官方模板恢复固定文字与格式")

# ---------- 6. 摘要内容 ----------
FW2HW = {"，": ", ", "。": ". ", "：": ": ", "；": "; ", "（": " (", "）": ") ", "！": "! ", "？": "? ", "“": '"', "”": '"', "‘": "'", "’": "'"}

def _kw_split(text, prefix_re):
    body = re.sub(prefix_re, "", text.strip())
    return [k.strip() for k in re.split(r"[，,；;、]", body) if k.strip()]

def check_abstract(doc, profile, out):
    ab = profile.get("abstract")
    if not ab: return
    rl = core.roles(doc, profile); cl = ab.get("clause")
    zh = [p for p, r in rl if r == "abs_zh_kw"]; en = [p for p, r in rl if r == "en_kw"]
    if not zh: out.append(_iss("abs_zh_kw_missing", "中文摘要", "关键词：……", "缺失", "关键词放摘要页最下方，从新的一行撰写", cl, fixable=False, suggestion="在中文摘要末尾另起一行写 3~5 个关键词"))
    for p in zh:
        ks = _kw_split(p.text, ab["zh_kw_prefix_re"])
        if not (ab["zh_kw_min"] <= len(ks) <= ab["zh_kw_max"]):
            out.append(_iss("abs_zh_kw_count", "中文关键词", f"{ab['zh_kw_min']}~{ab['zh_kw_max']} 个", len(ks), "关键词数量", cl, fixable=False,
                            suggestion="关键词数量属于内容决定，工具不自动增删"))
        body = re.sub(ab["zh_kw_prefix_re"], "", p.text.strip())
        if re.search(r"[,；;、]", body) or not p.text.strip().startswith(ab["zh_kw_prefix"]):
            out.append(_iss("abs_zh_kw_sep", "中文关键词", f"前缀「{ab['zh_kw_prefix']}」，以「，」分隔", p.text.strip()[:30], "关键词分隔符", cl))
    if not en: out.append(_iss("abs_en_kw_missing", "英文摘要", "KEY WORDS: ……", "缺失", "关键词从新的一行写起", cl, fixable=False, suggestion="在英文摘要末尾另起一行写 KEY WORDS"))
    for p in en:
        t = p.text.strip(); ks = _kw_split(t, ab["en_kw_prefix_re"])
        bad = not t.startswith(ab["en_kw_prefix"]) or re.search(r"[，；、;]", t) or any(k[:1] != k[:1].upper() for k in ks)
        if bad: out.append(_iss("abs_en_kw_format", "英文关键词", "KEY WORDS: Aa, Bb（大写、首字母大写、半角逗号）", t[:40], "KEY WORDS 格式", cl))
    for p, r in rl:
        if r in ("en_title", "en_meta", "en_body", "en_kw") and re.search("[" + "".join(FW2HW) + "]", p.text):
            out.append(_iss("abs_en_fullwidth", f"英文摘要「{p.text.strip()[:12]}」", "英文标点", "含全角标点", "标点符号用英文标点", cl))
    regs = {}; core.roles(doc, profile, regs)
    for p in doc.paragraphs:
        if regs.get(p._p) == "abstract_zh" and (list(p._p.iter(qn("w:drawing"))) or list(p._p.iter(qn("w:pict")))):
            out.append(_iss("abs_has_figure", "中文摘要", "无图表", "含图片", "摘要不能出现图片、表格", cl, fixable=False, suggestion="把图移到正文"))
    body = doc.element.body
    for tbl in body.iter(qn("w:tbl")):
        prev = tbl.getprevious()
        while prev is not None and prev.tag != qn("w:p"): prev = prev.getprevious()
        if prev is not None and regs.get(prev) == "abstract_zh":
            out.append(_iss("abs_has_table", "中文摘要", "无图表", "含表格", "摘要不能出现图片、表格", cl, fixable=False, suggestion="把表移到正文"))

def _set_text(p, text):
    """替换段落文字：写入第一个含文字的 run，清空其他 run 的文字；保留分页符、域等无文字 run。"""
    runs = [r for r in core.all_runs(p) if r._r.find(qn("w:t")) is not None]
    if not runs: p.add_run(text); return
    for r in runs[1:]:
        for t in r._r.findall(qn("w:t")): r._r.remove(t)
    t0 = runs[0]._r.findall(qn("w:t"))
    for t in t0[1:]: runs[0]._r.remove(t)
    t0[0].text = text; t0[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

def fix_abstract(doc, profile, changes):
    ab = profile.get("abstract")
    if not ab: return
    rl = core.roles(doc, profile)
    for p, r in rl:
        if r == "abs_zh_kw":
            ks = _kw_split(p.text, ab["zh_kw_prefix_re"]); _set_text(p, ab["zh_kw_prefix"] + "，".join(ks))
        elif r == "en_kw":
            ks = _kw_split(p.text, ab["en_kw_prefix_re"]); ks = [k[:1].upper() + k[1:] for k in ks]
            _set_text(p, ab["en_kw_prefix"] + ", ".join(ks))
        if r in ("en_title", "en_meta", "en_body", "en_kw"):
            for run in core.all_runs(p):
                t = run.text
                if not any(a in t for a in FW2HW): continue
                for a, b in FW2HW.items(): t = t.replace(a, b)
                t = re.sub(r" {2,}", " ", t).replace(" )", ")").replace("( ", "(")
                if t != run.text: run.text = t
    changes.append("摘要：关键词前缀与分隔符、KEY WORDS 大小写、英文摘要半角标点")

# ---------- 7. 目录 ----------
def _toc_fields(doc):
    res = []
    for it in doc.element.body.iter(qn("w:instrText")):
        if re.match(r"\s*TOC\b", it.text or ""): res.append(it)
    return res

def toc_entries(doc):
    """目录项段落（含 w:sdt 目录内容控件内的段落，doc.paragraphs 看不到它们）。"""
    ps = [P(doc, el) for el in doc.element.body.iter(qn("w:p"))]
    return [(int(re.search(r"(\d)", p.style.name).group(1)), p.text) for p in ps if core.is_toc_para(p)]

def expected_toc(doc, profile):
    """按标题计算应有的目录项（含自动编号的显示文本）。"""
    res = []; counters = [0, 0, 0, 0]; levels = {"heading1": 1, "heading2": 2, "heading3": 3}
    hnum = core.heading_num_ids(doc)
    for p, role in core.roles(doc, profile):
        if role not in levels: continue
        lv = levels[role]; nid, _ = core.para_numpr(p)
        txt = p.text.strip()
        if nid in hnum:
            counters[lv - 1] += 1
            for k in range(lv, 4): counters[k] = 0
            nums = profile["numbering"]["levels"][lv - 1]["display"]
            pre = nums.format(*([cn_num(counters[0])] if lv == 1 else counters[:lv]))
            txt = pre + txt
        res.append((lv, txt))
    return res

def cn_num(n):
    d = "零一二三四五六七八九"
    if n < 10: return d[n]
    if n < 20: return "十" + (d[n % 10] if n % 10 else "")
    return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")

def _norm_entry(t): return re.sub(r"\s+", "", re.sub(r"[\t ]*[\dIVXivx]+\s*$", "", t))

def check_toc(doc, profile, out):
    tc = profile.get("toc")
    if not tc: return
    cl = tc.get("clause"); flds = _toc_fields(doc)
    if not flds:
        out.append(_iss("toc_missing", "目录", "TOC 域", "无", "目录", cl)); return
    instr = flds[0].text
    m = re.search(r'\\o\s+"(\d)-(\d)"', instr)
    if not m or m.group(2) != str(tc["levels"]):
        out.append(_iss("toc_levels", "目录", f'\\o "1-{tc["levels"]}"', instr.strip(), "目录列到三级标题", cl))
    if "\\u" in instr:
        out.append(_iss("toc_switch_u", "目录", "不含 \\u", instr.strip(), "\\u 会把大纲级别段落也列入，可能把前置部分带进目录", cl))
    for role, r in profile["styles"].items():
        if r.get("outline_body"):
            for p, ro in core.roles(doc, profile):
                if ro == role and (p.style.name.lower().startswith("heading") or core.para_numpr(p)[0] not in (None, "0") and False):
                    out.append(_iss("toc_front_included", f"「{p.text.strip()}」", "不进入目录", f"样式 {p.style.name}（会进目录）", "目录从第一章开始，目录之前的内容及目录本身不列入", cl))
    for name, r in tc["styles"].items():
        st = core.find_style(doc, name)          # 名式不敏感：'TOC 1' 与 'toc 1' 是同一样式，不许误报缺失
        if st is None:
            out.append(_iss("toc_style_missing", name, "存在", "缺失", "目录样式", cl)); continue
        dummy = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:pStyle w:val="{st.style_id}"/></w:pPr><w:r><w:t>目录项Ab1</w:t></w:r></w:p>')
        core.check_para(doc, P(doc, dummy), r, f"目录样式 {name}", _adder(out))
    got = [(lv, _norm_entry(t)) for lv, t in toc_entries(doc)]
    want = [(lv, _norm_entry(t)) for lv, t in expected_toc(doc, profile)]
    if got != want:
        out.append(_iss("toc_stale", "目录", f"{len(want)} 项，与标题一致", f"{len(got)} 项", "目录内容与标题不一致（需更新目录）", cl))

TOC_INSTR = ' TOC \\o "1-{n}" \\h \\z '

def fix_toc(doc, profile, changes):
    tc = profile.get("toc")
    if not tc: return
    for role, r in profile["styles"].items():          # 前置标题改用不进目录的样式
        if r.get("outline_body"):
            core._apply_style_def(doc, r, changes)
    for name, r in tc["styles"].items():
        st = core._ensure_style(doc, name, changes); rr = dict(r, style=name); core._apply_style_def(doc, rr, changes)
        st.paragraph_format.first_line_indent = Pt(0)
        st.paragraph_format.left_indent = Pt(r.get("left_indent_pt", 0))
        ts = st.paragraph_format.tab_stops
        for t in list(ts): ts._pPr.tabs.remove(t._tab) if hasattr(ts, "_pPr") else None
        sec = doc.sections[-1]
        width = sec.page_width - sec.left_margin - sec.right_margin
        from docx.enum.text import WD_TAB_LEADER
        ts.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    flds = _toc_fields(doc)
    if flds:
        flds[0].text = TOC_INSTR.format(n=tc["levels"])
    else:
        regs = {}; core.roles(doc, profile, regs)
        head = next((p for p in doc.paragraphs if regs.get(p._p) == "toc" and re.match(tc["heading_pattern"], p.text.strip())), None)
        if head is not None:
            p = parse_xml(f'<w:p xmlns:w="{W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{TOC_INSTR.format(n=tc["levels"])}</w:instrText></w:r>'
                          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
            head._p.addnext(p)
    rebuild_toc_result(doc, profile, None)
    st = doc.settings.element
    if profile.get("_toc_mark_dirty", True) and st.find(qn("w:updateFields")) is None:
        # 页码未经排版计算时，让 Word 打开时提示更新域；LibreOffice 也会据此重建目录（其重建结果不含自动编号，见风险）
        uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); _settings_insert(st, uf)
    changes.append("目录：TOC 域 \\o 1-3、目录样式、前置标题不进目录、目录项按标题重建"
                   + ("（页码未计算：已设置打开时更新域）" if profile.get("_toc_mark_dirty", True) else ""))

def rebuild_toc_result(doc, profile, pages):
    """用当前标题重建 TOC 域的结果段落；pages 为 {目录项文本: 页码}（来自 LibreOffice 排版），缺省则页码留空。"""
    tc = profile["toc"]; flds = _toc_fields(doc)
    if not flds: return
    begin_p = _enclosing_p(flds[0])
    # 找到域结束所在段落（目录可能位于 w:sdt 内容控件中，按其所在容器遍历）
    body = list(begin_p.getparent()); i0 = body.index(begin_p); end_p = None; depth = 0
    for el in body[i0:]:
        for fc in el.iter(qn("w:fldChar")):
            t = fc.get(qn("w:fldCharType"))
            if t == "begin": depth += 1
            elif t == "end":
                depth -= 1
                if depth == 0: end_p = el; break
        if end_p is not None: break
    if end_p is None: return
    # 删除 begin_p 与 end_p 之间（含两端里的结果部分）的所有段落，重建
    idx_b, idx_e = body.index(begin_p), body.index(end_p)
    for el in body[idx_b + 1: idx_e + 1]: el.getparent().remove(el)
    # begin_p：保留 begin/instr/separate，删掉其后内容
    seen_sep = False
    for c in list(begin_p):
        if c.tag == qn("w:pPr"): continue
        if seen_sep: begin_p.remove(c); continue
        for fc in c.iter(qn("w:fldChar")):
            if fc.get(qn("w:fldCharType")) == "separate": seen_sep = True
        if c.find(qn("w:fldChar")) is not None and c.find(qn("w:fldChar")).get(qn("w:fldCharType")) == "end": begin_p.remove(c)
    entries = expected_toc(doc, profile)
    style_ids = {lv: core.require_style(doc, f"toc {lv}").style_id for lv in (1, 2, 3)}
    anchor = begin_p
    if not entries:
        r = OxmlElement("w:r"); fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), "end"); r.append(fc); begin_p.append(r); return
    for k, (lv, txt) in enumerate(entries):
        page = "" if pages is None else str(pages.get(_norm_entry(txt), ""))
        if k == 0:
            p = begin_p
            ppr = p.find(qn("w:pPr"))
            if ppr is None: ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
            ps = ppr.find(qn("w:pStyle"))
            if ps is None: ps = OxmlElement("w:pStyle"); ppr.insert(0, ps)
            ps.set(qn("w:val"), style_ids[lv])
            for c in list(ppr):
                if c.tag == qn("w:rPr"): ppr.remove(c)
            for r in p.findall(qn("w:r")):          # 域起止 run 不带粗体等直接格式
                rp = r.find(qn("w:rPr"))
                if rp is not None: r.remove(rp)
        else:
            p = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:pStyle w:val="{style_ids[lv]}"/></w:pPr></w:p>')
            anchor.addnext(p); anchor = p
        r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); t.text = txt; r.append(t); p.append(r)
        r = OxmlElement("w:r"); r.append(OxmlElement("w:tab")); p.append(r)
        r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = page; r.append(t); p.append(r)
    r = OxmlElement("w:r"); fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), "end"); r.append(fc); anchor.append(r)

# ---------- 8. 脚注 ----------
def _footnote_part(doc):
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/footnotes"): return rel.target_part
    return None

def _fn_root(part):
    from .revisions import _xml
    return _xml(part)

def _flush_part(part):
    if hasattr(part, "_guard_el"):
        from lxml import etree
        part._blob = etree.tostring(part._guard_el, xml_declaration=True, encoding="UTF-8", standalone=True)

def _real_footnotes(root):
    return [f for f in root.findall(qn("w:footnote")) if f.get(qn("w:type")) in (None, "normal")]

def check_footnotes(doc, profile, out):
    fn = profile.get("footnotes")
    if not fn: return
    cl = fn.get("clause"); part = _footnote_part(doc)
    refs = list(doc.element.body.iter(qn("w:footnoteReference")))
    if part is None or not refs: return
    root = _fn_root(part)
    for f in _real_footnotes(root):
        fid = f.get(qn("w:id"))
        for pel in f.findall(qn("w:p")):
            p = P(doc, pel); where = f"脚注{fid}"
            core.check_para(doc, p, fn, where, _adder(out))
            ind = core.indent_first_pt(p, fn["size_pt"])
            if abs(-ind - fn["hanging_chars"] * fn["size_pt"]) > 0.6:
                out.append(_iss("fn_hanging", where, f"悬挂 {fn['hanging_chars']} 字符", f"{-ind:.2f}pt", "悬挂缩进", cl))
            ref = next((r for r in pel.iter(qn("w:r")) if r.find(qn("w:footnoteRef")) is not None), None)
            if ref is not None:
                run = core.Document  # noqa
                from docx.text.run import Run
                rr = Run(ref, p)
                if _is_super(doc, p, rr): out.append(_iss("fn_mark_super", where, "正文（非上标）", "上标", "脚注处序号字体是正文不是上标", cl))
                nxt = ref.getnext(); txt = ""
                while nxt is not None and not txt:
                    txt = "".join(t.text or "" for t in nxt.iter(qn("w:t"))); nxt = nxt.getnext()
                if not txt.startswith((" ", "\u2002")) or txt.startswith("  "):
                    out.append(_iss("fn_gap", where, "序号后空半个汉字符", repr(txt[:3]), "序号与脚注内容之间", cl))
    for i, ref in enumerate(refs):
        from docx.text.run import Run
        rr = Run(ref.getparent(), P(doc, _enclosing_p(ref)))
        if not _is_super(doc, rr._parent, rr):
            out.append(_iss("fn_ref_not_super", f"正文脚注引用{i+1}", "上标", "非上标", "正文中脚注序号以上标标示", cl))
    for i, s in enumerate(doc.sections):
        fp = s._sectPr.find(qn("w:footnotePr"))
        fmt = None if fp is None or fp.find(qn("w:numFmt")) is None else fp.find(qn("w:numFmt")).get(qn("w:val"))
        rst = None if fp is None or fp.find(qn("w:numRestart")) is None else fp.find(qn("w:numRestart")).get(qn("w:val"))
        if i >= profile["header"].get("from_section", 0) and (fmt != fn["num_fmt"] or rst != fn["num_restart"]):
            out.append(_iss("fn_numbering", f"节{i+1}", f"{fn['num_fmt']}/{fn['num_restart']}", f"{fmt}/{rst}", "脚注序号①②③、按页编排", cl))

def _enclosing_p(el):
    while el is not None and el.tag != qn("w:p"): el = el.getparent()
    return el

def _is_super(doc, p, run):
    rpr = run._r.rPr
    va = None if rpr is None else rpr.find(qn("w:vertAlign"))
    if va is not None: return va.get(qn("w:val")) == "superscript"
    rs = None if rpr is None else rpr.find(qn("w:rStyle"))
    if rs is not None:
        st = next((s for s in doc.styles if s.style_id == rs.get(qn("w:val"))), None)
        while st is not None:
            e = None if st.element.rPr is None else st.element.rPr.find(qn("w:vertAlign"))
            if e is not None: return e.get(qn("w:val")) == "superscript"
            st = st.base_style
    return False

def _set_vert(run_el, val):
    rpr = run_el.find(qn("w:rPr"))
    if rpr is None: rpr = OxmlElement("w:rPr"); run_el.insert(0, rpr)
    va = rpr.find(qn("w:vertAlign"))
    if va is None:
        va = OxmlElement("w:vertAlign")
        anchor = next((c for c in rpr if c.tag.split("}")[1] in ("rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")), None)
        if anchor is not None: anchor.addprevious(va)
        else: rpr.append(va)
    va.set(qn("w:val"), val)

def fix_footnotes(doc, profile, changes):
    fn = profile.get("footnotes")
    if not fn: return
    part = _footnote_part(doc)
    if part is None: return
    st = core._ensure_style(doc, fn["style"], changes); core._apply_style_def(doc, dict(fn, style=fn["style"]), changes)
    pf = st.paragraph_format; h = Pt(fn["hanging_chars"] * fn["size_pt"])
    pf.left_indent = h; pf.first_line_indent = -h
    root = _fn_root(part); sid = st.style_id
    for f in _real_footnotes(root):
        for pel in f.findall(qn("w:p")):
            p = P(doc, pel)
            ppr = pel.find(qn("w:pPr"))
            if ppr is None: ppr = OxmlElement("w:pPr"); pel.insert(0, ppr)
            ps = ppr.find(qn("w:pStyle"))
            if ps is None: ps = OxmlElement("w:pStyle"); ppr.insert(0, ps)
            ps.set(qn("w:val"), sid)
            core._strip_para_direct(p)
            for r in pel.iter(qn("w:r")):
                if r.find(qn("w:footnoteRef")) is not None:
                    _set_vert(r, "baseline")
                    rpr = r.find(qn("w:rPr")); keep = {qn("w:rStyle"), qn("w:vertAlign")}
                    for c in list(rpr):
                        if c.tag not in keep: rpr.remove(c)
                    # 序号后半个汉字符：规范化为一个半角空格
                    nxt = r.getnext()
                    while nxt is not None and nxt.tag == qn("w:r") and not "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip():
                        nn = nxt.getnext(); nxt.getparent().remove(nxt); nxt = nn
                    if nxt is not None:
                        t = next(nxt.iter(qn("w:t")), None)
                        if t is not None: t.text = " " + (t.text or "").lstrip(); t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                else:
                    from docx.text.run import Run
                    core._strip_run_direct(Run(r, p), ["bold", "italic", "superscript", "subscript"])
    _flush_part(part)
    for ref in doc.element.body.iter(qn("w:footnoteReference")): _set_vert(ref.getparent(), "superscript")
    for i, s in enumerate(doc.sections):
        if i < profile["header"].get("from_section", 0): continue
        sp = s._sectPr; fp = sp.find(qn("w:footnotePr"))
        if fp is None:
            fp = OxmlElement("w:footnotePr")
            anchor = next((c for c in sp if c.tag.split("}")[1] not in ("headerReference", "footerReference")), None)
            anchor.addprevious(fp) if anchor is not None else sp.append(fp)
        for c in list(fp): fp.remove(c)
        nf = OxmlElement("w:numFmt"); nf.set(qn("w:val"), fn["num_fmt"]); fp.append(nf)
        nr = OxmlElement("w:numRestart"); nr.set(qn("w:val"), fn["num_restart"]); fp.append(nr)
    changes.append("脚注：小五号宋体/Times New Roman、单倍行距、悬挂缩进 1.5 字符、序号①②按页编排、注内序号非上标、正文引用上标")

# ---------- 9. 表格、题注位置、表注 ----------
def _body_tables(doc, profile):
    regs = {}; core.roles(doc, profile, regs); res = []
    for tbl in doc.element.body.iter(qn("w:tbl")):
        prev = tbl.getprevious()
        while prev is not None and prev.tag != qn("w:p"): prev = prev.getprevious()
        if prev is not None and regs.get(prev) == "body": res.append(tbl)
        elif prev is not None and regs.get(prev) is None:
            pp = prev
            while pp is not None and (pp.tag != qn("w:p") or regs.get(pp) is None): pp = pp.getprevious()
            if pp is not None and regs.get(pp) == "body": res.append(tbl)
    return res

def _ptext(el): return "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()

def _cap_after(tbl, rx, n=2):
    """表后 n 个非空段内的表题（中间可能隔着表注）。"""
    x = tbl
    for _ in range(n):
        x = _neighbor(x, +1)
        if x is None or x.tag != qn("w:p"): return None
        if rx.match(_ptext(x)): return x
    return None

def _neighbor(el, step):
    n = el.getnext() if step > 0 else el.getprevious()
    while n is not None and n.tag == qn("w:p") and not _ptext(n) and not list(n.iter(qn("w:drawing"))) and not list(n.iter(qn("w:pict"))):
        n = n.getnext() if step > 0 else n.getprevious()
    return n

def check_tables(doc, profile, out):
    tb = profile.get("tables")
    if not tb: return
    cl = tb.get("clause"); tcap = re.compile(profile["styles"]["table_caption"]["pattern"]); fcap = re.compile(profile["styles"]["figure_caption"]["pattern"])
    for k, tbl in enumerate(_body_tables(doc, profile)):
        before, after = _neighbor(tbl, -1), _neighbor(tbl, +1)
        if not (before is not None and before.tag == qn("w:p") and tcap.match(_ptext(before))):
            if _cap_after(tbl, tcap) is not None:
                out.append(_iss("table_caption_below", f"表格{k+1}", "表题在表上方", "在下方", "表序与表名置于表的上方", cl))
            else:
                out.append(_iss("table_caption_missing", f"表格{k+1}", "有表题", "无", "表应有表序与表名", cl, fixable=False, suggestion="在表上方补写「表 x.y　表名」"))
        for p_el in tbl.iter(qn("w:p")):
            p = P(doc, p_el)
            if p.text.strip(): core.check_para(doc, p, tb["cell"], f"表格{k+1}「{p.text.strip()[:8]}」", _adder(out))
    for p in doc.paragraphs:
        if list(p._p.iter(qn("w:drawing"))) or list(p._p.iter(qn("w:pict"))):
            regs = {}
            nxt, prv = _neighbor(p._p, +1), _neighbor(p._p, -1)
            if not (nxt is not None and nxt.tag == qn("w:p") and fcap.match(_ptext(nxt))):
                if prv is not None and prv.tag == qn("w:p") and fcap.match(_ptext(prv)):
                    out.append(_iss("figure_caption_above", f"图「{_ptext(prv)[:10]}」", "图题在图下方", "在上方", "图序与图名置于图的下方", cl))
    for role in ("figure_caption", "table_caption"):
        for p, r in core.roles(doc, profile):
            if r != role: continue
            m = re.match(r"^((续)?[图表]\s*[A-Z\d]+\.\d+)(\s*)(.*)$", p.text.strip())
            if m and m.group(3) != tb["caption_gap"]:
                out.append(_iss("caption_gap", f"「{p.text.strip()[:12]}」", "序号与名称之间空一个汉字符", repr(m.group(3)), "题注间隔", cl))

def fix_tables(doc, profile, changes):
    tb = profile.get("tables")
    if not tb: return
    tcap = re.compile(profile["styles"]["table_caption"]["pattern"]); fcap = re.compile(profile["styles"]["figure_caption"]["pattern"])
    st = core._ensure_style(doc, tb["cell"]["style"], changes); core._apply_style_def(doc, dict(tb["cell"], style=tb["cell"]["style"]), changes)
    st.paragraph_format.first_line_indent = Pt(0)
    for tbl in _body_tables(doc, profile):
        before, after = _neighbor(tbl, -1), _neighbor(tbl, +1)
        below = _cap_after(tbl, tcap)
        if not (before is not None and before.tag == qn("w:p") and tcap.match(_ptext(before))) and below is not None:
            tbl.addprevious(below); changes.append(f"表题「{_ptext(below)[:10]}」移到表上方")
        for p_el in tbl.iter(qn("w:p")):
            p = P(doc, p_el); p.style = st
            core._strip_para_direct(p, ("jc",))                              # 单元格对齐属于表格设计，保留
            for run in core.all_runs(p): core._strip_run_direct(run, ["bold", "italic", "superscript", "subscript"])
    for p in list(doc.paragraphs):
        if list(p._p.iter(qn("w:drawing"))) or list(p._p.iter(qn("w:pict"))):
            nxt, prv = _neighbor(p._p, +1), _neighbor(p._p, -1)
            if not (nxt is not None and nxt.tag == qn("w:p") and fcap.match(_ptext(nxt))) and \
               prv is not None and prv.tag == qn("w:p") and fcap.match(_ptext(prv)):
                p._p.addnext(prv); changes.append(f"图题「{_ptext(prv)[:10]}」移到图下方")
    for role in ("figure_caption", "table_caption"):
        for p, r in core.roles(doc, profile):
            if r != role: continue
            m = re.match(r"^((续)?[图表]\s*[A-Z\d]+\.\d+)(\s*)(.*)$", p.text.strip())
            if m and m.group(3) != tb["caption_gap"]:
                _replace_in_runs(p, m.group(1) + m.group(3), m.group(1) + tb["caption_gap"])
    changes.append("表格：单元格 11pt 宋体单倍行距段前后 3 磅；题注位置与间隔")

def _replace_in_runs(p, old, new):  # noqa: E302
    """在段落内替换文本（可跨 run），保留第一个 run 的格式。"""
    runs = [r for r in core.all_runs(p)]
    full = "".join(r.text for r in runs); i = full.find(old)
    if i < 0: return False
    pos = 0; done = False
    for r in runs:
        t = r.text; a, b = pos, pos + len(t); pos = b
        if done or b <= i or a >= i + len(old): continue
        s0, s1 = max(i, a) - a, min(i + len(old), b) - a
        if a <= i < b: r.text = t[:s0] + new + t[s1:]
        else: r.text = t[:s0] + t[s1:]
    return True

# ---------- 10. 公式 ----------
def _chapter_of(doc, profile):
    """每段所在章号（自动编号或手写"第X章"；附录用字母）。"""
    res = {}; ch = 0; appx = None; hnum = core.heading_num_ids(doc)
    for p, role in core.roles(doc, profile):
        if role == "heading1":
            t = p.text.strip(); m = re.match(r"附录\s*([A-Z])", t)
            if m: appx = m.group(1)
            elif re.match(r"^(参考文献|致谢)", t): pass
            else: ch += 1; appx = None
        res[p._p] = appx or str(ch)
    return res

def _eq_items(doc, profile):
    chs = _chapter_of(doc, profile); items = []
    for p, role in core.roles(doc, profile):
        if role == "equation":
            m = core.EQ_NUM.search(p.text.strip())
            items.append((p, chs.get(p._p), m))
    return items

def _expected_eq_numbers(items):
    cnt = {}; res = []
    for p, ch, m in items:
        cnt[ch] = cnt.get(ch, 0) + 1
        sep = "-" if ch and ch.isalpha() else "."
        res.append(f"{ch}{sep}{cnt[ch]}")
    return res

def check_equations(doc, profile, out):
    eq = profile.get("equations")
    if not eq: return
    cl = eq.get("clause"); items = _eq_items(doc, profile)
    sec = doc.sections[-1]; width = (sec.page_width - sec.left_margin - sec.right_margin)
    for (p, ch, m), want in zip(items, _expected_eq_numbers(items)):
        cur = f"{m.group(1)}{'-' if m.group(1).isalpha() else '.'}{m.group(2)}"
        if cur != want: out.append(_iss("eq_number", f"式({cur})", f"({want})", f"({cur})", "表达式按章编号", cl))
        txt = p.text.rstrip()
        if not re.search(r"\t\s*[（(][^)）]*[)）]$", txt):
            out.append(_iss("eq_number_pos", f"式({cur})", "序号用制表符置于行末", "未用右对齐制表位", "序号置于表达式右边行末", cl))
        tabs = [(t.alignment, t.position) for t in p.paragraph_format.tab_stops]
        if not any(a == WD_TAB_ALIGNMENT.RIGHT and abs(pos - width) < Twips(30) for a, pos in tabs):
            out.append(_iss("eq_tab_right", f"式({cur})", "右对齐制表位在行末", str(tabs), "序号右对齐", cl))
        if eq["mode"] == "center" and not any(a == WD_TAB_ALIGNMENT.CENTER and abs(pos - width // 2) < Twips(30) for a, pos in tabs):
            out.append(_iss("eq_tab_center", f"式({cur})", "居中制表位", str(tabs), "表达式居中书写", cl))
        if m.group(0).strip()[0] != eq["paren"][0]:
            out.append(_iss("eq_paren", f"式({cur})", eq["paren"], m.group(0).strip()[0], "序号括号", cl))
    # 正文中引用的式号必须存在
    nums = {f"{m.group(1)}{'-' if m.group(1).isalpha() else '.'}{m.group(2)}" for _, _, m in items}
    for p, role in core.roles(doc, profile):
        if role in ("body",):
            for r in re.finditer(r"式\s*[（(]\s*([A-Z\d]+[.\-]\d+)\s*[)）]", p.text):
                if r.group(1) not in nums:
                    out.append(_iss("eq_ref_dangling", f"「{p.text.strip()[:12]}」", "引用的式号存在", r.group(1), "正文引用式号", cl, fixable=False, suggestion="确认引用的是哪个公式"))

def fix_equations(doc, profile, changes):
    eq = profile.get("equations")
    if not eq: return
    items = _eq_items(doc, profile)
    sec = doc.sections[-1]; width = (sec.page_width - sec.left_margin - sec.right_margin)
    mapping = {}
    for (p, ch, m), want in zip(items, _expected_eq_numbers(items)):
        cur = f"{m.group(1)}{'-' if m.group(1).isalpha() else '.'}{m.group(2)}"
        if cur != want: mapping[cur] = want
        # 重写结构：[\t] 公式 [\t](x.y)
        runs = [r for r in p._p if r.tag in (qn("w:r"), MATHT("oMath"), MATHT("oMathPara"))]
        # 删除旧编号文字与多余制表符
        for r in list(p._p.iter(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is not None and t.text and (core.EQ_NUM.search(t.text) or re.fullmatch(r"\s*[（(]?\s*[A-Z\d]*[.\-]?\d*\s*[)）]?\s*", t.text)):
                if core.EQ_NUM.search(t.text) or t.text.strip() in ("", "(", ")", "（", "）"): r.getparent().remove(r)
            elif r.find(qn("w:tab")) is not None and r.getparent() is p._p: r.getparent().remove(r)
        first = next((c for c in p._p if c.tag != qn("w:pPr")), None)
        if eq["mode"] == "center" and first is not None:
            tr = OxmlElement("w:r"); tr.append(OxmlElement("w:tab")); first.addprevious(tr)
        tr = OxmlElement("w:r"); tr.append(OxmlElement("w:tab")); p._p.append(tr)
        nr = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = f"{eq['paren'][0]}{want}{eq['paren'][1]}"; nr.append(t); p._p.append(nr)
        ts = p.paragraph_format.tab_stops
        for tstop in list(ts): ts._pPr.tabs.remove(tstop._tab)
        if eq["mode"] == "center": ts.add_tab_stop(width // 2, WD_TAB_ALIGNMENT.CENTER)
        ts.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    if mapping:                               # 同步正文中的"式（x.y）"引用；用占位符避免链式替换
        for p, role in core.roles(doc, profile):
            if role == "equation": continue
            for k, (old, new) in enumerate(mapping.items()):
                for pat in (f"式（{old}）", f"式({old})", f"式 ({old})"):
                    while _replace_in_runs(p, pat, f"式（\ue010{k}\ue011）"): pass
            for k, (old, new) in enumerate(mapping.items()):
                while _replace_in_runs(p, f"\ue010{k}\ue011", new): pass
        changes.append(f"公式：按章重新编号 {mapping}，并同步正文引用")
    changes.append("公式：居中制表位 + 行末右对齐序号，单倍行距段前后 6 磅")

def MATHT(n): return core.MATH + n

# ---------- 11. 标题自动编号 ----------
def _numbering_part(doc):
    return doc.part.numbering_part

def _guard_num_id(doc):
    ids = core.heading_num_ids(doc)
    return sorted(ids)[0] if ids else None

def check_numbering(doc, profile, out):
    nb = profile.get("numbering")
    if not nb: return
    cl = nb.get("clause"); hnum = core.heading_num_ids(doc)
    unnum = re.compile(nb["unnumbered"]); manual = re.compile(nb["manual_prefix"])
    lv_of = {lv["role"]: k for k, lv in enumerate(nb["levels"])}
    for p, role in core.roles(doc, profile):
        if role not in lv_of: continue
        t = p.text.strip(); nid, ilvl = core.para_numpr(p)
        if role == "heading1" and unnum.match(t):
            if nid in hnum: out.append(_iss("num_unexpected", f"「{t[:10]}」", "不编号", "被编号", "参考文献/附录/致谢不编章号", cl))
            continue
        if nid not in hnum or ilvl != lv_of[role]:
            out.append(_iss("num_manual", f"「{t[:12]}」", "Word 自动编号（多级列表）", "手写编号" if manual.match(t) else f"numId={nid}", "标题自动编号", cl))
        elif manual.match(t):
            out.append(_iss("num_double", f"「{t[:12]}」", "仅自动编号", "自动编号+手写编号", "重复编号", cl))
    # 自动编号定义本身是否符合 1.7.1
    nid = _guard_num_id(doc)
    if nid:
        num = _numbering_part(doc).element
        absid = next(n.find(qn("w:abstractNumId")).get(qn("w:val")) for n in num.findall(qn("w:num")) if n.get(qn("w:numId")) == nid)
        ab = next(a for a in num.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == absid)
        for k, lv in enumerate(nb["levels"]):
            l = ab.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='{k}']")
            txt = l.find(qn("w:lvlText")).get(qn("w:val")); fmt = l.find(qn("w:numFmt")).get(qn("w:val"))
            if txt != lv["text"] or fmt != lv["fmt"]:
                out.append(_iss("num_definition", f"第{k+1}级", f"{lv['fmt']} {lv['text']!r}", f"{fmt} {txt!r}", "编号格式", cl))

def _repair_heading_list(doc, nb, nid):
    """GuardHeadings 定义被改动（如 lvlText 改成 Chapter %1）时按规则写回各级编号格式。"""
    num = _numbering_part(doc).element
    absid = next(n.find(qn("w:abstractNumId")).get(qn("w:val")) for n in num.findall(qn("w:num")) if n.get(qn("w:numId")) == nid)
    ab = next(a for a in num.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == absid)
    for k, lv in enumerate(nb["levels"]):
        l = ab.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='{k}']")
        l.find(qn("w:lvlText")).set(qn("w:val"), lv["text"]); l.find(qn("w:numFmt")).set(qn("w:val"), lv["fmt"])

def _ensure_heading_list(doc, nb):
    nid = _guard_num_id(doc)
    if nid: _repair_heading_list(doc, nb, nid); return nid
    num = _numbering_part(doc).element
    absid = str(max([int(a.get(qn("w:abstractNumId"))) for a in num.findall(qn("w:abstractNum"))] + [0]) + 1)
    lvls = ""
    for k, lv in enumerate(nb["levels"]):
        lvls += (f'<w:lvl w:ilvl="{k}"><w:start w:val="1"/><w:numFmt w:val="{lv["fmt"]}"/>'
                 + (f'<w:lvlRestart w:val="{k}"/>' if k else "")
                 + ('<w:isLgl/>' if lv.get("legal") else "")
                 + f'<w:suff w:val="nothing"/><w:lvlText w:val="{lv["text"]}"/><w:lvlJc w:val="left"/>'
                 '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>')
    ab = parse_xml(f'<w:abstractNum xmlns:w="{W}" w:abstractNumId="{absid}"><w:multiLevelType w:val="multilevel"/>'
                   f'<w:name w:val="GuardHeadings"/>{lvls}</w:abstractNum>')
    first_num = num.find(qn("w:num"))
    if first_num is not None: first_num.addprevious(ab)
    else: num.append(ab)
    nid = str(max([int(n.get(qn("w:numId"))) for n in num.findall(qn("w:num"))] + [0]) + 1)
    num.append(parse_xml(f'<w:num xmlns:w="{W}" w:numId="{nid}"><w:abstractNumId w:val="{absid}"/></w:num>'))
    return nid

def _set_numpr(p, nid, ilvl):
    ppr = p._p.get_or_add_pPr(); old = ppr.find(qn("w:numPr"))
    if old is not None: ppr.remove(old)
    np_ = OxmlElement("w:numPr"); il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(ilvl)); ni = OxmlElement("w:numId"); ni.set(qn("w:val"), nid)
    np_.append(il); np_.append(ni)
    anchor = next((c for c in ppr if c.tag.split("}")[1] not in ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl")), None)
    if anchor is not None: anchor.addprevious(np_)
    else: ppr.append(np_)

def fix_numbering(doc, profile, changes):
    nb = profile.get("numbering")
    if not nb: return
    unnum = re.compile(nb["unnumbered"]); manual = re.compile(nb["manual_prefix"])
    lv_of = {lv["role"]: k for k, lv in enumerate(nb["levels"])}
    targets = [(p, role) for p, role in core.roles(doc, profile) if role in lv_of]
    nid = _ensure_heading_list(doc, nb); n = 0
    for p, role in targets:
        t = p.text.strip()
        if role == "heading1" and unnum.match(t):
            _set_numpr(p, "0", 0); continue
        m = manual.match(t)
        if m: _replace_in_runs(p, p.text[:p.text.index(t) + len(m.group(0))], "")
        for r in p.runs:
            if r.text and r.text != r.text.lstrip(): r.text = r.text.lstrip(); break
            if r.text: break
        _set_numpr(p, nid, lv_of[role]); n += 1
    # 样式级 numPr（模板自带的、未定义的列表）去掉，避免叠加编号
    for role in lv_of:
        st = core.find_style(doc, profile["styles"][role]["style"])   # 名式不敏感：解析不到才跳过，不许把名式差异当缺失
        if st is None: continue
        ppr = st.element.pPr
        if ppr is not None and ppr.find(qn("w:numPr")) is not None: ppr.remove(ppr.find(qn("w:numPr")))
    changes.append(f"标题自动编号：{n} 个标题改为多级列表编号（第X章 / x.y / x.y.z），去掉手写编号")

# ---------- 不受管段落的版式冻结 ----------
# 封面、版权声明、原创性声明、模板说明文字等不属于任何角色，但大多继承 Normal。
# 修复会按正文规则重写 Normal（首行缩进 2 字符、字体），这会让封面标签折行。
# 因此修复前记录它们的有效格式，修复后若被样式改动波及，就写回为直接格式。
def _snap(doc, p):
    runs = [r for r in core.all_runs(p) if r.text]
    sz = core.eff_size(doc, p, runs[0]) if runs else 12
    return dict(first=core.indent_first_pt(p, sz or 12), left=core.eff_pf(p, "left_indent"),
                pf={k: core.eff_pf(p, k) for k in ("line_spacing", "line_spacing_rule", "space_before", "space_after", "alignment")},
                runs=[(r._r, core.eff_size(doc, p, r), core.eff_font(doc, p, r, "eastAsia"), core.eff_font(doc, p, r, "ascii")) for r in runs])

def _unmanaged(doc, profile):
    regs = {}; rl = dict((p._p, r) for p, r in core.roles(doc, profile, regs))
    keep = set(profile.get("frozen_regions", ["cover", "copyright", "declaration"]))
    for el in doc.element.body.iter(qn("w:p")):
        top = el
        while top.getparent() is not doc.element.body: top = top.getparent()
        reg = regs.get(top if top.tag == qn("w:p") else el)
        if top.tag == qn("w:tbl"):
            prev = top.getprevious()
            while prev is not None and prev.tag != qn("w:p"): prev = prev.getprevious()
            reg = regs.get(prev) if prev is not None else "cover"
            reg = reg or "cover"
            if reg not in keep: continue
        elif top.tag != qn("w:p"): continue
        elif rl.get(el) is not None or (reg not in keep and reg is not None): continue
        if core.is_toc_para(P(doc, el)): continue
        yield el

def freeze_unmanaged(doc, profile):
    doc._guard_frozen = [(el, _snap(doc, P(doc, el))) for el in _unmanaged(doc, profile)]

def restore_unmanaged(doc, profile, changes):
    n = 0
    for el, old in getattr(doc, "_guard_frozen", []):
        if el.getparent() is None: continue
        p = P(doc, el); new = _snap(doc, p)
        if abs(new["first"] - old["first"]) > 0.1:
            p.paragraph_format.first_line_indent = Pt(old["first"]); n += 1
            ind = p._p.pPr.find(qn("w:ind"))
            for a in ("firstLineChars", "hangingChars"):
                if ind is not None and ind.get(qn(f"w:{a}")) is not None: del ind.attrib[qn(f"w:{a}")]
        # 行距/段前后/对齐：Normal 改为“固定 20 磅”后，封面校徽（内嵌图片）所在段会被裁切、36pt 标题行会重叠
        for k, v in old["pf"].items():
            if new["pf"][k] != v:
                pf = p.paragraph_format
                if k == "line_spacing_rule":
                    if v is None: v = WD_LINE_SPACING.SINGLE
                    if v in (WD_LINE_SPACING.SINGLE, WD_LINE_SPACING.ONE_POINT_FIVE, WD_LINE_SPACING.DOUBLE): pf.line_spacing_rule = v
                    elif old["pf"]["line_spacing"] is not None: pf.line_spacing = old["pf"]["line_spacing"]
                    if v != WD_LINE_SPACING.EXACTLY and old["pf"]["line_spacing"] is not None and not isinstance(old["pf"]["line_spacing"], float):
                        pf.line_spacing_rule = v; pf.line_spacing = old["pf"]["line_spacing"]
                elif k == "line_spacing":
                    if v is None: pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    else:
                        pf.line_spacing = v
                        if old["pf"]["line_spacing_rule"] is not None: pf.line_spacing_rule = old["pf"]["line_spacing_rule"]
                elif k == "alignment":
                    if v is not None: pf.alignment = v
                else: setattr(pf, k, v if v is not None else Pt(0))
                n += 1
        for (r_el, sz, ea, la) in old["runs"]:
            from docx.text.run import Run
            r = Run(r_el, p)
            if core.eff_size(doc, p, r) != sz and sz: r.font.size = Pt(sz); n += 1
            if core.eff_font(doc, p, r, "eastAsia") != ea or core.eff_font(doc, p, r, "ascii") != la:
                rpr = r_el.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
                if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0 if rpr.find(qn("w:rStyle")) is None else 1, rf)
                if ea and not ea.startswith("theme:"): rf.set(qn("w:eastAsia"), ea)
                if la and not la.startswith("theme:"):
                    for sl in ("ascii", "hAnsi", "cs"): rf.set(qn(f"w:{sl}"), la)
                n += 1
    if n: changes.append(f"封面/声明等不受管段落：{n} 处因 Normal 样式改动而写回原有效格式（版式不变）")

# ---------- 汇总 ----------
CHECKS = [check_page, check_header_extra, check_page_numbers, check_cover, check_fixed_text, check_abstract,
          check_toc, check_footnotes, check_tables, check_equations, check_numbering,
          check_blank_pages, check_toc_leaks, check_cover_layout, check_keep_together, check_empty_filler]

def check_all(doc, profile):
    out = []
    for f in CHECKS: f(doc, profile, out)
    if profile.get("references"): out += check_references(doc, profile)
    return out

def fix_pre(doc, profile, changes):
    _ensure_default_para_style(doc)
    fix_fixed_text(doc, profile, changes)
    freeze_unmanaged(doc, profile)
    fix_numbering(doc, profile, changes)
    fix_tables(doc, profile, changes)
    fix_equations(doc, profile, changes)
    if profile.get("references"):
        fix_references(doc, profile, changes)
        from .references import finalize_cites; finalize_cites(doc)
    fix_abstract(doc, profile, changes)
    fix_page_numbers(doc, profile, changes)
    fix_blank_pages(doc, profile, changes)

def _ensure_default_para_style(doc):
    """LibreOffice 转出的文档可能没有 w:default="1" 的段落样式，无 pStyle 的段落会找不到样式。"""
    from docx.enum.style import WD_STYLE_TYPE
    if doc.styles.default(WD_STYLE_TYPE.PARAGRAPH) is None:
        core.require_style(doc, "Normal").element.set(qn("w:default"), "1")

def fix_post(doc, profile, changes):
    _ensure_default_para_style(doc)
    restore_unmanaged(doc, profile, changes)
    fix_page(doc, profile, changes)
    fix_toc_leaks(doc, profile, changes)          # 先建“不进目录”样式，页眉 STYLEREF 要引用它
    fix_header_extra(doc, profile, changes)
    fix_cover(doc, profile, changes)
    pin_cover(doc, profile, changes)
    fix_keep_together(doc, profile, changes)
    fix_footnotes(doc, profile, changes)
    fix_toc(doc, profile, changes)
