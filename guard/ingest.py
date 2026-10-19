"""Derive a DRAFT rule profile from a school template .docx.

Reads what a template file mechanically encodes: section page setup, named
style definitions (fonts, size, line spacing, indents, alignment), header
text/format and per-section page-number formats. What lives only in the
written guide -- which regions use roman vs arabic numerals, reference-list
semantics, footnote numbering rules, clause citations -- cannot be derived
from a file and is listed in the report instead of being guessed.

The output profile is a draft: values are read from the template itself.
Note that official templates can deviate from their own written guides
(the PKU template does, in 154 measured places), so a human/agent review
pass against the guide is required before a derived profile is trusted.
"""
import os
from docx import Document
from docx.oxml.ns import qn
import yaml

ROLE_BY_NAME = [
    ("body", ("normal", "正文")),
    ("heading1", ("heading 1", "标题 1")),
    ("heading2", ("heading 2", "标题 2")),
    ("heading3", ("heading 3", "标题 3")),
    ("caption", ("caption", "题注")),
    ("reference", ("reference", "参考文献")),
    ("abstract", ("abstract", "摘要")),
]

def _chain(style):
    seen, out = set(), []
    s = style
    while s is not None and id(s.element) not in seen:
        seen.add(id(s.element)); out.append(s)
        s = s.base_style
    return out

def _first(chain, xpath):
    for st in chain:
        el = st.element.find(xpath)
        if el is not None:
            return el
    return None

def _style_rule(style):
    ch = _chain(style)
    rule = {"style": style.name}
    rpr = _first(ch, qn("w:rPr"))
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            if rf.get(qn("w:eastAsia")): rule["east_asia"] = rf.get(qn("w:eastAsia"))
            if rf.get(qn("w:ascii")): rule["latin"] = rf.get(qn("w:ascii"))
        sz = rpr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            rule["size_pt"] = int(sz.get(qn("w:val"))) / 2
        b = rpr.find(qn("w:b"))
        if b is not None and b.get(qn("w:val")) not in ("0", "false", "off"):
            rule["bold"] = True
    ppr = _first(ch, qn("w:pPr"))
    if ppr is not None:
        sp = ppr.find(qn("w:spacing"))
        if sp is not None:
            line, rule_attr = sp.get(qn("w:line")), sp.get(qn("w:lineRule"))
            if line:
                v = int(line)
                if rule_attr == "exact": rule["line_exact_pt"] = v / 20
                elif rule_attr == "atLeast": rule["line_spacing"] = v / 20
                else: rule["line_spacing"] = round(v / 240, 3)
            if sp.get(qn("w:before")): rule["space_before_pt"] = int(sp.get(qn("w:before"))) / 20
            if sp.get(qn("w:after")): rule["space_after_pt"] = int(sp.get(qn("w:after"))) / 20
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            if ind.get(qn("w:firstLine")): rule["first_line_indent_pt"] = int(ind.get(qn("w:firstLine"))) / 20
            elif ind.get(qn("w:firstLineChars")): rule["first_line_indent_chars"] = int(ind.get(qn("w:firstLineChars"))) / 100
            if ind.get(qn("w:hanging")): rule["hanging_indent_pt"] = int(ind.get(qn("w:hanging"))) / 20
        jc = ppr.find(qn("w:jc"))
        if jc is not None and jc.get(qn("w:val")):
            rule["align"] = {"both": "justify", "distribute": "justify"}.get(jc.get(qn("w:val")), jc.get(qn("w:val")))
    return rule

def _paper_name(w_cm, h_cm):
    for name, (pw, ph) in {"A4": (21.0, 29.7), "A3": (29.7, 42.0), "Letter": (21.59, 27.94), "Legal": (21.59, 35.56)}.items():
        if abs(w_cm - pw) < 0.3 and abs(h_cm - ph) < 0.3: return name
    return f"{w_cm:.1f}x{h_cm:.1f}cm"

def _header_info(section):
    def read(hdr):
        for p in hdr.paragraphs:
            t = "".join(r.text for r in p.runs).strip()
            if not t: continue
            info = {"text": t}
            if p.runs:
                rpr = p.runs[0]._r.find(qn("w:rPr"))
                if rpr is not None:
                    rf = rpr.find(qn("w:rFonts"))
                    if rf is not None and rf.get(qn("w:eastAsia")): info["east_asia"] = rf.get(qn("w:eastAsia"))
                    sz = rpr.find(qn("w:sz"))
                    if sz is not None and sz.get(qn("w:val")): info["size_pt"] = int(sz.get(qn("w:val"))) / 2
            if p.alignment is not None: info["align"] = str(p.alignment).split(" ")[0].lower().replace("both", "justify")
            return info
        return None
    odd = read(section.header)
    even = None
    try: even = read(section.even_page_header)
    except Exception: pass
    return odd, even

def derive_profile(template_path):
    doc = Document(template_path)
    notes, derived = [], []
    name = os.path.splitext(os.path.basename(template_path))[0]
    prof = {"profile": f"{name}-draft", "derived_from_template": os.path.basename(template_path),
            "source": "AUTO-DERIVED DRAFT from the template file itself. Values below were read out of the "
                      "template; rules that exist only in the written guide are NOT included (see ingest "
                      "report). Review against the guide before trusting this profile."}
    secs = []
    for i, s in enumerate(doc.sections):
        pg = s._sectPr.find(qn("w:pgNumType"))
        secs.append({"i": i, "top": round(s.top_margin.cm, 2), "bottom": round(s.bottom_margin.cm, 2),
                     "left": round(s.left_margin.cm, 2), "right": round(s.right_margin.cm, 2),
                     "header_cm": round(s.header_distance.cm, 2), "footer_cm": round(s.footer_distance.cm, 2),
                     "paper": _paper_name(round(s.page_width.cm, 2), round(s.page_height.cm, 2)),
                     "pgfmt": (pg.get(qn("w:fmt")) if pg is not None else None)})
    main = max(secs, key=lambda d: sum(1 for o in secs if (o["top"], o["bottom"], o["left"], o["right"]) == (d["top"], d["bottom"], d["left"], d["right"])))
    prof["page"] = {"margin_cm": {"top": main["top"], "bottom": main["bottom"], "left": main["left"], "right": main["right"]},
                    "header_distance_cm": main["header_cm"], "footer_distance_cm": main["footer_cm"], "paper": main["paper"]}
    derived.append(f"page setup (from section #{main['i']}, the most common margins across {len(secs)} sections)")
    if len({(s['top'], s['bottom'], s['left'], s['right']) for s in secs}) > 1:
        notes.append("sections disagree on margins; the most common set was used, all sections are listed in the report")
    styles, unmapped = {}, []
    by_lower = {st.name.lower(): st for st in doc.styles if st.type is not None and str(st.type).startswith("PARAGRAPH")}
    used = set()
    for role, aliases in ROLE_BY_NAME:
        hit = next((by_lower[a] for a in aliases if a in by_lower), None)
        if hit is not None:
            styles[role] = _style_rule(hit); used.add(hit.name.lower()); derived.append(f"style {role} <- '{hit.name}'")
    for nm, st in by_lower.items():
        if nm not in used: unmapped.append(st.name)
    if styles: prof["styles"] = styles
    odd, even = _header_info(doc.sections[main["i"]])
    if odd:
        prof["header"] = odd; derived.append("header text/format")
        if even and even.get("text") and even["text"] != odd.get("text"):
            prof["header"]["even_text"] = even["text"]; prof["header"]["even_odd"] = True
    fmts = [s["pgfmt"] for s in secs]
    not_derived = ["region semantics (which parts are front matter / body) and roman-vs-arabic assignment per region",
                   "heading/caption numbering patterns and fixed texts (cover, declarations)",
                   "footnote/endnote numbering and format rules",
                   "reference-list semantics (citation style mixing, entry parsing)",
                   "clause citations — a derived profile has no guide clause numbers"]
    report = {"template": os.path.basename(template_path), "derived": derived,
              "sections": secs, "page_number_formats_by_section": fmts,
              "styles_found_not_mapped": sorted(unmapped), "not_derived": not_derived, "notes": notes}
    return prof, report

def write_profile(prof, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(prof, f, allow_unicode=True, sort_keys=False)

def format_report(rep):
    lines = [f"ingest report for {rep['template']}", "derived:"]
    lines += [f"  + {d}" for d in rep["derived"]]
    lines.append("sections (margins cm / paper / page-number format):")
    for s in rep["sections"]:
        lines.append(f"  #{s['i']}: {s['top']}/{s['bottom']}/{s['left']}/{s['right']} {s['paper']} pgfmt={s['pgfmt']}")
    if rep["styles_found_not_mapped"]:
        lines.append("styles present but not mapped to roles: " + ", ".join(rep["styles_found_not_mapped"][:20]))
    lines.append("NOT derived (need the written guide):")
    lines += [f"  - {d}" for d in rep["not_derived"]]
    lines += [f"note: {n}" for n in rep["notes"]]
    return "\n".join(lines)
