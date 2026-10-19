"""Coverage ledger: what a check actually covers — and what it cannot see.

check() verifies only the attributes a profile rule names. Everything else
is a silent gap: drift there passes unnoticed and the report still reads
"checked". coverage() censuses the document itself (styles in use, roles
assigned, sections) against the profile and sorts every dimension into
three states — ruled, waived (profile lists it under `waive:`), or
UNCOVERED — so a result can never claim completeness without stating its
blind spots. A paragraph whose style the profile does not know is counted
as unmapped instead of being silently treated as body text.
"""
from collections import Counter
from docx import Document
from .core import roles

CHECKABLE = {
    "east_asia": ("east_asia",), "latin": ("latin",), "size_pt": ("size_pt",),
    "bold": ("bold",), "line_spacing": ("line_exact_pt", "line_spacing"),
    "first_line_indent": ("first_line_indent_pt", "first_line_indent_chars"),
    "space_before_pt": ("space_before_pt",), "space_after_pt": ("space_after_pt",),
    "align": ("align",),
}
STRUCTURAL_PREFIXES = ("toc", "header", "footer", "page number")

def _state(name, ruled, waive):
    if name in waive: return "waived"
    return "ruled" if ruled else "uncovered"

def coverage(docx_path, profile):
    doc = Document(docx_path)
    waive = set(profile.get("waive", []))
    prof_styles = profile.get("styles", {})
    ruled_names = {r.get("style") for r in prof_styles.values()}
    census = Counter(p.style.name for p in doc.paragraphs if p.style is not None)
    unmapped = {}
    for name, n in census.items():
        if name in ruled_names or name.lower().startswith(STRUCTURAL_PREFIXES): continue
        if f"style:{name}" in waive: continue
        unmapped[name] = n
    role_counts = Counter(r for _, r in roles(doc, profile) if r)
    per_role, blind = {}, 0
    for role, n in sorted(role_counts.items()):
        rule = prof_styles.get(role, {})
        covered, uncovered = [], []
        for attr, keys in CHECKABLE.items():
            if any(k in rule for k in keys): covered.append(attr)
            elif f"{role}.{attr}" in waive: pass
            else: uncovered.append(attr)
        per_role[role] = {"paragraphs": n, "covered_attrs": covered, "uncovered_attrs": uncovered}
        blind += len(uncovered)
    areas = {}
    for area, key in (("page_setup", "page"), ("header", "header"),
                      ("page_numbers", "page_numbers"), ("cover_fixed_texts", "cover")):
        areas[area] = _state(area, bool(profile.get(key)), waive)
        if areas[area] == "uncovered": blind += 1
    blind += len(unmapped)
    return {"paragraphs": int(sum(census.values())), "roles": per_role, "areas": areas,
            "sections": len(doc.sections), "unmapped_styles": unmapped,
            "blind_spots": blind, "verdict": "complete" if blind == 0 else "partial"}

def format_coverage(cov):
    lines = [f"覆盖台账：共 {cov['paragraphs']} 段、{cov['sections']} 节"]
    for role, v in cov["roles"].items():
        lines.append(f"  角色 {role}（{v['paragraphs']} 段）：已检 {len(v['covered_attrs'])} 项"
                     + (f"，未覆盖 {', '.join(v['uncovered_attrs'])}" if v["uncovered_attrs"] else ""))
    for area, st in cov["areas"].items():
        lines.append(f"  {area}: {'已检' if st == 'ruled' else '已豁免' if st == 'waived' else '未覆盖'}")
    if cov["unmapped_styles"]:
        lines.append("  未识别样式（规则包不认识、按正文吞掉也不会报）："
                     + ", ".join(f"{k}×{v}" for k, v in cov["unmapped_styles"].items()))
    lines.append(f"结论：{'完整覆盖' if cov['verdict'] == 'complete' else '部分覆盖'}，盲区 {cov['blind_spots']} 处"
                 "（未覆盖≠没问题，只是没检）")
    return "\n".join(lines)
