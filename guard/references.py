"""参考文献内容检查（GB/T 7714-2015；兼容北大指南基于 2005 版的可选项）。
能安全修复的：半角标点、书名号、页码范围连接符、作者超 3 人截断、等/et al 语言、末尾句点、可高置信推断的文献类型标识、
顺序编码制下的重新编号与按首次引用顺序重排（正文引用全部可解析时）、著者-出版年制排序。
只能给建议的：无法推断类型、结构缺项（年卷期页、出版地出版者、引用日期）、西文姓名顺序、未被引用的文献、混用两种著录制。"""
import re
from . import core

VALID_TAGS = {"M", "C", "G", "N", "J", "D", "R", "S", "P", "DB", "CP", "EB", "A", "Z"}
TAG_RE = re.compile(r"\[([A-Z]{1,2})(?:/(OL|DK|MT|CD))?\]")
NUM_RE = re.compile(r"^\s*[\[［]\s*(\d+)\s*[\]］]\s*")
CITE_RE = re.compile(r"\[(\d+(?:\s*[-–~,，]\s*\d+)*)\]")
URL_RE = re.compile(r"(https?://|www\.)\S+", re.I)
CJK = re.compile(r"[\u4e00-\u9fff]")
FW = [("，", ", "), ("。", ". "), ("．", ". "), ("：", ": "), ("；", "; "), ("（", "("), ("）", ")"), ("［", "["), ("］", "]"), ("／", "/")]

def _iss(kind, where, expect, actual, cause, fixable=True, suggestion=None, clause=None):
    return dict(kind=kind, where=where, expect=expect, actual=actual, cause=cause, clause=clause, fixable=fixable, suggestion=suggestion)

def ref_paras(doc, profile):
    return [p for p, r in core.roles(doc, profile) if r == "reference"]

def _split(t):
    """→ (编号或 None, 主体)。"""
    m = NUM_RE.match(t)
    return (int(m.group(1)), t[m.end():]) if m else (None, t.strip())

def _author_end(body):
    """作者项结束位置：著者-出版年制为年份前；否则为第一个不属于姓名缩写的 ". "。"""
    m = re.match(r"^(.*?)[.,]?\s*(\d{4}[a-z]?)\.\s", body)
    if m and not TAG_RE.search(m.group(1)) and len(m.group(1)) < 200: return len(m.group(1))
    for m in re.finditer(r"\.\s", body):
        i = m.start()
        if i >= 1 and body[i - 1].isupper() and (i == 1 or not body[i - 2].isalpha()): continue   # J. / D.G.
        return i
    return None

def authors_of(body):
    e = _author_end(body)
    if e is None: return None, None
    seg = body[:e]
    if TAG_RE.search(seg) or not seg.strip(): return None, None
    parts = re.split(r";\s*|；", seg) if ";" in seg or "；" in seg else re.split(r",\s*(?:and\s+)?|，|\s+and\s+", seg)
    return [a.strip() for a in parts if a.strip()], seg

def _is_cjk(s): return bool(CJK.search(s))

def infer_tag(body):
    """高置信推断类型标识；返回 (标识, 插入位置) 或 (None, None)。"""
    url = URL_RE.search(body)
    m = re.search(r"\.\s+([^.]+?),\s*(\d{4})\s*,\s*\d+\s*(\(\d+\))?\s*:\s*\d+", body)          # 刊名, 年, 卷(期): 页
    if m: return ("J/OL" if url else "J"), m.start()
    m = re.search(r"\.\s+([^.:]+?):\s*([^,.]*?(大学|学院|研究所|University|Institute)[^,.]*),\s*\d{4}", body)   # 学位论文
    if m and re.search(r"(学位论文|硕士|博士|dissertation|thesis)", body, re.I) is None and False: pass
    if m and re.search(r"(硕士|博士|学位|dissertation|thesis)", body, re.I): return "D", m.start()
    m = re.search(r"\.\s+([^.:]+?):\s*([^,.]*?(出版社|Press|Publishing|Publishers))\s*,\s*\d{4}", body)
    if m: return "M", m.start()
    if url and body.count(". ") <= 2:
        a = re.search(r"\.\s+(?=\S)", body)
        rest = body[a.end():] if a else body
        b = re.search(r"\.\s", rest)
        if b: return "EB/OL", (a.end() if a else 0) + b.start()
    return None, None

def _norm_punct(body):
    t = body
    for a, b in FW: t = t.replace(a, b)
    t = re.sub(r"[《》〈〉]", "", t)
    t = re.sub(r"(?<=[\u4e00-\u9fff])\.(?=[^\s])", ". ", t)          # 中文后的著录分隔符“.”后补空格
    t = re.sub(r"(\d)\s*[~～–—－]\s*(\d)", r"\1-\2", t)
    t = re.sub(r"\bpp?\.\s*(?=\d)", "", t)
    t = re.sub(r"[ \t]{2,}", " ", t).replace(" ,", ",").replace(" .", ".").replace("( ", "(").replace(" )", ")")
    return t.strip()

def _fix_authors(body, max_n):
    au, seg = authors_of(body)
    if not au: return body
    cn = _is_cjk(seg); etal = [a for a in au if a in ("等", "et al", "et al.")]
    names = [a for a in au if a not in ("等", "et al", "et al.")]
    if len(names) > max_n: names = names[:max_n]; etal = ["x"]
    sep = "; " if ";" in seg else ", "
    new = sep.join(names) + ((sep + "等" if cn else sep + "et al") if etal else "")
    return new + body[len(seg):]

def check_entry(body, rc, idx):
    out = []; w = f"文献{idx}「{body[:16]}」"; cl = rc.get("clause")
    if _norm_punct(body) != body.strip():
        if re.search(r"[《》〈〉]", body): out.append(_iss("ref_booktitle_marks", w, "不加书名号", "含书名号", "A.4.3", clause=cl))
        if re.search("|".join(re.escape(a) for a, _ in FW), body): out.append(_iss("ref_fullwidth", w, "半角标点", "含全角标点", "1.8 标点用半角", clause=cl))
        if re.search(r"\d\s*[~～–—－]\s*\d|\bpp?\.\s*\d", body): out.append(_iss("ref_page_range", w, "起止页码用 - 连接", "其他连接符或 pp.", "页码", clause=cl))
        if not out: out.append(_iss("ref_spacing", w, "规范空格", "多余空格", "格式", clause=cl))
    tags = TAG_RE.findall(body)
    if tags:
        for t, carrier in tags:
            if t not in VALID_TAGS: out.append(_iss("ref_bad_tag", w, "/".join(sorted(VALID_TAGS)), t, "文献类型标识", fixable=False, suggestion="改为正确的类型标识", clause=cl))
        if URL_RE.search(body) and not any(c == "OL" for _, c in tags):
            out.append(_iss("ref_online_carrier", w, "电子资源加 /OL", f"[{tags[0][0]}]", "电子资源载体标识", clause=cl))
    else:
        need = rc.get("require_type_tag") or URL_RE.search(body)
        if need:
            tag, _ = infer_tag(body)
            out.append(_iss("ref_type_tag", w, "[J]/[M]/[D]/[EB/OL]…", "缺失", "文献类型标识（GB/T 7714-2015 必备）",
                            fixable=tag is not None, suggestion=None if tag else "无法从结构可靠推断类型，请手工补写", clause=cl))
    au, seg = authors_of(body)
    if au:
        names = [a for a in au if a not in ("等", "et al", "et al.")]
        if len(names) > rc["max_authors"]:
            out.append(_iss("ref_authors", w, f"前 {rc['max_authors']} 位 + 等/et al", f"{len(names)} 位", "A.4.4 / 2015 版 8.1.2", clause=cl))
        if "等" in au and not _is_cjk(", ".join(names)) or any(a.startswith("et al") for a in au) and _is_cjk(", ".join(names)):
            out.append(_iss("ref_etal_lang", w, "中文用「等」、西文用「et al」", seg, "等/et al 语种", clause=cl))
        for a in names:
            if re.fullmatch(r"[A-Z][a-z]+ [A-Z][a-z]+", a):
                out.append(_iss("ref_western_name", w, "姓在前、名缩写（如 SMITH J）", a, "西文姓名", fixable=False,
                                suggestion="工具无法区分姓与名，请手工确认（例：John Smith → SMITH J）", clause=cl))
    tag = tags[0][0] if tags else None
    if tag == "J" and not re.search(r"\]\.\s*[^,]+,\s*\d{4}(\s*,\s*\d+)?(\s*\(\d+\))?\s*:\s*\d+(-\d+)?", _norm_punct(body)):
        out.append(_iss("ref_j_fields", w, "刊名, 年, 卷(期): 起止页", "缺项或顺序不符", "A.1 期刊著录格式", fixable=False, suggestion="补全年、卷、期与页码", clause=cl))
    if tag == "M" and not re.search(r"\]\.\s*(\d+版\.\s*|\d+(st|nd|rd|th) ed\.\s*)?[^:]+:\s*[^,]+,\s*\d{4}", _norm_punct(body)):
        out.append(_iss("ref_m_fields", w, "出版地: 出版者, 出版年", "缺项或顺序不符", "A.1 专著著录格式", fixable=False, suggestion="补全出版地、出版者、出版年", clause=cl))
    if tags and any(c == "OL" for _, c in tags) or (tag == "EB"):
        if not re.search(r"\[\d{4}-\d{2}-\d{2}\]", body):
            out.append(_iss("ref_cite_date", w, "[引用日期 YYYY-MM-DD]", "缺失", "电子资源的引用日期", fixable=False, suggestion="补写访问日期，工具无法得知", clause=cl))
        if not URL_RE.search(body):
            out.append(_iss("ref_url", w, "获取和访问路径", "缺失", "电子资源 URL", fixable=False, suggestion="补写 URL", clause=cl))
    if not body.rstrip().endswith("."):
        out.append(_iss("ref_end_period", w, "以 . 结尾", body.rstrip()[-1:], "著录结尾", clause=cl))
    return out

def fix_entry(body, rc):
    t = _norm_punct(body)
    t = _fix_authors(t, rc["max_authors"])
    au, seg = authors_of(t)
    if au:
        names = [a for a in au if a not in ("等", "et al", "et al.")]
        if any(a in ("等", "et al", "et al.") for a in au):
            sep = "; " if ";" in seg else ", "
            t = (sep.join(names) + sep + ("等" if _is_cjk(", ".join(names)) else "et al")) + t[len(seg):]
    tags = TAG_RE.findall(t)
    if not tags and (rc.get("require_type_tag") or URL_RE.search(t)):
        tag, pos = infer_tag(t)
        if tag: t = t[:pos] + f"[{tag}]" + t[pos:]
    elif tags and URL_RE.search(t) and not any(c == "OL" for _, c in tags):
        m = TAG_RE.search(t); t = t[:m.start()] + f"[{m.group(1)}/OL]" + t[m.end():]
    if not t.endswith("."): t += "."
    return t

def _citations(doc, profile):
    res = []
    for p, r in core.roles(doc, profile):
        if r in ("reference",) or r is None: continue
        for m in CITE_RE.finditer(p.text): res.append((p, m))
    return res

def _expand(s):
    out = []
    for part in re.split(r"\s*[,，]\s*", s):
        a = re.split(r"\s*[-–~]\s*", part)
        if len(a) == 2 and a[0].isdigit() and a[1].isdigit(): out += list(range(int(a[0]), int(a[1]) + 1))
        elif part.isdigit(): out.append(int(part))
        else: return None
    return out

def _compress(nums):
    nums = sorted(set(nums)); out = []; i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1: j += 1
        out.append(str(nums[i]) if j == i else (f"{nums[i]},{nums[j]}" if j == i + 1 else f"{nums[i]}-{nums[j]}")); i = j + 1
    return ",".join(out)

def first_order(doc, profile):
    seen = []
    for p, m in _citations(doc, profile):
        ns = _expand(m.group(1))
        if ns is None: return None
        for n in ns:
            if n not in seen: seen.append(n)
    return seen

def _sort_key(body):
    from pypinyin import lazy_pinyin
    au, seg = authors_of(body); s = seg or body
    return (0 if _is_cjk(s) else 1, " ".join(lazy_pinyin(s)).lower() if _is_cjk(s) else s.lower())

def check_references(doc, profile):
    rc = profile["references"]; cl = rc.get("clause"); out = []
    paras = ref_paras(doc, profile)
    if not paras: return out
    parsed = [_split(p.text.strip()) for p in paras]
    numbered = [n is not None for n, _ in parsed]
    if any(numbered) and not all(numbered):
        out.append(_iss("ref_mixed_system", "参考文献", "全文统一一种著录制", f"{sum(numbered)} 条顺序编码 + {len(numbered)-sum(numbered)} 条著者-出版年",
                        "1.8 顺序编码制与著者-出版年制二选一", fixable=False, suggestion="选定一种体系后手工统一（混用时正文引用形式也不同，工具无法安全转换）", clause=cl))
    for k, (n, body) in enumerate(parsed): out += check_entry(body, rc, n if n is not None else k + 1)
    if all(numbered):
        nums = [n for n, _ in parsed]
        if nums != list(range(1, len(nums) + 1)):
            out.append(_iss("ref_numbering", "参考文献", "1..n 连续编号", str(nums[:12]), "顺序编码", clause=cl))
        order = first_order(doc, profile)
        if order is None:
            out.append(_iss("ref_cite_unparsed", "正文引用", "可解析的 [n]", "存在无法解析的引用", "顺序编码", fixable=False, suggestion="手工检查引用标注", clause=cl))
        else:
            missing = [n for n in order if n not in nums]
            if missing: out.append(_iss("ref_cite_dangling", "正文引用", "引用的编号存在", str(missing), "引用了不存在的文献", fixable=False, suggestion="补充文献或修正引用", clause=cl))
            uncited = [n for n in nums if n not in order]
            if uncited: out.append(_iss("ref_uncited", "参考文献", "每条都被引用", str(uncited), "未被正文引用", fixable=False, suggestion="删除或在正文中引用（内容取舍工具不代做）", clause=cl))
            cited = [n for n in nums if n in order]
            if not missing and [n for n in order] != sorted(order, key=lambda x: nums.index(x)) or (not missing and order != list(range(1, len(order) + 1)) and nums == list(range(1, len(nums) + 1))):
                out.append(_iss("ref_order", "参考文献", "按正文首次引用顺序编号", f"首次引用顺序 {order[:10]}", "顺序编码制", clause=cl))
    elif not any(numbered):
        keys = [_sort_key(b) for _, b in parsed]
        if keys != sorted(keys):
            out.append(_iss("ref_ay_order", "参考文献", "中文在前按拼音、西文按字母", "顺序不符", "著者-出版年制排序", clause=cl))
    return out

def _set_text(p, text):
    from .features import _set_text as st
    st(p, text)

def fix_references(doc, profile, changes):
    rc = profile["references"]; paras = ref_paras(doc, profile)
    if not paras: return
    parsed = [_split(p.text.strip()) for p in paras]
    numbered = [n is not None for n, _ in parsed]
    for p, (n, body) in zip(paras, parsed):
        nb = fix_entry(body, rc)
        if nb != body: _set_text(p, (f"[{n}] " if n is not None else "") + nb)
    changes.append("参考文献：半角标点、去书名号、页码连接符、作者截断与等/et al、类型标识（可推断时）、末尾句点")
    if all(numbered):
        order = first_order(doc, profile); nums = [n for n, _ in parsed]
        if order is not None and all(n in nums for n in order):
            new_seq = order + [n for n in nums if n not in order]          # 未引用的排在最后（只建议删除）
            remap = {old: i + 1 for i, old in enumerate(new_seq)}
            if any(remap[n] != i + 1 for i, n in enumerate(nums)) or nums != list(range(1, len(nums) + 1)):
                texts = {n: p.text.strip() for p, (n, _) in zip(paras, parsed)}
                for p, old in zip(paras, new_seq):
                    _set_text(p, f"[{remap[old]}] " + _split(texts[old])[1])
                for p, m in reversed(_citations(doc, profile)):
                    ns = _expand(m.group(1)); new = "[" + _compress([remap[x] for x in ns]) + "]"
                    if new != m.group(0): _replace_cite(p, m.group(0), new)
                changes.append(f"参考文献：按正文首次引用顺序重排并重新编号，同步正文引用（映射 {remap}）")
    elif not any(numbered):
        bodies = [p.text.strip() for p in paras]
        s = sorted(bodies, key=_sort_key)
        if s != bodies:
            for p, t in zip(paras, s): _set_text(p, t)
            changes.append("参考文献：著者-出版年制按中文拼音、西文字母排序")

def _replace_cite(p, old, new):
    from .features import _replace_in_runs
    # 用占位避免同段内多次替换互相干扰
    _replace_in_runs(p, old, new.replace("[", "\ue001").replace("]", "\ue002"))

def finalize_cites(doc):
    for p in doc.paragraphs:
        if "\ue001" in p.text:
            for r in core.all_runs(p): r.text = r.text.replace("\ue001", "[").replace("\ue002", "]")
