"""Word 兼容性结构校验（自写规则，不依赖第三方 XSD 文件）。返回问题列表，空表示通过。"""
import zipfile, posixpath
from lxml import etree
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
def q(ns, t): return f"{{{ns}}}{t}"

def validate(path):
    errs = []; z = zipfile.ZipFile(path); names = set(z.namelist()); xml = {}
    for n in names:
        if n.endswith(".xml") or n.endswith(".rels"):
            try: xml[n] = etree.fromstring(z.read(n))
            except etree.XMLSyntaxError as e: errs.append(f"{n} 非良构 XML：{e}")
    if errs: return errs
    ct = xml["[Content_Types].xml"]
    overrides = {o.get("PartName").lstrip("/") for o in ct.iter(q(CT, "Override"))}
    defaults = {d.get("Extension") for d in ct.iter(q(CT, "Default"))}
    for n in names:
        if n.startswith("word/") and n.endswith(".xml") and n not in overrides:
            errs.append(f"{n} 缺 Content_Types Override")
        ext = n.rsplit(".", 1)[-1]
        if not n.endswith("/") and n not in overrides and n.lstrip("/") != "[Content_Types].xml" and ext not in defaults and not n.startswith("word/"):
            errs.append(f"{n} 无内容类型")
    rels = xml.get("word/_rels/document.xml.rels")
    targets = {}
    for rel in rels.iter(q(R, "Relationship")):
        if rel.get("TargetMode") == "External": continue
        t = posixpath.normpath(posixpath.join("word", rel.get("Target")))
        targets[rel.get("Type").rsplit("/", 1)[-1]] = t
        if t not in names: errs.append(f"关系目标不存在 {t}")
    doc = xml["word/document.xml"]
    # 修订：属性齐全、id 唯一、删除内用 delText
    ids = []
    for tag in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange"):
        for e in doc.iter(q(W, tag)):
            for a in ("id", "author"):
                if e.get(q(W, a)) is None: errs.append(f"w:{tag} 缺 w:{a}")
            ids.append(e.get(q(W, "id")))
            if tag in ("del", "moveFrom") and list(e.iter(q(W, "t"))): errs.append(f"w:{tag} 内出现 w:t（应为 w:delText）")
    # rPrChange 必须是 rPr 的最后一个子元素；pPrChange 必须是 pPr 最后一个
    for tag, parent in (("rPrChange", "rPr"), ("pPrChange", "pPr")):
        for e in doc.iter(q(W, tag)):
            if e.getparent()[-1] is not e: errs.append(f"w:{tag} 不是 w:{parent} 的最后一个子元素")
    # 批注
    com = xml.get("word/comments.xml")
    cids = [c.get(q(W, "id")) for c in com.iter(q(W, "comment"))] if com is not None else []
    if len(cids) != len(set(cids)): errs.append("批注 id 重复")
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        got = [e.get(q(W, "id")) for e in doc.iter(q(W, tag))]
        for i in got:
            if i not in cids: errs.append(f"w:{tag} id={i} 无对应批注")
        if tag != "commentReference" and len(got) != len(set(got)): errs.append(f"w:{tag} id 重复")
        if set(got) != set(cids): errs.append(f"w:{tag} 与批注集合不一致 {sorted(got)} vs {sorted(cids)}")
    order = [(e.tag.rsplit('}')[1], e.get(q(W, "id"))) for e in doc.iter(q(W, "commentRangeStart"), q(W, "commentRangeEnd"))]
    seen = set()
    for t, i in order:
        if t == "commentRangeStart": seen.add(i)
        elif i not in seen: errs.append(f"批注 {i} 的 RangeEnd 出现在 RangeStart 之前")
    for c in (com.iter(q(W, "comment")) if com is not None else []):
        if c.find(q(W, "p")) is None: errs.append(f"批注 {c.get(q(W,'id'))} 无段落")
    ext = xml.get("word/commentsExtended.xml")
    if ext is not None:
        if "commentsExtended" not in targets: errs.append("commentsExtended.xml 未在 document.xml.rels 中关联")
        pids = {p.get(q(W14, "paraId")) for p in com.iter(q(W, "p"))}
        for e in ext.iter(q(W15, "commentEx")):
            if e.get(q(W15, "paraId")) not in pids: errs.append(f"commentEx paraId {e.get(q(W15,'paraId'))} 不存在")
            par = e.get(q(W15, "paraIdParent"))
            if par and par not in pids: errs.append(f"paraIdParent {par} 不存在")
        for pid in pids:
            if pid and (len(pid) != 8 or int(pid, 16) >= 0x80000000): errs.append(f"paraId {pid} 超出 Word 允许范围")
    return errs
