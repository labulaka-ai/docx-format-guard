import argparse, json, sys
from .core import load_profile, check, fix
from . import revisions as rv
def main(argv=None):
    ap = argparse.ArgumentParser("guard")
    sp = ap.add_subparsers(dest="cmd", required=True)
    c = sp.add_parser("check"); c.add_argument("docx"); c.add_argument("--profile", required=True); c.add_argument("--json", action="store_true")
    f = sp.add_parser("fix"); f.add_argument("docx"); f.add_argument("-o", required=True); f.add_argument("--profile", required=True)
    f.add_argument("--no-toc-pages", action="store_true", help="不调用 LibreOffice 计算目录页码（目录页码留空，打开时提示更新域）")
    pg = sp.add_parser("pages", help="LibreOffice 排版后逐页摘要（页眉/页脚/正文开头），标出空白页与几乎空的页"); pg.add_argument("docx")
    df = sp.add_parser("diff", help="比较改稿前后两份 docx 的段落格式（样式/字体/字号/行距/缩进/对齐）"); df.add_argument("before"); df.add_argument("after"); df.add_argument("--json", action="store_true")
    l = sp.add_parser("review", help="列出全部修订与批注"); l.add_argument("docx"); l.add_argument("--json", action="store_true")
    r = sp.add_parser("reply", help="回复批注 --comment ID，或在修订处加批注 --revision ID")
    r.add_argument("docx"); r.add_argument("-o", required=True); r.add_argument("--text", required=True)
    r.add_argument("--author", default="Author"); g = r.add_mutually_exclusive_group(required=True)
    ig = sp.add_parser("ingest", help="从学校模板 .docx 推导草案规则包（模板里读得到的才写，读不到的在报告里明说）")
    ig.add_argument("docx"); ig.add_argument("-o", required=True); ig.add_argument("--report")
    g.add_argument("--comment"); g.add_argument("--revision")
    a = ap.parse_args(argv)
    if a.cmd == "ingest":
        from .ingest import derive_profile, write_profile, format_report
        prof, rep = derive_profile(a.docx); write_profile(prof, a.o)
        txt = format_report(rep)
        if a.report: open(a.report, "w", encoding="utf-8").write(txt)
        print(txt); print(f"草案规则包 → {a.o}（草案：引用前须按学校指南复核）"); return 0
    if a.cmd == "pages":
        from .layout import summarize
        rows = summarize(a.docx)
        for r in rows:
            flag = "【空白】" if r["blank"] else "【几乎空】" if r["nearly_empty"] else ""
            print(f"第{r['page']:>2}页 页眉「{r['header']}」 页脚「{r['footer']}」 正文{r['lines']}行 「{r['first']}」…「{r['last']}」{flag}")
        from .layout import parity_notes
        for n in parity_notes(rows): print("注：" + n["note"])
        print(f"共 {len(rows)} 页"); return 0
    if a.cmd == "diff":
        from .diff import diff_format
        res = diff_format(a.before, a.after)
        if a.json: print(json.dumps(res, ensure_ascii=False, indent=1)); return 1 if res else 0
        for d in res:
            print(f"段{d['para_b']}「{d['text']}」{'（文字也有改动）' if d['text_changed'] else ''}")
            for x in d["changes"]: print("   ", x)
        print(f"格式变化的段落 {len(res)} 个"); return 1 if res else 0
    if a.cmd == "review":
        revs, coms = rv.list_revisions(a.docx), rv.list_comments(a.docx)
        if a.json: print(json.dumps({"revisions": revs, "comments": coms}, ensure_ascii=False, indent=1)); return 0
        print(f"修订 {len(revs)} 处：")
        for x in revs: print(f"  #{x['id']} [{x['type']}] 段{x['paragraph']} {x['author']} {x['date']}：{x['text'][:40]}" + (f" 旧格式{x['old']}→新{x['new']}" if x['old'] is not None else ""))
        print(f"批注 {len(coms)} 条：")
        for x in coms: print(f"  #{x['id']} {x['author']} {x['date']}{' ↳回复#'+x['reply_to'] if x['reply_to'] else ''} 锚定「{x['anchor_text'][:20]}」：{x['text']}")
        return 0
    if a.cmd == "reply":
        cid = rv.reply_comment(a.docx, a.o, a.comment, a.text, a.author) if a.comment else rv.comment_on_revision(a.docx, a.o, a.revision, a.text, a.author)
        print(f"已写入批注 #{cid} → {a.o}"); return 0
    prof = load_profile(a.profile)
    if a.cmd == "check":
        iss = check(a.docx, prof)
        if a.json: print(json.dumps(iss, ensure_ascii=False, indent=1))
        else:
            for i in iss: print(f"[{i['kind']}] {i['where']}: 期望 {i['expect']}，实际 {i['actual']}（{i['cause']}）")
            print(f"共 {len(iss)} 处问题")
        return 1 if iss else 0
    import shutil
    upd = not a.no_toc_pages and shutil.which("soffice") is not None
    if not a.no_toc_pages and not upd: print("提示：未找到 soffice，目录页码未计算，打开文档后需更新域")
    for ch in fix(a.docx, a.o, prof, update_toc=upd): print("-", ch)
    left = check(a.o, prof); print(f"修复后剩余问题 {len(left)} 处"); return 1 if left else 0
if __name__ == "__main__": sys.exit(main())
