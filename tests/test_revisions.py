import shutil, subprocess, pytest
from docx import Document
from docx.oxml.ns import qn
from fixtures import build_reviewed, build, CORRUPTIONS
from wordxml import validate
from guard.core import check, fix, eff_size, eff_font, all_runs
from guard import revisions as rv

@pytest.fixture
def reviewed(tmp_path): return build_reviewed(str(tmp_path / "reviewed.docx"))

def test_fixture_is_valid(reviewed):
    assert validate(reviewed) == []

def test_list_revisions(reviewed):
    revs = rv.list_revisions(reviewed)
    by = {r["id"]: r for r in revs}
    assert set(by) == {"101", "102", "103", "104", "105"}
    assert by["101"]["type"] == "插入" and "导师补充" in by["101"]["text"]
    assert by["102"]["type"] == "删除" and by["102"]["text"] == "平均精度提升 2.5%。"
    assert by["103"]["type"] == "字符格式修订" and "b" in by["103"]["new"]
    assert by["104"]["type"] == "段落格式修订" and "jc" in by["104"]["old"]
    assert by["105"]["type"] == "段落标记插入"
    assert all(r["author"] == "导师" and r["date"] == "2026-10-16T08:00:00Z" for r in revs)
    assert all(r["paragraph"] is not None for r in revs)

def test_list_comments(reviewed):
    cs = rv.list_comments(reviewed)
    assert len(cs) == 1 and cs[0]["author"] == "导师" and cs[0]["text"] == "这段缺少引用。"
    assert cs[0]["anchor_text"].startswith("深度学习") and cs[0]["reply_to"] is None

def test_reply_comment(tmp_path, reviewed):
    out = str(tmp_path / "replied.docx")
    cid = rv.reply_comment(reviewed, out, "0", "已补充引用 [12]。", author="学生")
    assert validate(out) == []
    cs = {c["id"]: c for c in rv.list_comments(out)}
    assert cs[cid]["reply_to"] == "0" and cs[cid]["author"] == "学生" and cs[cid]["text"] == "已补充引用 [12]。"
    out2 = str(tmp_path / "replied2.docx")                    # 回复的回复 / 连续回复
    cid2 = rv.reply_comment(out, out2, "0", "第二条回复", author="学生")
    assert validate(out2) == []
    cs = {c["id"]: c for c in rv.list_comments(out2)}
    assert cs[cid2]["reply_to"] == "0" and len(cs) == 3
    assert len(rv.list_revisions(out2)) == 5                  # 不动修订

@pytest.mark.parametrize("rev_id,expect", [("101", "导师补充"), ("102", "平均精度"), ("103", ""), ("104", ""), ("105", "")])
def test_comment_on_revision(tmp_path, reviewed, rev_id, expect):
    out = str(tmp_path / "c.docx")
    cid = rv.comment_on_revision(reviewed, out, rev_id, f"回应修订 {rev_id}")
    assert validate(out) == []
    c = {c["id"]: c for c in rv.list_comments(out)}[cid]
    assert c["text"] == f"回应修订 {rev_id}" and expect in c["anchor_text"] and c["anchor_text"]
    rev_para = {r["id"]: r["paragraph"] for r in rv.list_revisions(out)}[rev_id]
    assert c["paragraph"] == rev_para

def test_fix_formats_inserted_text(tmp_path, profile, reviewed):
    """上次的已知局限：修订插入文字自带 15pt，现在必须被检出并修复。"""
    kinds = {(i["kind"], i["actual"]) for i in check(reviewed, profile)}
    assert ("size", 15.0) in kinds
    out = str(tmp_path / "fixed.docx"); fix(reviewed, out, profile)
    assert check(out, profile) == []
    d = Document(out)
    ins = d.element.body.find(".//" + qn("w:ins"))
    p = next(p for p in d.paragraphs if p._p is ins.getparent())
    run = next(r for r in all_runs(p) if "导师补充" in r.text)
    assert eff_size(d, p, run) == 12 and eff_font(d, p, run, "eastAsia") == "宋体"
    assert validate(out) == []
    assert {r["id"] for r in rv.list_revisions(out)} == {"101", "102", "103", "104", "105"}  # 修订全部保留
    assert len(rv.list_comments(out)) == 1                                                 # 批注保留

def test_full_pipeline_valid(tmp_path, profile):
    src = build(str(tmp_path / "in.docx"), sorted(CORRUPTIONS), revision=True)
    out = str(tmp_path / "out.docx"); fix(src, out, profile)
    assert validate(out) == [] and check(out, profile) == []

@pytest.mark.skipif(shutil.which("soffice") is None, reason="需要 LibreOffice")
def test_libreoffice_opens_and_keeps_annotations(tmp_path, profile, reviewed):
    out = str(tmp_path / "final.docx")
    rv.reply_comment(reviewed, str(tmp_path / "a.docx"), "0", "已补充", author="学生")
    rv.comment_on_revision(str(tmp_path / "a.docx"), str(tmp_path / "b.docx"), "101", "同意")
    fix(str(tmp_path / "b.docx"), out, profile)
    lo = tmp_path / "lo"; lo.mkdir()
    for fmt in ("pdf", "docx"):
        r = subprocess.run(["soffice", "--headless", "--convert-to", fmt, "--outdir", str(lo), out], capture_output=True, timeout=120)
        assert r.returncode == 0 and (lo / f"final.{fmt}").exists(), r.stderr
    rt = str(lo / "final.docx")                     # LibreOffice 读入再存出，批注与修订仍在
    assert len(rv.list_comments(rt)) == 3
    assert {"插入", "删除"} <= {r["type"] for r in rv.list_revisions(rt)}
