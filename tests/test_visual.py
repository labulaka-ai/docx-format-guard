"""视觉回归：LibreOffice 转 PDF，修复稿与 golden 逐页像素比对。"""
import shutil, subprocess, pytest, pymupdf
from fixtures import build, CORRUPTIONS
from guard.core import fix
pytestmark = pytest.mark.skipif(shutil.which("soffice") is None, reason="需要 LibreOffice")

def to_pdf(docx, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), docx],
                   check=True, capture_output=True, timeout=120)
    return str(outdir / (docx.rsplit("/", 1)[-1][:-5] + ".pdf"))

def diff_ratio(a, b):
    da, db = pymupdf.open(a), pymupdf.open(b)
    assert len(da) == len(db), "页数不同"
    worst = 0.0
    for pa, pb in zip(da, db):
        ia, ib = pa.get_pixmap(dpi=60, colorspace=pymupdf.csGRAY), pb.get_pixmap(dpi=60, colorspace=pymupdf.csGRAY)
        assert (ia.width, ia.height) == (ib.width, ib.height)
        diff = sum(1 for x, y in zip(ia.samples, ib.samples) if abs(x - y) > 40)
        worst = max(worst, diff / len(ia.samples))
    return worst

def test_visual_fixed_matches_golden(tmp_path, golden, profile):
    src = build(str(tmp_path / "broken.docx"), sorted(CORRUPTIONS)); out = str(tmp_path / "fixed.docx")
    fix(src, out, profile)
    g, f, b = to_pdf(golden, tmp_path), to_pdf(out, tmp_path), to_pdf(src, tmp_path)
    assert diff_ratio(g, f) < 0.001, "修复稿与 golden 视觉差异过大"
    assert diff_ratio(g, b) > 0.005, "对照：乱掉的稿子应与 golden 明显不同（证明比对有效）"
