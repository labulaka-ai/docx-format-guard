import sys, pathlib
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests"))
import pytest
from guard.core import load_profile
@pytest.fixture(scope="session")
def profile(): return load_profile(str(ROOT / "profiles/example-master.yaml"))
@pytest.fixture(scope="session")
def golden(tmp_path_factory, profile):
    from fixtures import build
    from guard.core import fix
    d = tmp_path_factory.mktemp("golden"); raw = build(str(d / "raw.docx"))
    out = str(d / "golden.docx"); fix(raw, out, profile); return out
