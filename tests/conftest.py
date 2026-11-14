from __future__ import annotations

import sys
import shutil
import tempfile
from pathlib import Path

import pytest


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


@pytest.fixture
def work_tmp():
    base = Path(__file__).resolve().parent / ".tmp"
    base.mkdir(parents=True, exist_ok=True)
    path = Path(tempfile.mkdtemp(prefix="case_", dir=base))
    try:
        yield path
    finally:
        shutil.rmtree(path, ignore_errors=True)
