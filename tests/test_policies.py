from pathlib import Path
import pytest
from tales.core.policies import PolicyViolation, ensure_within_roots


def test_allowed_path(tmp_path: Path):
    root = tmp_path / "allowed"
    root.mkdir()
    target = root / "file.txt"
    target.write_text("ok")
    assert ensure_within_roots(target, [root]) == target.resolve()


def test_denied_path(tmp_path: Path):
    root = tmp_path / "allowed"
    root.mkdir()
    outside = tmp_path / "outside.txt"
    outside.write_text("no")
    with pytest.raises(PolicyViolation):
        ensure_within_roots(outside, [root])
