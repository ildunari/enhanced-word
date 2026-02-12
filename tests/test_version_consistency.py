from __future__ import annotations

import json
from pathlib import Path
import tomllib


def test_pyproject_version_matches_package_json():
    root = Path(__file__).resolve().parents[1]

    pkg = json.loads((root / "package.json").read_text(encoding="utf-8"))
    py = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))

    assert py["project"]["version"] == pkg["version"]
