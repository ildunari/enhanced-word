from __future__ import annotations

import json
from pathlib import Path
import tomllib

import word_document_server


def test_pyproject_version_matches_package_json():
    root = Path(__file__).resolve().parents[1]

    pkg = json.loads((root / "package.json").read_text(encoding="utf-8"))
    py = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))

    assert py["project"]["version"] == pkg["version"]


def test_module_version_matches_package_metadata():
    root = Path(__file__).resolve().parents[1]
    pkg = json.loads((root / "package.json").read_text(encoding="utf-8"))
    assert word_document_server.__version__ == pkg["version"]
