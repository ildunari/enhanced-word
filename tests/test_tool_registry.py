from __future__ import annotations

import re
from pathlib import Path


def _registered_tool_names_from_main_py(repo_root: Path) -> list[str]:
    main_py = repo_root / "word_document_server" / "main.py"
    text = main_py.read_text(encoding="utf-8")

    names: list[str] = []
    pattern = re.compile(
        r"^(?!\s*#).*mcp\.tool\(\)\(\s*(?:\w+\.)?(\w+)\s*\)",
        re.MULTILINE,
    )
    for m in pattern.finditer(text):
        names.append(m.group(1))
    return names


def test_registered_tool_list_matches_main_registration():
    import word_document_server.tools as tools

    repo_root = Path(__file__).resolve().parents[1]
    from_main = _registered_tool_names_from_main_py(repo_root)

    assert hasattr(tools, "REGISTERED_TOOL_NAMES"), "tools.REGISTERED_TOOL_NAMES is missing"
    assert sorted(from_main) == sorted(tools.REGISTERED_TOOL_NAMES)
    assert tools.REGISTERED_TOOL_COUNT == len(tools.REGISTERED_TOOL_NAMES)
