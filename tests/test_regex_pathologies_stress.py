from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

import word_document_server.tools.content_tools as content_tools
from word_document_server.tools.content_tools import enhanced_search_and_replace


pytestmark = pytest.mark.stress


def _write_doc(path: Path, text: str) -> None:
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(path))


def test_stress_nested_quantifier_regex_is_refused_without_mutation(tmp_path: Path):
    p = tmp_path / "nested.docx"
    _write_doc(p, ("a" * 5000) + "X")
    before = p.read_bytes()

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text=r"(a+)+$",
        replace_text="ok",
        use_regex=True,
    )

    assert "REGEX_COMPLEXITY_BLOCKED" in result
    assert "nested quantifiers" in result
    assert p.read_bytes() == before


def test_stress_valid_complex_regex_still_executes(tmp_path: Path):
    p = tmp_path / "valid.docx"
    _write_doc(p, "abcabc abc")

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text=r"(abc){2}",
        replace_text="Z",
        use_regex=True,
    )
    assert "Replaced" in result


def test_stress_timeout_contract_returns_clear_message(tmp_path: Path, monkeypatch):
    p = tmp_path / "timeout.docx"
    _write_doc(p, "aaaaa")

    monkeypatch.setenv("EW_REGEX_TIMEOUT_MS", "10")
    monkeypatch.setattr(content_tools, "is_regex_timeout_supported", lambda: False)

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text=r"a+",
        replace_text="b",
        use_regex=True,
    )

    assert "REGEX_COMPLEXITY_BLOCKED" in result
    assert "EW_REGEX_TIMEOUT_MS=10" in result
