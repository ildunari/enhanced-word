from __future__ import annotations

import asyncio
from pathlib import Path

from tests.helpers import write_docx
from word_document_server.tools.equation_tools import insert_equation


def test_insert_equation_requires_latex(tmp_path: Path):
    p = tmp_path / "e.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(insert_equation(filename=str(p), latex=""))
    assert "latex" in res.lower()


def test_insert_equation_rejects_invalid_position(tmp_path: Path):
    p = tmp_path / "e.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(insert_equation(filename=str(p), latex="x", position="nope"))
    assert "invalid position" in res.lower()


def test_insert_equation_propagates_conversion_failure(tmp_path: Path, monkeypatch):
    p = tmp_path / "e.docx"
    write_docx(p, paragraphs=["Hello"])

    import word_document_server.tools.equation_tools as mod
    monkeypatch.setattr(mod, "latex_to_omml", lambda _latex: (False, "boom"))

    res = asyncio.run(mod.insert_equation(filename=str(p), latex="x"))
    assert res == "boom"

