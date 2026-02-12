from __future__ import annotations

import asyncio
from pathlib import Path

from docx import Document

from word_document_server.tools.document_tools import create_document


def test_create_document_creates_docx_and_sets_properties(tmp_path: Path):
    p = tmp_path / "new_doc"
    res = asyncio.run(create_document(filename=str(p), title="T", author="A"))
    assert "created successfully" in res.lower()

    created = tmp_path / "new_doc.docx"
    assert created.exists()

    doc = Document(str(created))
    assert doc.core_properties.title == "T"
    assert doc.core_properties.author == "A"


def test_create_document_rejects_doc_extension(tmp_path: Path):
    p = tmp_path / "bad.doc"
    res = asyncio.run(create_document(filename=str(p)))
    assert ".doc files are not supported" in res.lower()

