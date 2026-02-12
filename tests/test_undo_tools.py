from __future__ import annotations

import asyncio
from pathlib import Path

from docx import Document

from tests.helpers import write_docx
from word_document_server.tools.session_tools import open_document
from word_document_server.tools.content_tools import add_text_content
from word_document_server.tools.undo_tools import session_undo


def _read_bytes(path: Path) -> bytes:
    return path.read_bytes()


def test_session_undo_uses_active_document_when_no_path_provided(tmp_path: Path):
    p = tmp_path / "x.docx"
    write_docx(p, paragraphs=["Hello"])

    res = open_document("main", str(p))
    assert "successfully opened" in res.lower()

    before = _read_bytes(p)
    asyncio.run(add_text_content(document_id="main", text="More text"))
    after = _read_bytes(p)
    assert after != before

    # No document_id/filename provided: should target active document ("main").
    undo_res = session_undo(action="undo", steps=1)
    assert "undo successful" in undo_res.lower() or "undo" in undo_res.lower()

    restored = _read_bytes(p)
    assert restored == before

