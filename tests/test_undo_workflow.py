from __future__ import annotations

import asyncio
import base64
from pathlib import Path

from tests.helpers import write_docx
from word_document_server.tools.content_tools import add_picture, add_table, enhanced_search_and_replace
from word_document_server.tools.session_tools import open_document
from word_document_server.tools.undo_tools import session_undo


def _read_bytes(path: Path) -> bytes:
    return path.read_bytes()


def test_undo_redo_after_search_replace(tmp_path: Path):
    p = tmp_path / "sr.docx"
    write_docx(p, paragraphs=["alpha beta"])

    assert "successfully opened" in open_document("doc", str(p)).lower()

    before = _read_bytes(p)
    result = enhanced_search_and_replace(document_id="doc", find_text="beta", replace_text="gamma")
    assert "replaced" in result.lower()
    changed = _read_bytes(p)
    assert changed != before

    undo_res = session_undo(action="undo", steps=1)
    assert "undo" in undo_res.lower()
    assert _read_bytes(p) == before

    redo_res = session_undo(action="redo", steps=1)
    assert "redo" in redo_res.lower()
    assert _read_bytes(p) == changed


def test_undo_after_add_table(tmp_path: Path):
    p = tmp_path / "table.docx"
    write_docx(p, paragraphs=["one"])

    assert "successfully opened" in open_document("table_doc", str(p)).lower()
    before = _read_bytes(p)

    add_res = asyncio.run(add_table(document_id="table_doc", rows=2, cols=2, data=[["a", "b"], ["c", "d"]]))
    assert "added" in add_res.lower()
    changed = _read_bytes(p)
    assert changed != before

    undo_res = session_undo(action="undo", steps=1)
    assert "undo" in undo_res.lower()
    assert _read_bytes(p) == before


def test_undo_after_add_picture(tmp_path: Path):
    p = tmp_path / "pic.docx"
    write_docx(p, paragraphs=["image"])

    png_path = tmp_path / "dot.png"
    png_path.write_bytes(
        base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6p1U8AAAAASUVORK5CYII=")
    )

    assert "successfully opened" in open_document("pic_doc", str(p)).lower()
    before = _read_bytes(p)

    add_res = asyncio.run(add_picture(document_id="pic_doc", image_path=str(png_path)))
    assert "added" in add_res.lower()
    changed = _read_bytes(p)
    assert changed != before

    undo_res = session_undo(action="undo", steps=1)
    assert "undo" in undo_res.lower()
    assert _read_bytes(p) == before

