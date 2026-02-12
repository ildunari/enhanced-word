from __future__ import annotations

from pathlib import Path

from docx import Document

from word_document_server.tools.review_tools import manage_comments


def _write_doc_with_paragraph(path: Path, text: str) -> None:
    doc = Document()
    doc.add_paragraph(text)
    doc.save(path)


def test_manage_comments_list_finds_legacy_marker(tmp_path: Path):
    p = tmp_path / "comments.docx"
    _write_doc_with_paragraph(p, "Hello [COMMENT-deadbeef by Alice: check this].")

    res = manage_comments(filename=str(p), action="list")
    assert "Found 1 comments" in res
    assert "deadbeef" in res.lower()
    assert "Alice" in res


def test_manage_comments_non_list_is_unsupported(tmp_path: Path):
    p = tmp_path / "comments.docx"
    _write_doc_with_paragraph(p, "Hello.")

    res = manage_comments(filename=str(p), action="add", paragraph_index=0, comment_text="x")
    assert "only" in res.lower()
    assert "list" in res.lower()

