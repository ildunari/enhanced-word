from __future__ import annotations

from pathlib import Path

from word_document_server.utils.session_utils import resolve_document_path
from word_document_server.session_manager import get_session_manager
from tests.helpers import write_docx


def test_resolve_document_path_rejects_doc_extension():
    path, err = resolve_document_path(filename="example.doc")
    assert path == ""
    assert ".doc" in err.lower()


def test_resolve_document_path_adds_docx_extension_when_missing(tmp_path: Path):
    file_path = tmp_path / "x.docx"
    write_docx(file_path, paragraphs=["hi"])

    resolved, err = resolve_document_path(filename=str(file_path.with_suffix("")))
    assert err == ""
    assert resolved.endswith(".docx")


def test_session_manager_open_document_rejects_doc_extension(tmp_path: Path):
    mgr = get_session_manager()
    res = mgr.open_document("doc1", str(tmp_path / "x.doc"))
    assert "not supported" in res.lower()
