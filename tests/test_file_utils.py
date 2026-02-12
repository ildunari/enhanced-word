from __future__ import annotations

from word_document_server.utils.file_utils import validate_docx_path


def test_validate_docx_path_rejects_doc_extension():
    ok, _, err = validate_docx_path("example.doc")
    assert not ok
    assert ".doc" in err.lower()


def test_validate_docx_path_adds_docx_extension_when_missing():
    ok, sanitized, err = validate_docx_path("example")
    assert ok, err
    assert sanitized.endswith(".docx")


def test_validate_docx_path_accepts_docx_extension():
    ok, sanitized, err = validate_docx_path("example.docx")
    assert ok, err
    assert sanitized.endswith(".docx")

