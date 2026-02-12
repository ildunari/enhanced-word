from __future__ import annotations

import asyncio
from pathlib import Path

from docx import Document

from tests.helpers import write_docx
from word_document_server.tools.protection_tools import add_digital_signature, verify_document


def test_digital_signature_verifies_until_document_changes(tmp_path: Path):
    p = tmp_path / "signed.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(add_digital_signature(filename=str(p), signer_name="Alice", reason="test"))
    assert "digital signature added" in res.lower()

    verify1 = asyncio.run(verify_document(filename=str(p)))
    assert "signature is valid" in verify1.lower()

    # Mutate document content after signing.
    doc = Document(str(p))
    doc.add_paragraph("CHANGED")
    doc.save(str(p))

    verify2 = asyncio.run(verify_document(filename=str(p)))
    assert "modified since it was signed" in verify2.lower()

