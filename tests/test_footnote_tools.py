from __future__ import annotations

import asyncio
from pathlib import Path

from tests.helpers import write_docx
from word_document_server.tools.footnote_tools import add_note


def test_add_note_is_disabled(tmp_path: Path):
    p = tmp_path / "x.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(add_note(filename=str(p), paragraph_index=0, note_text="note"))
    assert "footnotes/endnotes are not supported by python-docx" in res.lower()
