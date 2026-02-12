from __future__ import annotations

from pathlib import Path

import pytest

from tests.helpers import write_docx
from word_document_server.tools.footnote_tools import add_note


@pytest.mark.asyncio
async def test_add_note_is_disabled(tmp_path: Path):
    p = tmp_path / "x.docx"
    write_docx(p, paragraphs=["Hello"])

    res = await add_note(filename=str(p), paragraph_index=0, note_text="note")
    assert "not supported" in res.lower() or "disabled" in res.lower()
