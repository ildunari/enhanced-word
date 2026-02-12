from __future__ import annotations

import asyncio
import json
from pathlib import Path

from tests.helpers import write_docx
from word_document_server.tools.citation_tools import citations


def test_citations_list_on_plain_doc_returns_empty_json(tmp_path: Path):
    p = tmp_path / "c.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(citations(action="list", filename=str(p)))
    data = json.loads(res)
    assert data["total_citations"] == 0


def test_citations_rejects_invalid_action(tmp_path: Path):
    p = tmp_path / "c.docx"
    write_docx(p, paragraphs=["Hello"])

    res = asyncio.run(citations(action="nope", filename=str(p)))
    assert "invalid action" in res.lower()

