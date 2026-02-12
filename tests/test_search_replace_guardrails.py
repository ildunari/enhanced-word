from __future__ import annotations

import asyncio
import json
from pathlib import Path

from docx import Document

from word_document_server.tools.content_tools import add_text_content, enhanced_search_and_replace
from word_document_server.tools.document_tools import create_document, get_text


def test_guardrail_max_matches_per_call_limits_mutation(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("EW_MAX_MATCHES_PER_CALL", "2")

    p = tmp_path / "guard-limit.docx"
    asyncio.run(create_document(filename=str(p)))
    asyncio.run(add_text_content(filename=str(p), text="x x x x"))

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text="x",
        replace_text="y",
        whole_words_only=True,
    )

    assert "Replacement limit reached" in result

    doc = Document(str(p))
    text = "\n".join(par.text for par in doc.paragraphs)
    assert text.count("y") == 2
    assert text.count("x") >= 2


def test_guardrail_regex_pattern_length(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("EW_MAX_REGEX_PATTERN_CHARS", "3")

    p = tmp_path / "guard-pattern.docx"
    asyncio.run(create_document(filename=str(p)))
    asyncio.run(add_text_content(filename=str(p), text="hello world"))

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text=r"(hello)",
        replace_text="x",
        use_regex=True,
    )
    assert "Regex pattern too long" in result


def test_guardrail_regex_scan_size(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("EW_MAX_REGEX_SCAN_CHARS", "20")

    p = tmp_path / "guard-scan.docx"
    asyncio.run(create_document(filename=str(p)))
    asyncio.run(add_text_content(filename=str(p), text="a" * 200))

    result = enhanced_search_and_replace(
        filename=str(p),
        find_text=r"a+",
        replace_text="z",
        use_regex=True,
    )
    assert "Regex search refused due to scan-size guardrail" in result


def test_get_text_search_output_is_char_capped(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("EW_MAX_SEARCH_OUTPUT_CHARS", "700")

    p = tmp_path / "guard-output.docx"
    asyncio.run(create_document(filename=str(p)))
    for _ in range(20):
        asyncio.run(add_text_content(filename=str(p), text="needle " + ("x" * 80)))

    raw = asyncio.run(
        get_text(
            filename=str(p),
            scope="search",
            search_term="needle",
            include_formatting=True,
            max_results=200,
            match_case=True,
            whole_word=False,
        )
    )

    assert len(raw) <= 700
    payload = json.loads(raw)
    assert payload.get("truncated") is True
    assert payload.get("output_truncated") is True
    assert payload.get("returned_count", 0) <= payload.get("total_count", 0)

