from __future__ import annotations

import json
import os
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

import anyio
import pytest
from docx import Document
from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


pytestmark = pytest.mark.e2e


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _call_result_text(result) -> str:
    texts: list[str] = []
    for item in getattr(result, "content", []):
        text = getattr(item, "text", None)
        if isinstance(text, str):
            texts.append(text)
    return "\n".join(texts)


async def _with_session(run, extra_env: dict[str, str] | None = None):
    root = _project_root()
    merged_env = {**os.environ, "PYTHONUNBUFFERED": "1"}
    if extra_env:
        merged_env.update(extra_env)
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "word_document_server.main"],
        cwd=str(root),
        env=merged_env,
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            return await run(session)


def test_e2e_stdio_lists_registered_tools():
    async def _run(session: ClientSession):
        result = await session.list_tools()
        names = sorted([tool.name for tool in result.tools])
        assert len(names) == 25
        for expected in [
            "create_document",
            "add_text_content",
            "enhanced_search_and_replace",
            "get_text",
            "session_manager",
            "session_undo",
        ]:
            assert expected in names

    anyio.run(_run_wrapper, _run)


def test_e2e_stdio_core_tool_calls_round_trip():
    async def _run(session: ClientSession):
        with TemporaryDirectory(prefix="ew-e2e-") as tmp:
            doc_path = str(Path(tmp) / "roundtrip.docx")

            created = await session.call_tool("create_document", {"filename": doc_path})
            assert "created successfully" in _call_result_text(created).lower()

            added = await session.call_tool(
                "add_text_content",
                {
                    "filename": doc_path,
                    "text": "hello world",
                    "content_type": "paragraph",
                },
            )
            assert "added" in _call_result_text(added).lower()

            replaced = await session.call_tool(
                "enhanced_search_and_replace",
                {"filename": doc_path, "find_text": "world", "replace_text": "mcp"},
            )
            assert "replaced" in _call_result_text(replaced).lower()

            extracted = await session.call_tool(
                "get_text",
                {"filename": doc_path, "scope": "search", "search_term": "mcp"},
            )
            payload = json.loads(_call_result_text(extracted))
            assert payload.get("returned_count", payload.get("total_count", 0)) >= 1

            sessions = await session.call_tool("session_manager", {"action": "list"})
            assert isinstance(_call_result_text(sessions), str)

    anyio.run(_run_wrapper, _run)


async def _run_wrapper(coro):
    await _with_session(coro)


async def _run_wrapper_with_env(coro, extra_env: dict[str, str]):
    await _with_session(coro, extra_env=extra_env)


def _write_large_doc(path: Path, paragraphs: int = 80, width: int = 280) -> None:
    doc = Document()
    for i in range(paragraphs):
        doc.add_paragraph(f"{i:04d} " + ("x" * width))
    doc.save(str(path))


def test_e2e_stdio_guardrail_error_propagates_cleanly():
    async def _run(session: ClientSession):
        with TemporaryDirectory(prefix="ew-e2e-guard-") as tmp:
            doc_path = Path(tmp) / "large.docx"
            _write_large_doc(doc_path)

            extracted = await session.call_tool(
                "get_text",
                {"filename": str(doc_path), "scope": "search", "search_term": "x"},
            )
            text = _call_result_text(extracted)
            assert "DOC_TOO_LARGE" in text

    anyio.run(
        _run_wrapper_with_env,
        _run,
        {"EW_MAX_DOC_BYTES_PER_OPERATION": "5000"},
    )


def test_e2e_stdio_long_session_smoke_chain():
    async def _run(session: ClientSession):
        with TemporaryDirectory(prefix="ew-e2e-long-") as tmp:
            doc_path = str(Path(tmp) / "long.docx")

            created = await session.call_tool("create_document", {"filename": doc_path})
            assert "created successfully" in _call_result_text(created).lower()

            opened = await session.call_tool(
                "session_manager",
                {"action": "open", "document_id": "main", "file_path": doc_path},
            )
            assert "successfully opened" in _call_result_text(opened).lower()

            for i in range(25):
                added = await session.call_tool(
                    "add_text_content",
                    {
                        "document_id": "main",
                        "text": f"line-{i}",
                        "content_type": "paragraph",
                    },
                )
                assert "added" in _call_result_text(added).lower()

                replaced = await session.call_tool(
                    "enhanced_search_and_replace",
                    {
                        "document_id": "main",
                        "find_text": "line",
                        "replace_text": "line",
                        "whole_words_only": True,
                    },
                )
                assert "replaced" in _call_result_text(replaced).lower() or "no occurrences" in _call_result_text(replaced).lower()

            listed = await session.call_tool("session_undo", {"action": "list", "document_id": "main"})
            assert "history for" in _call_result_text(listed).lower()

    anyio.run(_run_wrapper, _run)
