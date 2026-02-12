from __future__ import annotations

import json
import os
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

import anyio
import pytest
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


async def _with_session(run):
    root = _project_root()
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "word_document_server.main"],
        cwd=str(root),
        env={**os.environ, "PYTHONUNBUFFERED": "1"},
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

