"""
Footnote and endnote tools for Word Document Server.

python-docx does not support creating native Word footnotes/endnotes safely, so this
tool is disabled to avoid producing corrupt documents.
"""

from __future__ import annotations

from typing import Optional

from word_document_server.utils.session_utils import resolve_document_path


async def add_note(
    document_id: str = None,
    filename: str = None,
    paragraph_index: int = None,
    note_text: str = None,
    note_type: str = "footnote",
    position: str = "end",
    symbol: Optional[str] = None,
) -> str:
    # Keep session/path resolution behavior consistent across tools.
    _, error_msg = resolve_document_path(document_id, filename)
    if error_msg:
        return error_msg

    return (
        "Footnotes/endnotes are not supported by python-docx. "
        "This tool is currently disabled to avoid corrupting document structure."
    )


__all__ = ["add_note"]

