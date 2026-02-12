from __future__ import annotations

import asyncio
import builtins
from pathlib import Path

from tests.helpers import write_docx
from word_document_server.tools.extended_document_tools import convert_to_pdf


class _RunResult:
    def __init__(self, returncode: int = 0, stderr: str = "", stdout: str = ""):
        self.returncode = returncode
        self.stderr = stderr
        self.stdout = stdout


def test_convert_to_pdf_errors_if_subprocess_reports_success_but_no_pdf_created(tmp_path: Path, monkeypatch):
    docx_path = tmp_path / "in.docx"
    write_docx(docx_path, paragraphs=["Hello"])

    out_pdf = tmp_path / "out.pdf"

    import word_document_server.tools.extended_document_tools as mod

    orig_import = builtins.__import__
    def _block_docx2pdf(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "docx2pdf":
            raise ImportError("blocked for test")
        return orig_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(mod.platform, "system", lambda: "Linux")
    monkeypatch.setattr(mod.subprocess, "run", lambda *args, **kwargs: _RunResult(returncode=0))
    monkeypatch.setattr(builtins, "__import__", _block_docx2pdf)

    res = asyncio.run(convert_to_pdf(filename=str(docx_path), output_filename=str(out_pdf)))
    assert "output pdf not found" in res.lower()
