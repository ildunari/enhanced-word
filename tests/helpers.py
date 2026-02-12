from __future__ import annotations

from pathlib import Path

from docx import Document


def write_docx(path: Path, paragraphs: list[str] | None = None) -> Path:
    doc = Document()
    for p in paragraphs or []:
        doc.add_paragraph(p)
    doc.save(path)
    return path

