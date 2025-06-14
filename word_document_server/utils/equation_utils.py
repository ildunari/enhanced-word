"""Utility helpers for converting LaTeX math to Office Math (OMML) XML.

Conversion pipeline preferred:
    LaTeX  --pandoc-->  MathML  --mml2omml.xsl-->  OMML

If either *pandoc* or *xsltproc* is not available at runtime, the helpers
return an error message so the calling tool can inform the user.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Tuple

MML2OMML_XSL_URL = (
    "https://raw.githubusercontent.com/plgonzalezrx/mathml2omml/master/mml2omml.xsl"
)


def _ensure_xsl_cached() -> Path | None:
    """Download *mml2omml.xsl* into a cache dir the first time it is needed."""
    cache_dir = Path(tempfile.gettempdir()) / "enhanced_word_omml_cache"
    cache_dir.mkdir(exist_ok=True)
    xsl_path = cache_dir / "mml2omml.xsl"
    if xsl_path.exists():
        return xsl_path
    # Attempt download; ignore failures (caller will handle)
    try:
        import urllib.request

        with urllib.request.urlopen(MML2OMML_XSL_URL, timeout=10) as resp:
            xsl_path.write_bytes(resp.read())
        return xsl_path
    except Exception:
        return None


def check_dependencies() -> Tuple[bool, str]:
    """Return (ok, msg) after verifying *pandoc* and *xsltproc* are in PATH."""
    for exe in ("pandoc", "xsltproc"):
        if shutil.which(exe) is None:
            return False, f"Required executable '{exe}' not found in PATH"
    xsl_path = _ensure_xsl_cached()
    if xsl_path is None or not xsl_path.exists():
        return False, "Could not download mml2omml.xsl converter stylesheet"
    return True, ""


def latex_to_omml(latex: str) -> Tuple[bool, str]:
    """Convert LaTeX equation to OMML.

    Returns (success, omml_or_error).
    """
    ok, msg = check_dependencies()
    if not ok:
        return False, msg

    xsl_path = _ensure_xsl_cached()
    assert xsl_path is not None and xsl_path.exists()

    with tempfile.TemporaryDirectory() as td:
        tmp_dir = Path(td)
        tex_input = tmp_dir / "in.tex"
        mathml_path = tmp_dir / "out_mathml.xml"

        tex_input.write_text(f"${latex}$", encoding="utf-8")

        # Call pandoc to get MathML
        try:
            subprocess.run(
                [
                    "pandoc",
                    "--from",
                    "latex",
                    "--to",
                    "mathml",
                    str(tex_input),
                    "-o",
                    str(mathml_path),
                ],
                check=True,
                capture_output=True,
            )
        except subprocess.CalledProcessError as e:
            return False, f"pandoc conversion failed: {e.stderr.decode(errors='ignore')[:200]}"

        # Call xsltproc to convert MathML → OMML
        try:
            result = subprocess.run(
                ["xsltproc", str(xsl_path), str(mathml_path)],
                check=True,
                capture_output=True,
            )
        except subprocess.CalledProcessError as e:
            return False, f"xsltproc failed: {e.stderr.decode(errors='ignore')[:200]}"

        omml = result.stdout.decode()
        return True, omml.strip() 