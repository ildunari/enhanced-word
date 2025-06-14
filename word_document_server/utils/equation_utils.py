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
import urllib.request
import contextlib
from pathlib import Path
from typing import Tuple

import importlib

MML2OMML_XSL_URL = (
    "https://raw.githubusercontent.com/plgonzalezrx/mathml2omml/master/mml2omml.xsl"
)


def _ensure_xsl_cached() -> Path | None:
    """Locate ``mml2omml.xsl``.

    Search order:
      1. File bundled next to this module (preferred, works offline)
      2. Previously cached copy in temp dir
    """
    module_dir = Path(__file__).parent
    local_path = module_dir / "mml2omml.xsl"
    if local_path.exists():
        return local_path

    cache_dir = Path(tempfile.gettempdir()) / "enhanced_word_omml_cache"
    cache_dir.mkdir(exist_ok=True)
    xsl_path = cache_dir / "mml2omml.xsl"
    if xsl_path.exists():
        return xsl_path

    # Attempt download if internet access is available
    try:
        with contextlib.closing(
            urllib.request.urlopen(MML2OMML_XSL_URL, timeout=10)
        ) as response:
            data = response.read()
            # Write the downloaded stylesheet to cache for future runs
            xsl_path.write_bytes(data)
            return xsl_path
    except Exception:
        # Silent failure – caller will report missing dependency
        return None


def _has_module(module_name: str) -> bool:
    """Return True if *module_name* can be imported."""
    spec = importlib.util.find_spec(module_name)
    return spec is not None


def check_dependencies() -> Tuple[bool, str]:
    """Return (ok, msg) after verifying required Python modules and XSL file."""
    if not _has_module("latex2mathml.converter"):
        return False, (
            "Python package 'latex2mathml' is missing. Install it via `pip install latex2mathml`."
        )

    xsl_path = _ensure_xsl_cached()
    if xsl_path is None or not xsl_path.exists():
        return False, (
            "mml2omml.xsl stylesheet not found. Provide it next to equation_utils.py "
            "or ensure internet access to download automatically."
        )

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

    # Pure-python conversion pipeline
    try:
        from latex2mathml.converter import convert as latex2mathml_convert
    except ImportError:
        return False, (
            "Python package 'latex2mathml' is missing. Install it via `pip install latex2mathml`."
        )

    try:
        mathml_str = latex2mathml_convert(latex)
    except Exception as e:
        return False, f"latex2mathml conversion failed: {e}"

    # Apply XSLT using lxml
    try:
        import lxml.etree as ET

        xsl_doc = ET.parse(str(xsl_path))
        transform = ET.XSLT(xsl_doc)
        mathml_doc = ET.fromstring(mathml_str.encode())
        omml_doc = transform(mathml_doc)
        omml = ET.tostring(omml_doc, encoding="unicode")
        return True, omml.strip()
    except Exception as e:
        return False, f"XSLT transform failed: {e}" 