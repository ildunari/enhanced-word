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
import sys

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
        # Attempt silent install
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "latex2mathml>=3.0.0", "--quiet"],
                check=True,
            )
        except FileNotFoundError:
            # pip not available – bootstrap it via ensurepip then retry
            try:
                import ensurepip; ensurepip.bootstrap()
                subprocess.run([sys.executable, "-m", "pip", "install", "latex2mathml>=3.0.0", "--quiet"], check=True)
            except Exception:
                return False, (
                    "Failed to bootstrap pip and install latex2mathml automatically. "
                    "Please run `pip install latex2mathml` in the server's Python environment."
                )
        # Recheck
        if not _has_module("latex2mathml.converter"):
            return False, (
                "Python package 'latex2mathml' could not be imported even after attempted installation."
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
        if not mathml_str:
            return False, "latex2mathml returned empty result"
    except Exception as e:
        return False, f"latex2mathml conversion failed: {e}"

    # Apply XSLT using lxml
    try:
        import lxml.etree as ET

        # The XSLT file has the root template commented out, so we need to uncomment it
        xsl_content = xsl_path.read_text()
        
        # Uncomment the root template that wraps result in oMath
        xsl_content = xsl_content.replace('<!--\n  <xsl:template match="/">', '<xsl:template match="/">')
        xsl_content = xsl_content.replace('</xsl:template>\n-->', '</xsl:template>')
        
        # Parse the modified XSLT
        xsl_doc = ET.fromstring(xsl_content.encode())
        transform = ET.XSLT(xsl_doc)
        
        # Parse MathML and transform
        mathml_doc = ET.fromstring(mathml_str.encode())
        omml_doc = transform(mathml_doc)
        
        if omml_doc is None:
            return False, "XSLT transformation returned None"
            
        # Get the root element
        root = omml_doc.getroot()
        if root is None:
            return False, "XSLT transformation produced empty result"
            
        # Fix n-ary operators with empty <e> elements
        _fix_nary_operators(root)
            
        # Serialize the OMML
        omml = ET.tostring(root, encoding="unicode")
        if omml is None:
            return False, "Failed to serialize OMML result"
            
        return True, omml.strip()
    except Exception as e:
        return False, f"XSLT transform failed: {e}"

# Add this to equation_utils.py

def _fix_nary_operators(root):
    """Fix n-ary operators (integrals, sums, etc.) with empty <e> elements.
    
    The default XSLT conversion leaves <e> empty and puts the content after
    the nary element. This function moves that content into the <e> element.
    """
    import lxml.etree as ET
    
    # Define namespaces
    namespaces = {
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    # Find all nary elements
    nary_elements = root.xpath('.//m:nary', namespaces=namespaces)
    
    for nary in nary_elements:
        # Find the <e> element within this nary
        e_elem = nary.find('m:e', namespaces)
        if e_elem is None:
            continue
            
        # Check if <e> is empty or only has whitespace
        if len(e_elem) > 0 or (e_elem.text and e_elem.text.strip()):
            continue
            
        # Collect elements that should be inside <e>
        # These are siblings after the nary element until we hit another nary or the end
        parent = nary.getparent()
        nary_index = list(parent).index(nary)
        elements_to_move = []
        
        # Collect following siblings that should be part of the integrand/summand
        for i in range(nary_index + 1, len(parent)):
            next_elem = parent[i]
            # Stop if we hit another nary operator
            if next_elem.tag.endswith('nary'):
                break
            # Also stop at certain elements that indicate end of the expression
            tag_local = next_elem.tag.split('}')[-1] if '}' in next_elem.tag else next_elem.tag
            if tag_local in ['oMath', 'oMathPara']:
                break
            elements_to_move.append(next_elem)
            
        # Move collected elements into <e>
        for elem in elements_to_move:
            parent.remove(elem)
            e_elem.append(elem)