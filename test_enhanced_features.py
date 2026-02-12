#!/usr/bin/env python3

from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
import venv


def main() -> int:
    # Compatibility runner for existing docs and `npm test`.
    #
    # This repo is hybrid (Node wrapper + Python server). `npm test` runs in a
    # plain Python environment, so we create a local venv and install dev deps
    # there to keep tests reproducible without requiring global installs.
    root = Path(__file__).resolve().parent
    venv_dir = root / ".venv-test"

    # Determine venv python path (cross-platform).
    if os.name == "nt":
        venv_python = venv_dir / "Scripts" / "python.exe"
    else:
        venv_python = venv_dir / "bin" / "python"

    if not venv_python.exists():
        venv.EnvBuilder(with_pip=True).create(venv_dir)

    # Install runtime + dev deps into the venv.
    deps_cmd = [
        str(venv_python),
        "-m",
        "pip",
        "install",
        "-r",
        str(root / "requirements.txt"),
        "-r",
        str(root / "requirements-dev.txt"),
    ]
    install = subprocess.run(deps_cmd)
    if install.returncode != 0:
        return install.returncode

    result = subprocess.run([str(venv_python), "-m", "pytest", "-q"])
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
