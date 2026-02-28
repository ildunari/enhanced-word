## Cursor Cloud specific instructions

This is a hybrid **Python + Node.js** MCP server with no external services (no DB, no Docker, no network dependencies). Everything runs locally on the filesystem.

### Quick reference

- **Dev commands** are in `CLAUDE.md` (install, test, run).
- **Test commands**: `pytest -q -m "not e2e and not stress"` (unit/integration), `pytest -q -m e2e` (end-to-end), `npm test` (full via venv).
- **Run server**: `python -m word_document_server.main` (stdio JSON-RPC; no HTTP server).

### Non-obvious caveats

- The package must be installed in **editable mode** (`pip install -e .`) for `word_document_server` to be importable. The update script handles this.
- `python3.12-venv` system package is required for `npm test` (it creates a `.venv-test` venv internally via `test_enhanced_features.py`). If the `.venv-test` directory exists but is broken (e.g. created before `python3.12-venv` was installed), delete it and re-run.
- `~/.local/bin` must be on `PATH` for pip-installed scripts (pytest, mcp, etc.). The update script ensures this.
- The MCP server uses **stdio transport only** — it reads JSON-RPC from stdin and writes to stdout. There is no HTTP endpoint to curl. To test interactively, use the `mcp` Python client library (see e2e tests in `tests/test_e2e_mcp_stdio.py`).
- `convert_to_pdf` tool requires LibreOffice (`libreoffice` system package) on Linux. It is not installed by default and is not needed for the core test suite.
