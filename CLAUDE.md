# Enhanced Word MCP Server - Project Guide

## Project Overview
Enhanced Word document manipulation MCP server with **25 registered tools** exposed by `word_document_server/main.py`.

Core domains:
- Session/document lifecycle
- Text extraction and search/replace
- Review, sections, and protection
- Citations, equations, and undo/redo

## Runtime & Versions
- Python: `>=3.10`
- Package version: `2.7.13` (Node + Python metadata aligned)

## Development Commands

```bash
# Install runtime deps
pip install -r requirements.txt

# Install test/dev deps
pip install -r requirements-dev.txt

# Run server locally
python -m word_document_server.main

# Full local verification
python -m compileall -q word_document_server
pytest -q
npm test
```

## Test Matrix

```bash
# Unit + integration tests
pytest -q -m "not e2e"

# End-to-end MCP stdio tests
pytest -q -m e2e tests/test_e2e_mcp_stdio.py
```

## Search/Replace Guardrails

The following environment variables control safety limits:

- `EW_MAX_MATCHES_PER_CALL` (default `1000`)
- `EW_MAX_SEARCH_OUTPUT_CHARS` (default `200000`)
- `EW_MAX_REGEX_PATTERN_CHARS` (default `5000`)
- `EW_MAX_REGEX_SCAN_CHARS` (default `2000000`)

Behavior:
- replacement/search operations return explicit truncation/guardrail messages when limits are hit.

## Project Structure

```
word_document_server/
├── main.py                 # FastMCP server registration (source of truth for tools)
├── tools/                  # MCP tool implementations
├── utils/                  # Shared helpers
├── core/                   # Core manipulation logic
└── session_manager.py      # Session state

tests/                      # Unit/integration/e2e test suite
test_enhanced_features.py   # npm test compatibility runner
```

## CI
- Workflow: `.github/workflows/tests.yml`
- Jobs:
  - `unit-integration`: compile + non-e2e pytest
  - `e2e-stdio`: stdio MCP round-trip tests

