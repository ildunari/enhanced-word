# AGENTS Instructions

This repository contains an enhanced Model Context Protocol (MCP) server for Microsoft Word documents.
The server logic is written in Python (see `word_document_server/`) with a small Node.js wrapper (`index.js` and scripts in `bin/`).

## Contribution Guidelines

- **Python version**: Use Python 3.9 or higher (tested with 3.11).
- **Formatting**: Follow PEP 8 style. Provide docstrings for public functions.
- **JavaScript**: Keep Node code minimal and compatible with Node 14+.
- **Testing**: After modifications run:
  ```bash
  python test_enhanced_features.py
  ```
  Ensure the tests pass before committing.
- **File structure**: Place new tools under `word_document_server/tools` and helper functions under `word_document_server/utils`.
- **Documentation**: Update `README.md` or `README_ENHANCED.md` when adding significant features.

