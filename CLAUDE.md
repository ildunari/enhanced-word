# Enhanced Word MCP Server - Project Guide

## Project Overview
Enhanced Word document manipulation MCP server with 24 consolidated tools for comprehensive document processing, editing, and analysis.

## Development Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run server locally
python -m word_document_server.main

# Test functionality  
python test_enhanced_features.py

# Publish to NPM
npm publish
```

## Installation & Usage

```bash
# Install via NPX (recommended)
claude mcp add word-mcp -s user -- npx enhanced-word-mcp-server

# Verify installation
claude mcp list
```

## Project Structure

```
word_document_server/
├── main.py                 # FastMCP server registration (28 tools)
├── tools/
│   ├── document_tools.py   # Core document operations
│   ├── content_tools.py    # Text content and search/replace
│   ├── footnote_tools.py   # Footnotes and endnotes
│   ├── section_tools.py    # Document structure analysis
│   ├── review_tools.py     # Comments and track changes
│   ├── protection_tools.py # Document security
│   └── citation_tools.py   # EndNote citation management
└── utils/
    ├── document_utils.py   # Core Word document utilities
    ├── file_utils.py       # File path handling
    └── citation_utils.py   # EndNote citation extraction
```

## Consolidated Tools (28 total)

**6 Consolidated Tools:**
- `get_text` (replaces 3 tools)
- `manage_track_changes` (replaces 2 tools) 
- `add_note` (replaces 2 tools)
- `add_text_content` (replaces 2 tools)
- `get_sections` (replaces 2 tools)
- `manage_protection` (replaces 2 tools)

**18 Essential Tools:**
- Document management (create, copy, info, merge)
- Advanced features (search/replace, tables, images, PDF)
- Formatting and analysis tools

**4 Citation Tools (NEW):**
- `list_citations` - Extract all EndNote citations with metadata
- `get_citation_at_position` - Get citation details at specific location
- `copy_existing_citation` - Copy citation for reuse
- `analyze_citation_distribution` - Statistical analysis of citations

## Testing Approach

**Test Documents:**
- `comprehensive_test_document.docx` - Main test file
- `comment_test.docx` - Comment functionality testing

**Known Issues:**
- Comment persistence bug: Comments report as added but don't persist (needs investigation)

**Testing Commands:**
```bash
# Test all consolidated tools
python test_enhanced_features.py

# Manual testing via Claude Code
claude mcp get word-mcp
```

## Version History

- **v2.8.0**: Current version with 42 tools and EndNote citation support
- **v2.7.12**: Previous version with 38 tools and equation support
- **v2.2.1**: Consolidated tools version
- **v2.1.1**: Pre-consolidation (47 tools)

## Code Style
- Python type hints required
- Comprehensive error handling
- Detailed docstrings with examples
- FastMCP tool decorators

## Common Issues

1. **File path handling**: Always use absolute paths for reliability
2. **Tool name length**: Keep under 64 characters (why we use 'word-mcp' not 'enhanced-word-mcp')
3. **Import errors**: Check all module imports after file changes
4. **Comment persistence**: Known bug in manage_comments tool