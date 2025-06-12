# Enhanced Word Document MCP Server

[![smithery badge](https://smithery.ai/badge/@GongRzhe/Office-Word-MCP-Server)](https://smithery.ai/server/@GongRzhe/Office-Word-MCP-Server)

A powerful, consolidated Model Context Protocol (MCP) server for creating, reading, and manipulating Microsoft Word documents. This enhanced version provides 22 optimized tools (reduced from 47) for comprehensive Word document operations through a standardized interface.

<a href="https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server/badge" alt="Office Word Server MCP server" />
</a>

![](https://badge.mcpx.dev?type=server "MCP Server")

## Overview

Enhanced-Word-MCP-Server implements the [Model Context Protocol](https://modelcontextprotocol.io/) with a focus on consolidation and efficiency. It provides 22 powerful tools that replace 47 individual functions, offering:

- **53% tool reduction** while preserving all functionality
- **Consolidated operations** for better usability  
- **Enhanced functionality** with regex support and advanced formatting
- **Academic and professional workflow optimization**
- **Comprehensive error handling and validation**

### Example Usage

#### Creating Academic Documents
```python
# Create document with proper structure
create_document("thesis.docx", title="AI in Healthcare", author="John Doe")

# Add structured content
add_text_content("thesis.docx", "Introduction", content_type="heading", level=1)
add_text_content("thesis.docx", "This paper explores...", content_type="paragraph")

# Add references and notes
add_note("thesis.docx", paragraph_index=0, note_text="See methodology section", note_type="footnote")
```

## Features Overview

### 🎯 Consolidated Tools (6 Tools)
Unified operations that replace multiple individual functions:

- **`get_text`** - Unified text extraction (replaces 3 tools)
- **`manage_track_changes`** - Track changes management (replaces 2 tools)  
- **`add_note`** - Footnote/endnote creation (replaces 2 tools)
- **`add_text_content`** - Paragraph/heading creation (replaces 2 tools)
- **`get_sections`** - Section extraction (replaces 2 tools)
- **`manage_protection`** - Document protection (replaces 2 tools)

### 📄 Essential Document Tools (7 Tools)
Core document management functionality:

- **Document Lifecycle**: `create_document`, `copy_document`, `merge_documents`
- **Content Operations**: `enhanced_search_and_replace`, `add_table`, `add_picture`
- **Export**: `convert_to_pdf`

### 🔧 Advanced Features (5 Tools)
Specialized functionality for professional workflows:

- **Collaboration**: `manage_comments`, `extract_track_changes`, `generate_review_summary`
- **Document Structure**: `generate_table_of_contents`
- **Security**: `add_digital_signature`, `verify_document`

## Key Enhancements

### 🚀 Enhanced Search & Replace
- **Regex support** for complex pattern matching
- **Case-insensitive** search options
- **Whole word matching**
- **Advanced formatting** application to replaced text
- **Group substitutions** for regex patterns

```python
# Regex date format conversion
enhanced_search_and_replace("doc.docx", 
    find_text=r"(\d{4})-(\d{2})-(\d{2})", 
    replace_text=r"$2/$3/$1", 
    use_regex=True)

# Case-insensitive formatting
enhanced_search_and_replace("doc.docx", 
    find_text="important", 
    replace_text="CRITICAL",
    match_case=False, 
    apply_formatting=True, 
    bold=True, color="red")
```

### 📝 Unified Text Extraction
```python
# Extract full document with formatting
get_text("doc.docx", scope="document", include_formatting=True)

# Search within document
get_text("doc.docx", scope="search", search_term="methodology", match_case=False)

# Extract specific paragraph
get_text("doc.docx", scope="paragraph", paragraph_index=5)
```

### 📑 Flexible Section Management
```python
# Extract all sections with formatting
get_sections("doc.docx", extraction_type="all", include_formatting=True)

# Get specific section content
get_sections("doc.docx", extraction_type="specific", section_title="Results")
```

### 🔒 Advanced Protection Management
```python
# Password protection
manage_protection("doc.docx", action="protect", protection_type="password", password="secure123")

# Read-only protection with exceptions
manage_protection("doc.docx", action="protect", protection_type="editing", 
                 allowed_editing="comments", password="review123")
```

## Installation

### NPX Installation (Recommended)
```bash
# Install via NPX (latest version)
npx enhanced-word-mcp-server

# Or install globally
npm install -g enhanced-word-mcp-server
```

### Add to Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "npx",
      "args": ["enhanced-word-mcp-server"]
    }
  }
}
```

### Manual Installation
```bash
# Clone repository
git clone https://github.com/your-username/enhanced-word-mcp-server.git
cd enhanced-word-mcp-server

# Install dependencies
npm install

# Global installation
npm install -g .
```

## Usage Examples

### Academic Writing Workflow
```python
# Create research paper structure
create_document("research_paper.docx", title="Machine Learning Analysis", author="Dr. Smith")

# Add structured content
add_text_content("research_paper.docx", "Abstract", content_type="heading", level=1)
add_text_content("research_paper.docx", "This study examines...", content_type="paragraph", 
                style="Normal", position="end")

# Add citations and notes
add_note("research_paper.docx", paragraph_index=1, 
         note_text="See Smith et al. (2023) for detailed methodology", 
         note_type="footnote")

# Format academic terms
format_research_paper_terms("research_paper.docx")

# Extract sections for review
sections = get_sections("research_paper.docx", extraction_type="all", max_level=2)
```

### Document Review Workflow
```python
# Extract all review elements
comments = extract_comments("draft.docx")
changes = extract_track_changes("draft.docx")

# Generate comprehensive review summary
summary = generate_review_summary("draft.docx")

# Manage track changes selectively
manage_track_changes("draft.docx", action="accept", change_type="insertions")
manage_track_changes("draft.docx", action="reject", change_type="deletions", author="Reviewer1")
```

### Document Security Workflow
```python
# Apply comprehensive protection
manage_protection("confidential.docx", action="protect", 
                 protection_type="password", password="secure123")

# Add digital signature
add_digital_signature("contract.docx", signer_name="John Doe", 
                     reason="Document approval")

# Verify document integrity
verification = verify_document("contract.docx")
```

## Tool Reference

### Consolidated Tools

#### `get_text(filename, scope, **options)`
Unified text extraction with multiple modes:
- `scope`: "document" | "paragraph" | "search" | "range"
- `include_formatting`: Extract formatting information
- `search_term`: Text to search for (when scope="search")
- `paragraph_index`: Specific paragraph (when scope="paragraph")

#### `manage_track_changes(filename, action, **filters)`
Comprehensive track changes management:
- `action`: "accept" | "reject" | "extract"
- `change_type`: "all" | "insertions" | "deletions" | "formatting"
- `author`: Filter by specific author
- `date_range`: Filter by date range

#### `add_note(filename, paragraph_index, note_text, note_type, **options)`
Unified footnote/endnote creation:
- `note_type`: "footnote" | "endnote"
- `custom_symbol`: Use custom reference symbol
- `position`: Note positioning options

#### `add_text_content(filename, text, content_type, **options)`
Unified content creation:
- `content_type`: "paragraph" | "heading"
- `level`: Heading level (1-6)
- `style`: Apply document style
- `position`: "start" | "end" | specific index

#### `get_sections(filename, extraction_type, **options)`
Advanced section extraction:
- `extraction_type`: "all" | "specific" | "by_level"
- `section_title`: Specific section to extract
- `max_level`: Maximum heading level
- `include_formatting`: Preserve formatting

#### `manage_protection(filename, action, **options)`
Document protection management:
- `action`: "protect" | "unprotect" | "check"
- `protection_type`: "password" | "editing" | "readonly"
- `password`: Protection password
- `allowed_editing`: Editing permissions

## Error Handling

All tools provide comprehensive error handling:

```python
# Typical error responses
{
  "status": "error",
  "message": "Document not found: nonexistent.docx",
  "error_type": "FileNotFoundError",
  "suggestions": ["Check file path", "Ensure file exists"]
}
```

## Development

### Project Structure
```
enhanced-word-mcp-server/
├── word_document_server/
│   ├── main.py              # MCP server entry point
│   ├── tools/               # Tool implementations
│   │   ├── document_tools.py      # Document management
│   │   ├── content_tools.py       # Content creation
│   │   ├── review_tools.py        # Review and collaboration
│   │   ├── section_tools.py       # Document structure
│   │   ├── protection_tools.py    # Security features
│   │   └── footnote_tools.py      # Notes and references
│   └── utils/               # Utility modules
├── bin/
│   └── enhanced-word-mcp-server.js  # NPX entry point
├── package.json
└── README.md
```

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

### Testing

```bash
# Run test suite
python test_enhanced_features.py

# Test specific functionality
python -c "from word_document_server.tools.content_tools import enhanced_search_and_replace; print(enhanced_search_and_replace('test.docx', 'old', 'new'))"
```

## License

MIT License - see LICENSE file for details.

## Version History

### v2.0.0 (Enhanced)
- 🎯 **53% tool reduction** (47 → 22 tools)
- 🚀 **Enhanced search & replace** with regex support
- 📝 **Consolidated operations** for better usability
- 🔧 **Improved error handling** and validation
- 📚 **Comprehensive documentation** with examples

### v1.0.0 (Original)
- Initial release with 47 individual tools
- Basic Word document operations
- Simple MCP server implementation

## Support

For issues, feature requests, or questions:
- 📧 Create an issue on GitHub
- 📖 Check the documentation and examples
- 🔍 Review error messages for troubleshooting guidance