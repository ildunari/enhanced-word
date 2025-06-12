# Enhanced Word MCP Server

🚀 **Advanced Word document manipulation for academic research collaboration**

An enhanced version of the Office-Word-MCP-Server with revolutionary features designed specifically for academic research workflows, thesis synthesis, and scientific writing collaboration.

## 🌟 Key Enhancements

This enhanced server solves critical limitations in LLM-friendly document manipulation while adding powerful features for academic research:

### 🎯 **Tier 1 Features - Production Ready**

#### 1. **Enhanced Search & Replace with Formatting**
- ✅ **Solves LLM character positioning problem** - No more counting characters!
- ✅ Semantic text targeting with regex support (use `$1` style groups)
- ✅ Simultaneous text replacement and formatting
- ✅ Batch word formatting for research terms
- ✅ Academic research helpers (PCL terminology, statistical notation)

#### 2. **Review Tools & Collaboration**
- ✅ Extract and manage comments with author/timestamp
- ✅ Read, accept, and reject track changes
- ✅ Generate comprehensive review summaries
- ✅ Author-specific change management
- ✅ Add comments programmatically

#### 3. **Section Management via Heading Styles**
- ✅ Extract document structure by heading hierarchy
- ✅ Extract specific section content
- ✅ Generate/update table of contents automatically
- ✅ Document structure statistics
- ✅ Multi-document section merging for thesis synthesis

## 🚀 Quick Start

### Installation

```bash
# Install the enhanced server
npm install -g @kosta/enhanced-word-mcp-server

# Or using npx (no installation needed)
npx @kosta/enhanced-word-mcp-server
```

### Claude Desktop Configuration

Add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "enhanced-word-server": {
      "command": "npx",
      "args": ["@kosta/enhanced-word-mcp-server"]
    }
  }
}
```

## 📖 Enhanced Features Documentation

### Enhanced Search & Replace

The enhanced search and replace feature solves the fundamental problem of character positioning that makes the original format_text tool impractical for LLMs.

```python
# Example: Format research terms with semantic targeting
await enhanced_search_and_replace(
    filename="research_paper.docx",
    find_text="polycaprolactone",
    replace_text="polycaprolactone",
    apply_formatting=True,
    bold=True,
    color="green"
)

# Format multiple research terms at once
await format_research_paper_terms("research_paper.docx")
```

**Key Benefits:**
- No character counting required
- Semantic text matching
- Simultaneous replacement and formatting
- Support for regex patterns and whole-word matching
- `$1` style group references for regex replacements
- Academic research terminology helpers

### Review Tools

Perfect for academic collaboration and thesis advisor feedback:

```python
# Extract all comments and track changes
review_summary = await generate_review_summary("thesis_draft.docx")

# Get changes by specific author
author_changes = await get_author_specific_changes("thesis_draft.docx", "Dr. Smith")

# Accept all track changes (create clean version)
await accept_all_changes("thesis_draft.docx")
```

### Section Management

Ideal for organizing large academic documents and thesis synthesis:

```python
# Extract document structure
structure = await extract_sections_by_heading("thesis.docx")

# Extract specific section for analysis
methods_section = await extract_section_content("thesis.docx", "Methods")

# Generate table of contents
await generate_table_of_contents("thesis.docx", max_level=3)

# Merge sections from multiple documents (thesis synthesis)
await merge_sections_from_documents(
    target_filename="combined_research.docx",
    source_files=["paper1.docx", "paper2.docx", "paper3.docx"],
    section_mapping={
        "paper1.docx": "Results",
        "paper2.docx": "Methods",
        "paper3.docx": "Discussion"
    }
)
```

## 🔬 Academic Research Use Cases

### PCL Mesophase Research Example

```python
# Format scientific terminology consistently
await format_specific_words(
    filename="pcl_research.docx",
    word_list=["dolutegravir", "meloxicam", "dexamethasone", "DTG", "MLX", "DEX"],
    bold=True,
    color="blue"
)

# Format statistical notation
await format_specific_words(
    filename="pcl_research.docx", 
    word_list=["p < 0.05", "r²", "±"],
    bold=True,
    color="red"
)

# Extract and combine results sections from multiple student theses
await merge_sections_from_documents(
    target_filename="combined_pcl_results.docx",
    source_files=[
        "mbaye_dolutegravir_thesis.docx",
        "shah_dexamethasone_thesis.docx", 
        "sharan_meloxicam_thesis.docx"
    ],
    section_mapping={
        "mbaye_dolutegravir_thesis.docx": "Results and Discussion",
        "shah_dexamethasone_thesis.docx": "Results and Discussion",
        "sharan_meloxicam_thesis.docx": "Results and Discussion"
    }
)
```

### Thesis Review Workflow

```python
# Generate comprehensive review summary for advisor
review = await generate_review_summary("student_thesis_v3.docx")

# Extract changes by specific reviewer
advisor_feedback = await get_author_specific_changes("student_thesis_v3.docx", "Dr. Johnson")

# After addressing feedback, accept all changes
await accept_all_changes("student_thesis_final.docx")
```

## 🛠️ Available Tools

### Enhanced Content Tools
- `enhanced_search_and_replace` - Semantic text targeting with formatting
- `format_specific_words` - Batch formatting of terminology
- `format_research_paper_terms` - Academic terminology formatting

### Review & Collaboration Tools
- `extract_comments` - Get all document comments
- `extract_track_changes` - Get all track changes
- `generate_review_summary` - Comprehensive review report
- `accept_all_changes` - Create clean document version
- `reject_all_changes` - Revert to original
- `add_comment` - Add comments programmatically
- `get_author_specific_changes` - Filter by author

### Section Management Tools
- `extract_sections_by_heading` - Document structure analysis
- `extract_section_content` - Get specific section text
- `generate_table_of_contents` - Auto-generate TOC
- `reorganize_sections` - Reorder document sections
- `merge_sections_from_documents` - Combine sections from multiple docs
- `get_section_statistics` - Document metrics and analysis

### Original Features
All original Office-Word-MCP-Server features are preserved:
- Document creation and manipulation
- Text formatting and styling
- Table operations
- Protection and security
- Footnotes and endnotes
- PDF conversion

## 🏗️ Architecture

The enhanced server maintains the original modular architecture while adding three new tool categories:

```
word_document_server/
├── tools/
│   ├── content_tools.py     # Enhanced with new search/replace
│   ├── review_tools.py      # NEW: Collaboration features
│   ├── section_tools.py     # NEW: Document organization
│   ├── document_tools.py    # Original functionality
│   ├── format_tools.py      # Original formatting
│   └── ...
├── core/                    # Shared utilities
└── utils/                   # Helper functions
```

## 🧪 Testing

Run the comprehensive test suite:

```bash
python test_enhanced_features.py
```

This creates a sample academic document and tests all enhanced features.

## 🤝 Contributing

This enhanced server builds upon the excellent foundation of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server).

### Enhancements by Kosta Vučković
- Enhanced search/replace solving LLM character positioning
- Academic collaboration review tools
- Document section management for thesis synthesis
- Academic research workflow optimization

## 📄 License

MIT License - See LICENSE file for details.

## 🙏 Acknowledgments

- Original Office-Word-MCP-Server by [GongRzhe](https://github.com/GongRzhe)
- Enhanced for academic research collaboration at Brown University
- Designed for PCL mesophase drug delivery research synthesis

---

**Perfect for:** Academic researchers, thesis advisors, graduate students, scientific writing collaboration, multi-document synthesis, and any workflow requiring advanced Word document manipulation with LLM-friendly interfaces.