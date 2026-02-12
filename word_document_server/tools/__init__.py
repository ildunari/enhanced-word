"""
MCP tool implementations for the Enhanced Word Document Server.

This package exports tool callables and provides a stable list of the tool names
registered by `word_document_server/main.py`.
"""

# ========== CONSOLIDATED TOOLS ==========
# These unified tools replace numerous legacy functions; total registered tools = 25.

# Document tools - consolidated
from word_document_server.tools.document_tools import (
    get_text,  # Replaces: get_document_text, get_paragraph_text_from_document, find_text_in_document
    document_utility,  # Consolidated (replaces 3 tools)
    create_document, 
    get_document_info, 
    get_document_outline, 
    list_available_documents, 
    copy_document, 
    merge_documents
)

# Content tools - consolidated  
from word_document_server.tools.content_tools import (
    add_text_content,  # Replaces: add_paragraph, add_heading
    enhanced_search_and_replace,  # Enhanced version with regex support
    format_document,  # Consolidated (replaces 2 tools)
    format_specific_words,
    format_research_paper_terms,
    add_table, 
    add_picture
)

# Review tools - consolidated
from word_document_server.tools.review_tools import (
    manage_track_changes,  # Replaces: accept_all_changes, reject_all_changes  
    manage_comments,  # list-only marker scanning (replaces extract_comments)
    extract_track_changes,
    generate_review_summary
)

# Section tools - consolidated
from word_document_server.tools.section_tools import (
    get_sections,  # Replaces: extract_sections_by_heading, extract_section_content
    generate_table_of_contents
)

# Protection tools - consolidated
from word_document_server.tools.protection_tools import (
    manage_protection,  # Replaces: protect_document, unprotect_document
    add_digital_signature,
    verify_document
)

# Footnote tools - consolidated
from word_document_server.tools.footnote_tools import (
    add_note  # Disabled (python-docx footnote/endnote limitation)
)

# Extended document tools
from word_document_server.tools.extended_document_tools import (
    convert_to_pdf
)

# Session management tools  
from word_document_server.tools.session_tools import (
    session_manager,  # Consolidated (replaces 5 tools)
    open_document,
    close_document,
    list_open_documents,
    set_active_document,
    close_all_documents
)

# Undo / Redo tool (single entry for multiple actions)
from word_document_server.tools.undo_tools import session_undo  # noqa: F401

# Equation tool
from word_document_server.tools.equation_tools import insert_equation  # noqa: F401

# Citation tool (consolidated entry point)
from word_document_server.tools.citation_tools import citations  # noqa: F401

# Tools registered by `word_document_server/main.py` (source of truth for tool count).
REGISTERED_TOOL_NAMES = [
    "session_manager",
    "get_text",
    "manage_track_changes",
    "add_note",
    "add_text_content",
    "get_sections",
    "manage_protection",
    "manage_comments",
    "document_utility",
    "create_document",
    "copy_document",
    "merge_documents",
    "enhanced_search_and_replace",
    "add_table",
    "add_picture",
    "convert_to_pdf",
    "format_document",
    "extract_track_changes",
    "generate_review_summary",
    "generate_table_of_contents",
    "add_digital_signature",
    "verify_document",
    "citations",
    "session_undo",
    "insert_equation",
]

# Exported symbols for compatibility (includes legacy names not registered by main.py).
CONSOLIDATED_TOOLS = [
    # 3 Consolidated Wrapper Tools (replaces 10 original tools)
    'session_manager',  # Replaces 5 session tools
    'document_utility',  # Replaces 3 document info tools  
    'format_document',  # Replaces 2 formatting tools
    
    # 6 Unified Tools (already consolidated)
    'get_text', 'manage_track_changes', 'manage_comments', 'add_note', 'add_text_content', 
    'get_sections', 'manage_protection',
    
    # 7 Essential Document Tools  
    'create_document', 'copy_document', 'merge_documents', 'enhanced_search_and_replace', 
    'add_table', 'add_picture', 'convert_to_pdf',
    
    # 5 Advanced Features
    'extract_track_changes', 'generate_review_summary', 'generate_table_of_contents',
    'add_digital_signature', 'verify_document',
    
    # Legacy tools (for backward compatibility - not registered in main.py)
    'open_document', 'close_document', 'list_open_documents', 'set_active_document', 'close_all_documents',
    'get_document_info', 'get_document_outline', 'list_available_documents',
    'format_specific_words', 'format_research_paper_terms',

    # New undo/redo capability
    'session_undo',
    'insert_equation'
]

# Total tools registered by main.py (must match `REGISTERED_TOOL_NAMES`).
REGISTERED_TOOL_COUNT = len(REGISTERED_TOOL_NAMES)
TOTAL_TOOL_COUNT = REGISTERED_TOOL_COUNT

__all__ = CONSOLIDATED_TOOLS + [
    "CONSOLIDATED_TOOLS",
    "REGISTERED_TOOL_NAMES",
    "REGISTERED_TOOL_COUNT",
    "TOTAL_TOOL_COUNT",
]
