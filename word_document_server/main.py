"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
"""

import os
import sys
from mcp.server.fastmcp import FastMCP
from word_document_server.tools import (
    document_tools,
    content_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    review_tools,
    section_tools,
    session_tools,
    undo_tools,
    equation_tools,
    citation_tools,
)



# Initialize FastMCP server
mcp = FastMCP("word-document-server")

def register_tools():
    """Register all tools with the MCP server - CONSOLIDATED VERSION WITH SESSION MANAGEMENT."""
    
    # ========== SESSION MANAGEMENT TOOLS (CONSOLIDATED) ==========
    # Unified session management (replaces 5 individual tools)
    mcp.tool()(session_tools.session_manager)
    
    # ========== CONSOLIDATED TOOLS (NEW) ==========
    # These replace multiple existing tools with enhanced functionality
    
    # Unified text extraction (replaces get_document_text, get_paragraph_text_from_document, find_text_in_document)
    mcp.tool()(document_tools.get_text)
    
    # Unified track changes management (replaces accept_all_changes, reject_all_changes)
    mcp.tool()(review_tools.manage_track_changes)
    
    # Unified note addition (replaces add_footnote_to_document, add_endnote_to_document)
    mcp.tool()(footnote_tools.add_note)
    
    # Unified text content addition (replaces add_paragraph, add_heading)
    mcp.tool()(content_tools.add_text_content)
    
    # Unified section extraction (replaces extract_sections_by_heading, extract_section_content)
    mcp.tool()(section_tools.get_sections)
    
    # Unified protection management (replaces protect_document, unprotect_document)
    mcp.tool()(protection_tools.manage_protection)
    
    # Enhanced comment management (replaces extract_comments with full lifecycle management)
    mcp.tool()(review_tools.manage_comments)
    
    # ========== CONSOLIDATED DOCUMENT TOOLS (NEW) ==========
    # Unified document utilities (replaces 3 individual tools)
    mcp.tool()(document_tools.document_utility)
    
    # ========== ESSENTIAL DOCUMENT TOOLS (7) ==========
    # Core document management that cannot be consolidated
    mcp.tool()(document_tools.create_document)
    mcp.tool()(document_tools.copy_document)
    mcp.tool()(document_tools.merge_documents)
    mcp.tool()(content_tools.enhanced_search_and_replace)
    mcp.tool()(content_tools.add_table)
    mcp.tool()(content_tools.add_picture)
    mcp.tool()(extended_document_tools.convert_to_pdf)
    
    # ========== CONSOLIDATED FORMATTING TOOLS (NEW) ==========
    # Unified document formatting (replaces 2 individual tools)
    mcp.tool()(content_tools.format_document)
    
    # ========== ADVANCED FEATURES (5) ==========
    # Specialized functionality for advanced use cases
    mcp.tool()(review_tools.extract_track_changes)
    mcp.tool()(review_tools.generate_review_summary)
    mcp.tool()(section_tools.generate_table_of_contents)
    mcp.tool()(protection_tools.add_digital_signature)
    mcp.tool()(protection_tools.verify_document)

    # ========== CITATION MANAGEMENT (4) ==========
    # EndNote and field-based citation support
    # Consolidated to keep total tools ≤ 25
    mcp.tool()(citation_tools.citations)
    
    # ========== LEGACY COMPATIBILITY (OPTIONAL) ==========
    # These maintain backwards compatibility - can be removed after transition
    # Uncomment if you need backwards compatibility during transition period
    
    # Legacy text extraction tools (now replaced by get_text)
    # mcp.tool()(document_tools.get_document_text)
    # mcp.tool()(extended_document_tools.get_paragraph_text_from_document)
    # mcp.tool()(extended_document_tools.find_text_in_document)
    
    # Legacy track changes tools (now replaced by manage_track_changes)
    # mcp.tool()(review_tools.accept_all_changes)
    # mcp.tool()(review_tools.reject_all_changes)
    
    # Legacy note tools (now replaced by add_note)
    # mcp.tool()(footnote_tools.add_footnote_to_document)
    # mcp.tool()(footnote_tools.add_endnote_to_document)
    
    # Legacy content tools (now replaced by add_text_content)
    # mcp.tool()(content_tools.add_paragraph)
    # mcp.tool()(content_tools.add_heading)
    
    # Legacy section tools (now replaced by get_sections)
    # mcp.tool()(section_tools.extract_sections_by_heading)
    # mcp.tool()(section_tools.extract_section_content)
    
    # Legacy protection tools (now replaced by manage_protection)
    # mcp.tool()(protection_tools.protect_document)
    # mcp.tool()(protection_tools.unprotect_document)
    
    # Legacy basic search (now replaced by enhanced_search_and_replace and get_text with search scope)
    # mcp.tool()(content_tools.search_and_replace)

    # ========== UNDO / REDO (1) ==========
    # Provides history management while keeping total tool count <= 25
    mcp.tool()(undo_tools.session_undo)

    # ========== EQUATION INSERTION (1) ==========
    mcp.tool()(equation_tools.insert_equation)






def run_server():
    """Run the Word Document MCP Server."""
    # Register all tools
    register_tools()
    
    # Run the server
    mcp.run(transport='stdio')
    return mcp

if __name__ == "__main__":
    run_server()
