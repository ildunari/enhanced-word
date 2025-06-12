"""
Session management tools for Word Document Server.

Provides MCP tools for managing document sessions with simple IDs,
eliminating the need to pass full file paths for every operation.
"""
from word_document_server.session_manager import (
    DocumentSessionManager,
    get_session_manager,
)


def open_document(document_id: str, file_path: str) -> str:
    """
    Open a Word document and assign it a simple ID for future operations.
    
    This tool allows you to open a Word document and reference it by a simple ID
    instead of using the full file path for every subsequent operation.
    
    Args:
        document_id (str): Simple identifier for the document
            - Examples: "main", "draft", "review", "doc1", "thesis"
            - Must be unique among currently open documents
            - Case-sensitive
        file_path (str): Full path to the Word document file
            - Absolute or relative path to .docx file
            - .docx extension will be added automatically if missing
    
    Returns:
        str: Success message with document info, or error message
    
    Examples:
        # Open a document with ID "main"
        result = open_document("main", "/Users/john/Documents/thesis.docx")
        # Returns: "Successfully opened document 'main' from '/Users/john/Documents/thesis.docx'"
        
        # Open multiple documents
        open_document("draft", "./draft_v2.docx")
        open_document("review", "/shared/review_comments.docx")
        
        # Now use document IDs in other tools instead of file paths
        get_text(document_id="main", scope="all")
        manage_comments(document_id="review", action="list")
    
    Use Cases:
        📚 Academic Writing: Open thesis chapters as "intro", "methods", "results"
        👥 Collaboration: Open "original", "review", "final" versions
        🔄 Document Comparison: Open multiple versions for side-by-side work
        ⚡ Efficiency: Avoid typing long file paths repeatedly
    
    Error Handling:
        - File not found: Returns error with file path
        - Document ID already in use: Returns error with suggestion
        - Invalid document format: Returns error with details
        - Empty parameters: Returns parameter validation error
    """
    session_manager = get_session_manager()
    return session_manager.open_document(document_id, file_path)


def close_document(document_id: str) -> str:
    """
    Close a document and remove it from the session.
    
    This tool closes an open document and frees up its ID for reuse.
    The document file itself is not deleted, only removed from the session.
    
    Args:
        document_id (str): ID of the document to close
            - Must be a currently open document ID
            - Case-sensitive
    
    Returns:
        str: Success message with document info, or error message
    
    Examples:
        # Close a specific document
        result = close_document("draft")
        # Returns: "Successfully closed document 'draft' (was '/path/to/draft.docx')"
        
        # Close multiple documents
        close_document("review")
        close_document("temp")
    
    Use Cases:
        🧹 Memory Management: Free up memory from large documents
        🔄 ID Reuse: Close old version to reopen with same ID
        🎯 Focus: Remove distracting documents from session
        ✅ Cleanup: Close completed work documents
    
    Behavior:
        - If closed document was active, another open document becomes active
        - If no other documents open, no active document is set
        - Document ID becomes available for reuse immediately
        - File remains unchanged on disk
    
    Error Handling:
        - Document ID not found: Returns error with available IDs
        - No documents open: Returns appropriate error message
    """
    session_manager = get_session_manager()
    return session_manager.close_document(document_id)


def list_open_documents() -> str:
    """
    List all currently open documents with their metadata.
    
    This tool shows all documents currently available in the session,
    including their IDs, file paths, and basic document information.
    
    Returns:
        str: Formatted list of open documents with metadata
    
    Output Format:
        ```
        Open documents (2):
        
        ID: main (ACTIVE)
          Path: /Users/john/Documents/thesis.docx
          Paragraphs: 156
          Sections: 5
          File size: 2547823 bytes
        
        ID: review
          Path: /shared/review_comments.docx
          Paragraphs: 23
          Sections: 1
          File size: 45123 bytes
        ```
    
    Examples:
        # Check what documents are open
        result = list_open_documents()
        
        # Use in workflow to see available documents
        list_open_documents()  # See what's available
        get_text(document_id="main", scope="all")  # Use an available ID
    
    Use Cases:
        🔍 Discovery: See what documents are available to work with
        📊 Overview: Quick summary of document sizes and content
        🎯 Active Document: See which document is currently active
        🗂️ Organization: Review your current document workspace
        🚨 Debugging: Verify documents opened correctly
    
    Information Displayed:
        - Document ID and active status
        - Full file path
        - Number of paragraphs
        - Number of sections
        - File size in bytes
    
    Special Cases:
        - No open documents: Returns "No documents currently open"
        - Active document marked with "(ACTIVE)" indicator
        - Metadata calculated when document was opened
    """
    session_manager = get_session_manager()
    return session_manager.list_open_documents()


def set_active_document(document_id: str) -> str:
    """
    Set the active/default document for operations that support it.
    
    This tool sets which document should be considered "active" or default
    for operations that might support working with the active document.
    
    Args:
        document_id (str): ID of the document to make active
            - Must be a currently open document ID
            - Case-sensitive
    
    Returns:
        str: Success message showing the change, or error message
    
    Examples:
        # Set main document as active
        result = set_active_document("main")
        # Returns: "Active document changed from 'draft' to 'main'"
        
        # Set first active document
        set_active_document("thesis")
        # Returns: "Active document set to 'thesis'"
    
    Use Cases:
        🎯 Default Context: Set primary document for workflow
        🔄 Context Switching: Change focus between documents
        📝 Primary Document: Mark main document in multi-doc workflow
        ⚡ Efficiency: Reduce need to specify document_id repeatedly
    
    Behavior:
        - Active document is marked with "(ACTIVE)" in list_open_documents()
        - First opened document automatically becomes active
        - When active document is closed, another becomes active
        - Some tools may use active document as default (future feature)
    
    Error Handling:
        - Document ID not found: Returns error with available IDs
        - Empty document_id: Returns parameter validation error
    
    Note:
        Currently this is primarily for organizational purposes.
        Future versions may allow tools to operate on active document by default.
    """
    session_manager = get_session_manager()
    return session_manager.set_active_document(document_id)


def close_all_documents() -> str:
    """
    Close all open documents and clear the session.
    
    This tool closes all currently open documents and resets the session.
    Useful for cleanup or starting fresh.
    
    Returns:
        str: Success message with count of closed documents
    
    Examples:
        # Close everything
        result = close_all_documents()
        # Returns: "Closed 3 documents"
        
        # Start fresh workflow
        close_all_documents()
        open_document("new", "path/to/new_document.docx")
    
    Use Cases:
        🧹 Session Cleanup: Clear all documents to start fresh
        🔄 Workflow Reset: End current work and begin new task
        💾 Memory Management: Free up memory from all open documents
        🚨 Emergency Reset: Clear session if documents are problematic
    
    Behavior:
        - All document IDs become available for reuse
        - No active document after operation
        - Files remain unchanged on disk
        - Session state is completely reset
    
    Special Cases:
        - No open documents: Returns "Closed 0 documents"
        - Cannot be undone: Must reopen documents individually
    """
    session_manager = get_session_manager()
    return session_manager.close_all_documents()


def session_manager(
    action: str,
    document_id: str = None,
    file_path: str = None
) -> str:
    """Unified session management function for all document session operations.
    
    This consolidated tool replaces 5 individual session management functions with a single
    action-based interface, reducing tool count while preserving 100% functionality.
    
    Args:
        action (str): Session operation to perform:
            - "open": Open document with session ID (requires document_id and file_path)
            - "close": Close document session (requires document_id)  
            - "list": List all open document sessions
            - "set_active": Set active document (requires document_id)
            - "close_all": Close all document sessions
        document_id (str, optional): Session document identifier for targeted operations
        file_path (str, optional): File path for opening documents
        
    Returns:
        str: Operation result message or session information
        
    Examples:
        # Open document with session ID
        session_manager("open", document_id="main", file_path="report.docx")
        
        # List all open documents
        session_manager("list")
        
        # Set active document
        session_manager("set_active", document_id="draft")
        
        # Close specific document
        session_manager("close", document_id="main")
        
        # Close all documents
        session_manager("close_all")
    """
    # Validate action parameter
    valid_actions = ["open", "close", "list", "set_active", "close_all"]
    if action not in valid_actions:
        return f"Invalid action: {action}. Must be one of: {', '.join(valid_actions)}"
    
    # Delegate to appropriate original function based on action
    if action == "open":
        if not document_id or not file_path:
            return "Error: Both 'document_id' and 'file_path' are required for action 'open'"
        return open_document(document_id, file_path)
        
    elif action == "close":
        if not document_id:
            return "Error: 'document_id' is required for action 'close'"
        return close_document(document_id)
        
    elif action == "list":
        return list_open_documents()
        
    elif action == "set_active":
        if not document_id:
            return "Error: 'document_id' is required for action 'set_active'"
        return set_active_document(document_id)
        
    elif action == "close_all":
        return close_all_documents()


# Export consolidated tool list for reference
CONSOLIDATED_TOOLS = [
    'session_manager',  # Consolidated (replaces 5 tools)
    'open_document', 'close_document', 'list_open_documents', 'set_active_document', 'close_all_documents'  # Original tools (for backward compatibility)
]

__all__ = CONSOLIDATED_TOOLS + ['DocumentSessionManager', 'get_session_manager']