"""
Session utility functions for backward compatibility and document access.

Provides helper functions to support both document_id and filename parameters
in tools, allowing for gradual migration to session-based document management.
"""
import os
from typing import Optional, Tuple
from word_document_server.session_manager import get_session_manager
from word_document_server.utils.file_utils import ensure_docx_extension


def resolve_document_path(document_id: Optional[str] = None, filename: Optional[str] = None) -> Tuple[str, str]:
    """
    Resolve document_id or filename to actual file path.
    
    Supports both new session-based approach (document_id) and legacy approach (filename)
    for backward compatibility during transition period.
    
    Args:
        document_id: Optional session document ID
        filename: Optional direct file path (legacy)
        
    Returns:
        Tuple of (file_path, error_message)
        - If successful: (actual_file_path, "")
        - If error: ("", error_message)
    
    Priority:
        1. If document_id provided, use session manager
        2. If filename provided, use directly (legacy mode)
        3. If neither provided, return error
        4. If both provided, prefer document_id with warning
    """
    session_manager = get_session_manager()
    
    # Validate input parameters
    if not document_id and not filename:
        return "", "Error: Either 'document_id' or 'filename' parameter is required"
    
    # Handle case where both are provided
    if document_id and filename:
        # Prefer document_id but warn about dual usage
        validation_error = session_manager.validate_document_id(document_id)
        if validation_error:
            return "", f"Error: Both document_id and filename provided. {validation_error}"
        
        file_path = session_manager.get_document_path(document_id)
        return file_path, ""
    
    # Handle document_id approach (preferred)
    if document_id:
        validation_error = session_manager.validate_document_id(document_id)
        if validation_error:
            return "", validation_error
        
        file_path = session_manager.get_document_path(document_id)
        if not file_path:
            return "", f"Error: Could not retrieve file path for document_id '{document_id}'"
        
        return file_path, ""
    
    # Handle filename approach (legacy)
    if filename:
        file_path = filename.strip()
        _, ext = os.path.splitext(file_path)
        if ext and ext.lower() == ".doc":
            return "", "Error: .doc files are not supported. Please provide a .docx file."
        if not ext:
            file_path = ensure_docx_extension(file_path)
        elif ext.lower() != ".docx":
            return "", "Error: Invalid file extension. Only .docx is supported."
        return file_path, ""
    
    # Should never reach here
    return "", "Error: Unexpected parameter resolution failure"


def get_session_document(document_id: str):
    """
    Get a document object from the session.
    
    Args:
        document_id: Session document ID
        
    Returns:
        Document object if found, None otherwise
    """
    session_manager = get_session_manager()
    handle = session_manager.get_document(document_id)
    return handle.document if handle else None


def update_session_document(document_id: str, updated_document) -> bool:
    """
    Update a document object in the session after modifications.
    
    Args:
        document_id: Session document ID
        updated_document: Modified document object
        
    Returns:
        True if updated successfully, False otherwise
    """
    session_manager = get_session_manager()
    handle = session_manager.get_document(document_id)
    
    if handle:
        handle.document = updated_document
        return True
    
    return False
