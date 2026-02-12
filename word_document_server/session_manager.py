"""
Document Session Manager for Word Document Server.

Provides in-memory storage and management of open Word documents with simple IDs,
eliminating the need to pass full file paths for every operation.
"""
import os
from typing import Dict, Optional, List, Any
from dataclasses import dataclass
from docx import Document
from word_document_server.utils.file_utils import ensure_docx_extension


@dataclass
class DocumentHandle:
    """Container for an open Word document with metadata."""
    document_id: str
    file_path: str
    document: Document
    metadata: Dict[str, Any]
    
    def __post_init__(self):
        """Ensure file path has .docx extension."""
        self.file_path = ensure_docx_extension(self.file_path)


class DocumentSessionManager:
    """Manages open Word documents with simple string IDs."""
    
    def __init__(self):
        self._documents: Dict[str, DocumentHandle] = {}
        self._active_document_id: Optional[str] = None
    
    def open_document(self, document_id: str, file_path: str) -> str:
        """
        Open a Word document and assign it a simple ID for future operations.
        
        Args:
            document_id: Simple identifier for the document (e.g., "main", "draft", "review")
            file_path: Full path to the Word document file
            
        Returns:
            Success/error message string
        """
        try:
            # Validate inputs
            if not document_id or not document_id.strip():
                return "Error: document_id cannot be empty"
                
            if not file_path or not file_path.strip():
                return "Error: file_path cannot be empty"

            # Ensure/validate .docx extension. Reject .doc explicitly.
            base, ext = os.path.splitext(file_path)
            if ext and ext.lower() == ".doc":
                return "Error: .doc files are not supported. Please provide a .docx file."
            if not ext:
                file_path = ensure_docx_extension(file_path)
            elif ext.lower() != ".docx":
                return "Error: Invalid file extension. Only .docx is supported."
            
            # Check if file exists
            if not os.path.exists(file_path):
                return f"Error: Document file '{file_path}' does not exist"
            
            # Check if document_id already in use
            if document_id in self._documents:
                return f"Error: Document ID '{document_id}' is already in use. Use close_document() first or choose a different ID."
            
            # Try to open the document
            try:
                doc = Document(file_path)
            except Exception as e:
                return f"Error: Failed to open document '{file_path}': {str(e)}"
            
            # Create document handle with metadata
            metadata = {
                "opened_at": str(os.path.getmtime(file_path)),
                "paragraph_count": len(doc.paragraphs),
                "section_count": len(doc.sections),
                "file_size": os.path.getsize(file_path)
            }
            
            handle = DocumentHandle(
                document_id=document_id,
                file_path=file_path,
                document=doc,
                metadata=metadata
            )
            
            # Store in session
            self._documents[document_id] = handle
            
            # Set as active if it's the first document
            if self._active_document_id is None:
                self._active_document_id = document_id
                
            return f"Successfully opened document '{document_id}' from '{file_path}'"
            
        except Exception as e:
            return f"Error opening document: {str(e)}"
    
    def close_document(self, document_id: str) -> str:
        """
        Close a document and remove it from the session.
        
        Args:
            document_id: ID of the document to close
            
        Returns:
            Success/error message string
        """
        try:
            if document_id not in self._documents:
                return f"Error: Document ID '{document_id}' not found in session"
            
            # Get handle before removing
            handle = self._documents[document_id]
            
            # Remove from session
            del self._documents[document_id]
            
            # Update active document if needed
            if self._active_document_id == document_id:
                # Set active to another document if available, or None
                self._active_document_id = next(iter(self._documents.keys())) if self._documents else None
            
            return f"Successfully closed document '{document_id}' (was '{handle.file_path}')"
            
        except Exception as e:
            return f"Error closing document: {str(e)}"
    
    def list_open_documents(self) -> str:
        """
        List all currently open documents with their metadata.
        
        Returns:
            Formatted string with document information
        """
        try:
            if not self._documents:
                return "No documents currently open"
            
            result = f"Open documents ({len(self._documents)}):\n\n"
            
            for doc_id, handle in self._documents.items():
                is_active = " (ACTIVE)" if doc_id == self._active_document_id else ""
                result += f"ID: {doc_id}{is_active}\n"
                result += f"  Path: {handle.file_path}\n"
                result += f"  Paragraphs: {handle.metadata.get('paragraph_count', 'Unknown')}\n"
                result += f"  Sections: {handle.metadata.get('section_count', 'Unknown')}\n"
                result += f"  File size: {handle.metadata.get('file_size', 'Unknown')} bytes\n\n"
            
            return result.rstrip()
            
        except Exception as e:
            return f"Error listing documents: {str(e)}"
    
    def set_active_document(self, document_id: str) -> str:
        """
        Set the active/default document for operations that support it.
        
        Args:
            document_id: ID of the document to make active
            
        Returns:
            Success/error message string
        """
        try:
            if document_id not in self._documents:
                return f"Error: Document ID '{document_id}' not found in session"
            
            old_active = self._active_document_id
            self._active_document_id = document_id
            
            if old_active:
                return f"Active document changed from '{old_active}' to '{document_id}'"
            else:
                return f"Active document set to '{document_id}'"
                
        except Exception as e:
            return f"Error setting active document: {str(e)}"
    
    def get_document(self, document_id: str) -> Optional[DocumentHandle]:
        """
        Get a document handle by ID.
        
        WARNING: This returns a handle containing the live Document object.
        Callers can directly mutate the document, bypassing any access control.
        Future versions should provide controlled access methods instead.
        
        Args:
            document_id: ID of the document to retrieve
            
        Returns:
            DocumentHandle if found, None otherwise
        """
        return self._documents.get(document_id)
    
    def get_document_path(self, document_id: str) -> Optional[str]:
        """
        Get the file path for a document by ID.
        
        Args:
            document_id: ID of the document
            
        Returns:
            File path if document found, None otherwise
        """
        handle = self.get_document(document_id)
        return handle.file_path if handle else None
    
    def validate_document_id(self, document_id: str) -> str:
        """
        Validate that a document ID exists in the session.
        
        Args:
            document_id: ID to validate
            
        Returns:
            Empty string if valid, error message if invalid
        """
        if not document_id:
            return "Error: document_id parameter is required"
        
        if document_id not in self._documents:
            available = list(self._documents.keys())
            if available:
                return f"Error: Document ID '{document_id}' not found. Available: {', '.join(available)}"
            else:
                return f"Error: Document ID '{document_id}' not found. No documents are currently open. Use open_document() first."
        
        return ""  # Valid
    
    def close_all_documents(self) -> str:
        """
        Close all open documents.
        
        Returns:
            Success message with count
        """
        count = len(self._documents)
        self._documents.clear()
        self._active_document_id = None
        return f"Closed {count} documents"


# Global session manager instance
_session_manager = DocumentSessionManager()


def get_session_manager() -> DocumentSessionManager:
    """Get the global document session manager instance."""
    return _session_manager
