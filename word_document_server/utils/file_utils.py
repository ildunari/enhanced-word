"""
File utility functions for Word Document Server.
"""
import os
from typing import Tuple, Optional, List
import shutil


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to, with special handling for Word documents.
    
    Args:
        filepath: Path to the file
        
    Returns:
        Tuple of (is_writeable, error_message)
    """
    import platform
    import tempfile
    
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == '':
            directory = '.'
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""
    
    # Check basic file permissions
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"
    
    # Check for Word-specific lock files
    if filepath.lower().endswith(('.docx', '.doc')):
        directory = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        
        # Check for temporary Word lock files
        word_lock_patterns = [
            f"~${filename}",  # Word lock file pattern
            f".~lock.{filename}#",  # LibreOffice lock pattern
            f"~WRL{filename[-4:]}.tmp"  # Another Word temp pattern
        ]
        
        for lock_pattern in word_lock_patterns:
            lock_file = os.path.join(directory, lock_pattern)
            if os.path.exists(lock_file):
                return False, f"Document appears to be open in Word (lock file: {lock_pattern})"
    
    # Try to open the file for writing with exclusive access
    try:
        # Create a temporary backup to test exclusive access
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_path = temp_file.name
        
        # Try to copy the file to temp location (tests read access)
        try:
            import shutil
            shutil.copy2(filepath, temp_path)
        except Exception as e:
            os.unlink(temp_path) if os.path.exists(temp_path) else None
            return False, f"Cannot read file {filepath}: {str(e)}"
        
        # Try to open original file for writing (tests write access and locks)
        try:
            # Use different approach based on platform
            if platform.system() == "Windows":
                # On Windows, try to open with exclusive access
                import msvcrt
                try:
                    with open(filepath, 'r+b') as f:
                        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                except (OSError, IOError):
                    os.unlink(temp_path) if os.path.exists(temp_path) else None
                    return False, f"File {filepath} is locked (likely open in Word)"
            else:
                # On Unix-like systems, try to get an exclusive lock
                import fcntl
                try:
                    with open(filepath, 'r+b') as f:
                        fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.flock(f.fileno(), fcntl.LOCK_UN)
                except (OSError, IOError):
                    os.unlink(temp_path) if os.path.exists(temp_path) else None
                    return False, f"File {filepath} is locked (likely open in Word)"
        
        except Exception as e:
            os.unlink(temp_path) if os.path.exists(temp_path) else None
            return False, f"File {filepath} is not writeable: {str(e)}"
        
        # Clean up temp file
        os.unlink(temp_path) if os.path.exists(temp_path) else None

        # --------------------------------------------------------------
        # Capture undo snapshot *after* we have verified the file can be
        # written to but *before* the caller mutates the document.  This
        # guarantees that every successful modification path has a
        # corresponding restore point.
        # --------------------------------------------------------------
        try:
            from word_document_server.undo_manager import get_undo_manager

            if filepath.lower().endswith((".docx", ".doc")) and os.path.isfile(filepath):
                undo_mgr = get_undo_manager()
                undo_mgr.snapshot(filepath)
        except Exception:
            # We never want undo bookkeeping to block the actual operation.
            pass

        return True, ""
        
    except Exception as e:
        return False, f"Error checking file permissions: {str(e)}"



def create_document_copy(source_path: str, dest_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.
    
    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'
        
    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None
    
    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"
    
    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension (case-insensitive check).
    
    Args:
        filename: The filename to check
        
    Returns:
        Filename with .docx extension
    """
    if not filename.lower().endswith('.docx'):
        return filename + '.docx'
    return filename


def sanitize_file_path(filepath: str, allowed_extensions: List[str] = None) -> Tuple[bool, str, str]:
    """
    Sanitize file path to prevent path traversal attacks and ensure valid extensions.
    
    Args:
        filepath: The file path to sanitize
        allowed_extensions: List of allowed file extensions (e.g., ['.docx', '.doc'])
        
    Returns:
        Tuple of (is_valid, sanitized_path, error_message)
    """
    import os.path
    from pathlib import Path
    import re
    
    if not filepath or not isinstance(filepath, str):
        return False, "", "Invalid file path provided"
    
    try:
        # Convert to Path object for better handling
        path = Path(filepath)
        path_str = str(path)

        # Check for path traversal attempts via '..' segments
        if '..' in path.parts:
            return False, "", "Path traversal detected in file path"

        # Validate colon usage (allow Windows drive letters like C:\\)
        if ':' in path_str:
            if not re.match(r'^[a-zA-Z]:[\\/]', path_str):
                return False, "", "Invalid ':' in file path"

        # If absolute path, ensure resolved path does not contain traversal
        if path.is_absolute():
            resolved_path = path.resolve()
            if '..' in resolved_path.parts:
                return False, "", "Path traversal detected in file path"
        
        # Check for dangerous characters
        dangerous_chars = ['<', '>', '|', '*', '?', '"']
        if any(char in str(path) for char in dangerous_chars):
            return False, "", "Invalid characters in file path"
        
        # Validate extension if specified
        if allowed_extensions:
            file_ext = path.suffix.lower()
            if file_ext not in [ext.lower() for ext in allowed_extensions]:
                return False, "", f"Invalid file extension. Allowed: {', '.join(allowed_extensions)}"
        
        # Convert back to string and normalize
        sanitized_path = str(path).replace('\\', '/')
        
        return True, sanitized_path, ""
        
    except Exception as e:
        return False, "", f"Error sanitizing path: {str(e)}"


def validate_docx_path(filepath: str) -> Tuple[bool, str, str]:
    """
    Validate and sanitize a Word document path.
    
    Args:
        filepath: Path to validate
        
    Returns:
        Tuple of (is_valid, sanitized_path, error_message)
    """
    # First sanitize the general path
    is_valid, sanitized_path, error = sanitize_file_path(filepath, ['.docx', '.doc'])
    
    if not is_valid:
        return False, "", error
    
    # Ensure .docx extension
    sanitized_path = ensure_docx_extension(sanitized_path)
    
    return True, sanitized_path, ""