"""
Extended document tools for Word Document Server.

These tools provide enhanced document content extraction and search capabilities.
"""
import os
import subprocess
import platform
import shutil
from typing import Optional
from docx import Document

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, sanitize_file_path
from word_document_server.utils.session_utils import resolve_document_path


async def convert_to_pdf(document_id: str = None, filename: str = None, output_filename: Optional[str] = None) -> str:
    """Convert a Word document to PDF format.
    
    Args:
        document_id (str, optional): Session document identifier (preferred)
        filename (str, optional): Path to the Word document
        output_filename: Optional path for the output PDF. If not provided, 
                         will use the same name with .pdf extension
    """
    # Resolve document path from session or filename
    filename, error_msg = resolve_document_path(document_id, filename)
    if error_msg:
        return error_msg
    
    filename = ensure_docx_extension(filename)
    
    # Sanitize the file path to prevent command injection
    is_valid, sanitized_filename, error = sanitize_file_path(filename, allowed_extensions=['.docx', '.doc'])
    if not is_valid:
        return f"Invalid file path: {error}"
    filename = sanitized_filename
    
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    
    # Generate output filename if not provided
    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith('.pdf'):
        output_filename = f"{output_filename}.pdf"
    
    # Sanitize output filename
    is_valid, sanitized_output, error = sanitize_file_path(output_filename, allowed_extensions=['.pdf'])
    if not is_valid:
        return f"Invalid output file path: {error}"
    output_filename = sanitized_output
    
    # Convert to absolute path if not already
    if not os.path.isabs(output_filename):
        output_filename = os.path.abspath(output_filename)
    
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_filename)
    if not output_dir:
        output_dir = os.path.abspath('.')
    
    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Check if output file can be written
    is_writeable, error_message = check_file_writeable(output_filename)
    if not is_writeable:
        return f"Cannot create PDF: {error_message} (Path: {output_filename}, Dir: {output_dir})"
    
    try:
        # Determine platform for appropriate conversion method
        system = platform.system()
        
        if system == "Windows":
            # On Windows, try docx2pdf which uses Microsoft Word
            try:
                from docx2pdf import convert
                convert(filename, output_filename)
                return f"Document successfully converted to PDF: {output_filename}"
            except (ImportError, Exception) as e:
                return f"Failed to convert document to PDF: {str(e)}\nNote: docx2pdf requires Microsoft Word to be installed."
                
        elif system in ["Linux", "Darwin"]:  # Linux or macOS
            # Try using LibreOffice if available (common on Linux/macOS)
            try:
                # Choose the appropriate command based on OS
                if system == "Darwin":  # macOS
                    lo_commands = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
                else:  # Linux
                    lo_commands = ["libreoffice", "soffice"]
                
                # Try each possible command
                conversion_successful = False
                errors = []
                
                for cmd_name in lo_commands:
                    try:
                        # Construct LibreOffice conversion command
                        output_dir = os.path.dirname(output_filename)
                        # If output_dir is empty, use current directory
                        if not output_dir:
                            output_dir = '.'
                        # Ensure the directory exists
                        os.makedirs(output_dir, exist_ok=True)
                        
                        cmd = [
                            cmd_name, 
                            '--headless', 
                            '--convert-to', 
                            'pdf', 
                            '--outdir', 
                            output_dir, 
                            filename
                        ]
                        
                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                        
                        if result.returncode == 0:
                            # LibreOffice creates the PDF with the same basename
                            base_name = os.path.basename(filename)
                            pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                            created_pdf = os.path.join(os.path.dirname(output_filename) or '.', pdf_base_name)
                            
                            if not os.path.exists(created_pdf):
                                errors.append(
                                    f"{cmd_name} reported success but output PDF not found: {created_pdf}"
                                )
                                continue

                            # If the created PDF is not at the desired location, move it
                            if created_pdf != output_filename:
                                shutil.move(created_pdf, output_filename)

                            if not os.path.exists(output_filename):
                                errors.append(
                                    f"{cmd_name} reported success but output PDF not found: {output_filename}"
                                )
                                continue
                            
                            conversion_successful = True
                            break  # Exit the loop if successful
                        else:
                            errors.append(f"{cmd_name} error: {result.stderr}")
                    except (subprocess.SubprocessError, FileNotFoundError) as e:
                        errors.append(f"{cmd_name} error: {str(e)}")
                
                if conversion_successful:
                    return f"Document successfully converted to PDF: {output_filename}"
                else:
                    # If all LibreOffice attempts failed, try docx2pdf as fallback
                    try:
                        from docx2pdf import convert
                        convert(filename, output_filename)
                        return f"Document successfully converted to PDF: {output_filename}"
                    except (ImportError, Exception) as e:
                        error_msg = "Failed to convert document to PDF using LibreOffice or docx2pdf.\n"
                        error_msg += "LibreOffice errors: " + "; ".join(errors) + "\n"
                        error_msg += f"docx2pdf error: {str(e)}\n"
                        error_msg += "To convert documents to PDF, please install either:\n"
                        error_msg += "1. LibreOffice (recommended for Linux/macOS)\n"
                        error_msg += "2. Microsoft Word (required for docx2pdf on Windows/macOS)"
                        return error_msg
                        
            except Exception as e:
                return f"Failed to convert document to PDF: {str(e)}"
        else:
            return f"PDF conversion not supported on {system} platform"
            
    except Exception as e:
        return f"Failed to convert document to PDF: {str(e)}"
