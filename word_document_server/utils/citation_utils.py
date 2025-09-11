"""
Citation utilities for handling EndNote and other citation fields in Word documents.

This module provides functionality to extract, parse, and manage citations
that are stored as fields in Word documents, particularly EndNote citations.
"""

from typing import List, Dict, Any, Optional, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.oxml.ns import qn
from lxml import etree
import re
import json


def extract_fields_from_run(run: Run) -> List[Dict[str, Any]]:
    """
    Extract field information from a run element.
    
    Fields in Word can be:
    - Simple fields: <w:fldSimple w:instr="ADDIN EN.CITE ...">
    - Complex fields: <w:fldChar w:fldCharType="begin"/>...<w:fldChar w:fldCharType="end"/>
    - Hyperlinks: <w:hyperlink w:anchor="_ENREF_XX"> (common for EndNote citations)
    
    Args:
        run: A python-docx Run object
        
    Returns:
        List of field dictionaries containing:
        - type: 'simple', 'complex', or 'hyperlink'
        - instruction: The field instruction code
        - content: The displayed text
        - field_type: Parsed field type (e.g., 'ADDIN', 'REF', 'HYPERLINK')
        - metadata: Additional parsed metadata
    """
    fields = []
    
    # Get the run's XML element
    run_element = run._element
    
    # Check if this run is part of a hyperlink (EndNote often uses this format)
    parent = run_element.getparent()
    if parent is not None and parent.tag == qn('w:hyperlink'):
        anchor = parent.get(qn('w:anchor'), '')
        tooltip = parent.get(qn('w:tooltip'), '')
        
        # EndNote citations typically have anchors like "_ENREF_47"
        if anchor.startswith('_ENREF_'):
            ref_num = anchor.replace('_ENREF_', '')
            
            # Get the display text from the run
            display_text = run.text
            
            field_info = {
                'type': 'hyperlink',
                'instruction': f'HYPERLINK \\l "{anchor}"',
                'content': display_text,
                'field_type': 'HYPERLINK',
                'metadata': {
                    'citation_type': 'endnote',
                    'reference_id': anchor,
                    'record_number': ref_num,
                    'tooltip': tooltip,
                    'anchor': anchor
                }
            }
            
            # Parse tooltip for author/year info if available
            if tooltip:
                # Common format: "Author, Year #RecordNum"
                import re
                match = re.match(r'([^,]+),\s*(\d{4})', tooltip)
                if match:
                    field_info['metadata']['author'] = match.group(1)
                    field_info['metadata']['year'] = match.group(2)
            
            fields.append(field_info)
            return fields
    
    # Original field extraction code follows...
    """
    Extract field information from a run element.
    
    Fields in Word can be simple or complex:
    - Simple fields: <w:fldSimple w:instr="ADDIN EN.CITE ...">
    - Complex fields: <w:fldChar w:fldCharType="begin"/>...<w:fldChar w:fldCharType="end"/>
    
    Args:
        run: A python-docx Run object
        
    Returns:
        List of field dictionaries containing:
        - type: 'simple' or 'complex'
        - instruction: The field instruction code
        - content: The displayed text
        - field_type: Parsed field type (e.g., 'ADDIN', 'REF', 'HYPERLINK')
        - metadata: Additional parsed metadata
    """
    fields = []
    
    # Get the run's XML element
    run_element = run._element
    
    # Check for simple fields in the run
    # Use namespaced xpath
    w = qn('w:fldSimple')
    for fld_simple in run_element.findall('.//' + w):
        instruction = fld_simple.get(qn('w:instr'), '')
        
        # Extract the text content from the field
        field_text = ''
        for text_elem in fld_simple.findall('.//' + qn('w:t')):
            field_text += text_elem.text or ''
        
        field_info = parse_field_instruction(instruction)
        field_info.update({
            'type': 'simple',
            'content': field_text,
            'xml_element': etree.tostring(fld_simple, encoding='unicode', pretty_print=True)
        })
        
        fields.append(field_info)
    
    # Check for complex fields
    # Complex fields span multiple runs, so we need to check if this run is part of one
    parent = run_element.getparent()
    if parent is not None:
        # Look for field char elements
        fld_chars = run_element.findall('.//' + qn('w:fldChar'))
        
        for fld_char in fld_chars:
            fld_type = fld_char.get(qn('w:fldCharType'), '')
            
            if fld_type == 'begin':
                # This is the start of a complex field
                # We need to find the instruction and end
                field_data = extract_complex_field_from_position(parent, run_element)
                if field_data:
                    fields.append(field_data)
    
    return fields


def extract_complex_field_from_position(parent_element, start_run_element) -> Optional[Dict[str, Any]]:
    """
    Extract a complete complex field starting from a given run position.
    
    Complex fields consist of:
    1. Begin marker: <w:fldChar w:fldCharType="begin"/>
    2. Instruction: <w:instrText>ADDIN EN.CITE ...</w:instrText>
    3. Separator: <w:fldChar w:fldCharType="separate"/>
    4. Result text: The displayed content
    5. End marker: <w:fldChar w:fldCharType="end"/>
    """
    runs = parent_element.findall('.//' + qn('w:r'))
    start_index = runs.index(start_run_element)
    
    instruction = ''
    content = ''
    in_instruction = False
    in_result = False
    
    for i in range(start_index, len(runs)):
        run = runs[i]
        
        # Check for field char markers
        fld_chars = run.findall('.//' + qn('w:fldChar'))
        for fld_char in fld_chars:
            fld_type = fld_char.get(qn('w:fldCharType'), '')
            
            if fld_type == 'begin':
                in_instruction = True
            elif fld_type == 'separate':
                in_instruction = False
                in_result = True
            elif fld_type == 'end':
                # Field complete
                field_info = parse_field_instruction(instruction)
                field_info.update({
                    'type': 'complex',
                    'content': content.strip()
                })
                return field_info
        
        # Extract instruction text
        if in_instruction:
            for instr in run.findall('.//' + qn('w:instrText')):
                instruction += instr.text or ''
        
        # Extract result text
        if in_result:
            for text in run.findall('.//' + qn('w:t')):
                content += text.text or ''
    
    return None


def parse_field_instruction(instruction: str) -> Dict[str, Any]:
    """
    Parse a field instruction to extract type and metadata.
    
    Common EndNote citation format:
    ADDIN EN.CITE <EndNote><Cite><Author>Smith</Author><Year>2023</Year>...</Cite></EndNote>
    
    Other field types:
    REF _Ref123456789 \\h
    HYPERLINK "http://example.com"
    """
    field_info = {
        'instruction': instruction,
        'field_type': '',
        'metadata': {}
    }
    
    # Extract field type (first word)
    parts = instruction.strip().split(None, 1)
    if parts:
        field_info['field_type'] = parts[0]
    
    # Parse EndNote citations (both EN.CITE and EN.JS.CITE formats)
    if field_info['field_type'] == 'ADDIN' and ('EN.CITE' in instruction or 'EN.JS.CITE' in instruction):
        field_info['metadata']['citation_type'] = 'endnote'
        
        # Try to extract EndNote XML data
        endnote_match = re.search(r'<EndNote>(.*?)</EndNote>', instruction, re.DOTALL)
        if endnote_match:
            try:
                # Parse the EndNote XML
                endnote_xml = endnote_match.group(0)
                # For now, just store the raw XML
                field_info['metadata']['endnote_xml'] = endnote_xml
                
                # Extract basic citation info
                author_match = re.search(r'<Author>(.*?)</Author>', endnote_xml)
                year_match = re.search(r'<Year>(.*?)</Year>', endnote_xml)
                record_match = re.search(r'<RecNum>(\d+)</RecNum>', endnote_xml)
                
                if author_match:
                    field_info['metadata']['author'] = author_match.group(1)
                if year_match:
                    field_info['metadata']['year'] = year_match.group(1)
                if record_match:
                    field_info['metadata']['record_number'] = record_match.group(1)
                    
            except Exception as e:
                field_info['metadata']['parse_error'] = str(e)
    
    # Parse REF fields (cross-references)
    elif field_info['field_type'] == 'REF':
        ref_match = re.match(r'REF\s+(\S+)', instruction)
        if ref_match:
            field_info['metadata']['reference_id'] = ref_match.group(1)
    
    # Parse HYPERLINK fields
    elif field_info['field_type'] == 'HYPERLINK':
        link_match = re.search(r'"([^"]+)"', instruction)
        if link_match:
            field_info['metadata']['url'] = link_match.group(1)
    
    return field_info


def extract_citations_from_paragraph(paragraph: Paragraph) -> List[Dict[str, Any]]:
    """
    Extract all citations from a paragraph.
    
    Returns a list of citation objects with their position and metadata.
    """
    citations = []
    char_position = 0
    
    for run_idx, run in enumerate(paragraph.runs):
        # First check if the run contains any fields
        fields = extract_fields_from_run(run)
        
        for field in fields:
            # Check for EndNote citations in various formats
            is_endnote = False
            if field['field_type'] == 'ADDIN':
                # Check if it's any EndNote format
                if 'EN.' in field.get('instruction', ''):
                    is_endnote = True
                    field['metadata']['citation_type'] = 'endnote'
            elif field['field_type'] == 'HYPERLINK' and field.get('metadata', {}).get('citation_type') == 'endnote':
                is_endnote = True
                
            if is_endnote:
                citation = {
                    'run_index': run_idx,
                    'character_position': char_position,
                    'display_text': field['content'],
                    'field_type': field['field_type'],
                    'metadata': field['metadata'],
                    'instruction': field['instruction']
                }
                citations.append(citation)
        
        # Add the length of the run text to track position
        # Note: For fields, we use the display content length
        if fields:
            char_position += len(fields[0].get('content', ''))
        else:
            char_position += len(run.text)
    
    return citations


def extract_all_citations_from_document(doc: Document) -> Dict[str, Any]:
    """
    Extract all citations from a Word document.
    
    Returns a dictionary containing:
    - total_citations: Total number of citations found
    - citations_by_paragraph: List of paragraphs containing citations with details
    - unique_references: Set of unique references cited
    - citation_summary: Summary statistics
    """
    result = {
        'total_citations': 0,
        'citations_by_paragraph': [],
        'unique_references': set(),
        'citation_summary': {}
    }
    
    for para_idx, paragraph in enumerate(doc.paragraphs):
        citations = extract_citations_from_paragraph(paragraph)
        
        if citations:
            para_info = {
                'paragraph_index': para_idx,
                'paragraph_text': paragraph.text,
                'citations': citations,
                'citation_count': len(citations)
            }
            
            result['citations_by_paragraph'].append(para_info)
            result['total_citations'] += len(citations)
            
            # Track unique references
            for citation in citations:
                record_num = citation.get('metadata', {}).get('record_number')
                if record_num:
                    result['unique_references'].add(record_num)
    
    # Convert set to list for JSON serialization
    result['unique_references'] = list(result['unique_references'])
    
    # Create summary
    result['citation_summary'] = {
        'total_paragraphs_with_citations': len(result['citations_by_paragraph']),
        'total_unique_references': len(result['unique_references']),
        'average_citations_per_paragraph': (
            result['total_citations'] / len(result['citations_by_paragraph'])
            if result['citations_by_paragraph'] else 0
        )
    }
    
    return result


def create_citation_field_xml(citation_text: str, record_number: str, 
                            author: str = "", year: str = "") -> str:
    """
    Create the XML structure for an EndNote citation field.
    
    This creates a field instruction that can be inserted into a Word document.
    Note: This is a simplified version and may need adjustment based on 
    EndNote's specific requirements.
    """
    endnote_xml = f"""<EndNote><Cite><Author>{author}</Author><Year>{year}</Year><RecNum>{record_number}</RecNum></Cite></EndNote>"""
    
    instruction = f"ADDIN EN.CITE {endnote_xml}"
    
    return instruction


def format_run_with_citation_awareness(run: Run, formatting_detail: str = "basic") -> Dict[str, Any]:
    """
    Enhanced run formatting that includes citation field information.
    
    This should be used instead of the standard extract_run_formatting
    when citation awareness is needed.
    """
    # First get any fields in this run
    fields = extract_fields_from_run(run)
    
    formatting = {
        "text": run.text,
    }
    
    # If there are fields, include field information
    if fields:
        formatting["fields"] = fields
        # Use the field content as the display text
        if fields[0].get('content'):
            formatting["display_text"] = fields[0]['content']
    
    # Add standard formatting based on detail level
    if formatting_detail in ["basic", "detailed", "comprehensive"]:
        formatting.update({
            "bold": run.bold,
            "italic": run.italic,
            "underline": run.underline,
        })
    
    if formatting_detail in ["detailed", "comprehensive"]:
        formatting.update({
            "font_name": run.font.name,
            "font_size": str(run.font.size) if run.font.size else None,
            "font_color": str(run.font.color.rgb) if run.font.color.rgb else None,
            "highlight_color": str(run.font.highlight_color) if run.font.highlight_color else None,
            "strike": run.font.strike,
            "double_strike": run.font.double_strike,
            "superscript": run.font.superscript,
            "subscript": run.font.subscript,
            "small_caps": run.font.small_caps,
            "all_caps": run.font.all_caps,
        })
    
    # Clean up None values
    return {k: v for k, v in formatting.items() if v is not None}