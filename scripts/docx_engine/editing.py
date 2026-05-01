"""Text editing: insert, delete, replace, append paragraphs and headings."""
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx_engine.core import load_document, save_document


def insert_paragraph(doc_path, text, index=None, after_heading=None, style=None, output_path=None):
    doc, err = load_document(doc_path)
    if err:
        return err

    if after_heading:
        for i, para in enumerate(doc.paragraphs):
            if para.style.name.startswith('Heading') and after_heading.lower() in para.text.lower():
                new_p = doc.add_paragraph(text, style=style)
                para._p.addnext(new_p._p)
                return save_document(doc, doc_path, output_path) + f"\nText inserted after heading '{after_heading}'."
        return f"ERROR: Heading '{after_heading}' not found."

    if index is not None:
        paras = doc.paragraphs
        if index < 0 or index >= len(paras):
            return f"ERROR: Invalid index. Document contains {len(paras)} paragraphs."
        new_p = doc.add_paragraph(text, style=style)
        paras[index]._p.addnext(new_p._p)
        return save_document(doc, doc_path, output_path) + f"\nText inserted after P{index}."

    doc.add_paragraph(text, style=style)
    return save_document(doc, doc_path, output_path) + "\nText appended to the end of the document."


def insert_heading(doc_path, text, level=1, index=None, output_path=None):
    doc, err = load_document(doc_path)
    if err:
        return err

    if index is not None:
        paras = doc.paragraphs
        if index < 0 or index >= len(paras):
            return f"ERROR: Invalid index."
        new_p = doc.add_heading(text, level=level)
        paras[index]._p.addnext(new_p._p)
    else:
        doc.add_heading(text, level=level)

    return save_document(doc, doc_path, output_path) + f"\nH{level} heading added: '{text}'"


def delete_paragraphs(doc_path, indices, output_path=None):
    doc, err = load_document(doc_path)
    if err:
        return err

    paras = doc.paragraphs
    # Convert string indices "1,2,3" to list if needed
    if isinstance(indices, str):
        indices = [int(x.strip()) for x in indices.split(',')]
        
    sorted_indices = sorted(indices, reverse=True)
    deleted = []
    for idx in sorted_indices:
        if 0 <= idx < len(paras):
            p = paras[idx]._p
            p.getparent().remove(p)
            deleted.append(idx)

    if not deleted:
        return "ERROR: No paragraphs were deleted."
    return save_document(doc, doc_path, output_path) + f"\nDeleted paragraphs: {deleted}"


def replace_text(doc_path, find_pattern, replace_with, use_regex=False, output_path=None):
    doc, err = load_document(doc_path)
    if err:
        return err

    count = 0
    # Search in paragraphs
    for para in doc.paragraphs:
        count += _surgical_replace(para, find_pattern, replace_with, use_regex)

    # Also search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    count += _surgical_replace(para, find_pattern, replace_with, use_regex)

    if count == 0:
        return f"'{find_pattern}' not found, no changes made."
    return save_document(doc, doc_path, output_path) + f"\nReplaced in {count} occurrences."


def _surgical_replace(paragraph, find_str, replace_str, use_regex=False):
    """
    Ultimate Surgical Replace: Preserves formatting by mapping text to runs.
    This handles cases where the find_str is split across multiple runs.
    """
    total_replaced = 0
    
    # 1. Get full text and map each character to its run
    full_text = ""
    char_map = [] # list of (run_index, char_index_in_run)
    
    for r_idx, run in enumerate(paragraph.runs):
        for c_idx, char in enumerate(run.text):
            full_text += char
            char_map.append((r_idx, c_idx))
            
    # 2. Find occurrences
    matches = []
    if use_regex:
        for m in re.finditer(find_str, full_text):
            matches.append((m.start(), m.end()))
    else:
        start = 0
        while True:
            idx = full_text.find(find_str, start)
            if idx == -1: break
            matches.append((idx, idx + len(find_str)))
            start = idx + len(find_str)
            
    if not matches:
        return 0
        
    # 3. Replace from end to beginning to keep indices valid
    for start, end in reversed(matches):
        total_replaced += 1
        
        # Identify runs involved
        start_run_idx, _ = char_map[start]
        end_run_idx, _ = char_map[end - 1]
        
        # Simple Case: Match is within a single run
        if start_run_idx == end_run_idx:
            run = paragraph.runs[start_run_idx]
            match_in_run = full_text[start:end]
            run.text = run.text.replace(match_in_run, replace_str, 1)
        else:
            # Complex Case: Match spans multiple runs
            # Strategy: Put the replacement in the first run, clear the rest of the match in others
            
            # Start run: Keep text before match, append replacement
            first_run = paragraph.runs[start_run_idx]
            _, char_idx_in_first = char_map[start]
            prefix = first_run.text[:char_idx_in_first]
            first_run.text = prefix + replace_str
            
            # Middle runs: Clear completely
            for r_idx in range(start_run_idx + 1, end_run_idx):
                paragraph.runs[r_idx].text = ""
                
            # End run: Keep text after match
            last_run = paragraph.runs[end_run_idx]
            _, char_idx_in_last = char_map[end - 1]
            last_run.text = last_run.text[char_idx_in_last + 1:]
            
    return total_replaced


def append_text(doc_path, text, style=None, output_path=None):
    doc, err = load_document(doc_path)
    if err:
        return err
    doc.add_paragraph(text, style=style)
    return save_document(doc, doc_path, output_path) + "\nText appended to the end of the document."
