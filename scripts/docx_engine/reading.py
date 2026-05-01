"""Content reading: outline, full_text, read_section, read_range, search."""
import re
import json
import zipfile
import io
from typing import List, Dict, Any, Optional, Union
from lxml import etree
from docx_engine.core import load_document, NSMAP
from docx_engine import errors

# Compact formatting codes
def _run_fmt_compact(run: Any) -> str:
    parts = []
    if run.bold: parts.append("B")
    if run.italic: parts.append("I")
    if run.underline: parts.append("U")
    if run.font.strike: parts.append("S")
    if run.font.size: parts.append(f"{int(run.font.size.pt)}pt")
    if run.font.color and run.font.color.rgb:
        parts.append(f"#{run.font.color.rgb}")
    return ",".join(parts) if parts else ""

def _run_fmt_verbose(run: Any) -> str:
    parts = []
    if run.bold: parts.append("B")
    if run.italic: parts.append("I")
    if run.underline: parts.append("U")
    if run.font.strike: parts.append("S")
    if run.font.name: parts.append(f"font:{run.font.name}")
    if run.font.size: parts.append(f"size:{run.font.size.pt}pt")
    if run.font.color and run.font.color.rgb:
        parts.append(f"color:#{run.font.color.rgb}")
    return ",".join(parts) if parts else "plain"


def get_outline(doc_path: str, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    doc, err = load_document(doc_path)
    if err:
        return err
    headings = []
    lines = []
    for i, para in enumerate(doc.paragraphs):
        if para.style.name.startswith('Heading'):
            level_str = para.style.name.replace('Heading ', '').strip()
            if level_str.isdigit():
                level = int(level_str)
                indent = "  " * (level - 1)
                text = para.text.strip()
                headings.append({"index": i, "level": level, "text": text})
                lines.append(f"{indent}[P{i}] (H{level}) {text}")
    if not headings:
        if json_mode:
            return {"status": "WARNING", "headings": [], "message": "No standard 'Heading' styles found."}
        return errors.warn("reading", "get_outline", "No standard 'Heading' styles found.")
    if json_mode:
        return {"status": "SUCCESS", "headings": headings}
    return "=== DOCUMENT STRUCTURE ===\n" + "\n".join(lines)


def full_text(doc_path: str, include_formatting: bool = False, compact: bool = False, max_chars: Optional[int] = None, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    if not include_formatting:
        # LXML fast path for massive speedup on large docs
        try:
            with zipfile.ZipFile(doc_path, 'r') as z:
                xml_data = z.read('word/document.xml')
            context = etree.iterparse(io.BytesIO(xml_data), events=('end',), tag='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            
            paragraphs_data = []
            lines = []
            if not json_mode:
                lines = ["=== FULL TEXT ==="]
            
            total_chars = 0
            truncated = False
            
            for i, (_, elem) in enumerate(context):
                style = "Normal"
                pPr = elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                if pPr is not None:
                    pStyle = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                    if pStyle is not None:
                        style = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or "Normal"
                
                texts = elem.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                text = "".join([t.text for t in texts if t.text])
                elem.clear() # Memory optimization
                
                if compact and not text.strip():
                    continue
                
                if not compact and not json_mode and not text.strip():
                    lines.append(f"[P{i}] ({style}) <empty>")
                    continue
                
                if json_mode:
                    paragraphs_data.append({"index": i, "style": style, "text": text})
                else:
                    if compact:
                        prefix = f"H{style.replace('Heading','').strip()}" if style.startswith('Heading') else ""
                        line = f"[P{i}]{' '+prefix if prefix else ''} {text}"
                    else:
                        line = f"[P{i}] ({style}) {line_content if 'line_content' in locals() else text}"
                        # Fixed line above to use 'text'
                        line = f"[P{i}] ({style}) {text}"
                    total_chars += len(line)
                    if max_chars and total_chars > max_chars:
                        lines.append(f"[TRUNCATED at {max_chars} chars; use read_range --start {i} to continue]")
                        truncated = True
                        break
                    lines.append(line)
                    
            if json_mode:
                return {"status": "SUCCESS", "paragraphs": paragraphs_data, "truncated": truncated}
            return "\n".join(lines)
        except Exception:
            pass # Fall back to python-docx if fast path fails

    doc, err = load_document(doc_path)
    if err:
        return err

    paragraphs_data = []
    lines = []

    if not json_mode:
        lines = ["=== FULL TEXT ==="]

    total_chars = 0
    truncated = False

    for i, para in enumerate(doc.paragraphs):
        style = para.style.name
        text = para.text

        # Compact: skip empties
        if compact and not text.strip():
            continue

        if not compact and not json_mode and not text.strip() and not include_formatting:
            lines.append(f"[P{i}] ({style}) <empty>")
            continue

        if include_formatting:
            if compact:
                runs_info = []
                for r in para.runs:
                    fmt = _run_fmt_compact(r)
                    if fmt:
                        runs_info.append(f"[{fmt}]{r.text}")
                    else:
                        runs_info.append(r.text)
                line_content = "".join(runs_info)
            else:
                runs_info = []
                for r in para.runs:
                    fmt = _run_fmt_verbose(r)
                    runs_info.append(f'[{fmt}]"{r.text}"')
                line_content = "".join(runs_info)

            if json_mode:
                paragraphs_data.append({"index": i, "style": style, "runs": line_content})
            else:
                line = f"[P{i}] ({style}) {line_content}"
                total_chars += len(line)
                if max_chars and total_chars > max_chars:
                    lines.append(f"[TRUNCATED at {max_chars} chars; use read_range --start {i} to continue]")
                    truncated = True
                    break
                lines.append(line)
        else:
            if json_mode:
                paragraphs_data.append({"index": i, "style": style, "text": text})
            else:
                # Compact: don't print heading style name separately as it's redundant
                if compact:
                    prefix = f"H{style.replace('Heading ','').strip()}" if style.startswith('Heading') else ""
                    line = f"[P{i}]{' '+prefix if prefix else ''} {text}"
                else:
                    line = f"[P{i}] ({style}) {text}"
                total_chars += len(line)
                if max_chars and total_chars > max_chars:
                    lines.append(f"[TRUNCATED at {max_chars} chars; use read_range --start {i} to continue]")
                    truncated = True
                    break
                lines.append(line)

    if json_mode:
        return {"status": "SUCCESS", "paragraphs": paragraphs_data, "truncated": truncated}
    return "\n".join(lines)


def read_section(doc_path: str, target_heading: str, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    doc, err = load_document(doc_path)
    if err:
        return err
    content = []
    paragraphs_data = []
    capture = False
    target_level = 0
    for i, para in enumerate(doc.paragraphs):
        is_heading = para.style.name.startswith('Heading')
        current_level = 99
        if is_heading:
            level_str = para.style.name.replace('Heading ', '').strip()
            if level_str.isdigit():
                current_level = int(level_str)
        if capture:
            if is_heading and current_level <= target_level:
                break
            content.append(f"[P{i}] ({para.style.name}) {para.text}")
            paragraphs_data.append({"index": i, "style": para.style.name, "text": para.text})
        else:
            if is_heading and target_heading.lower() in para.text.lower():
                capture = True
                target_level = current_level
                content.append(f"--- SECTION: {para.text.strip()} (P{i}) ---")
                paragraphs_data.append({"index": i, "style": para.style.name, "text": para.text, "is_heading": True})
    if not capture:
        reason = f"Heading '{target_heading}' not found."
        if json_mode:
            return {"status": "ERROR", "message": errors.err("reading", "read_section", reason)}
        return errors.err("reading", "read_section", reason)
    if json_mode:
        return {"status": "SUCCESS", "section": target_heading, "paragraphs": paragraphs_data}
    return "\n".join(content)


def read_range(doc_path: str, start_idx: int, end_idx: int, include_formatting: bool = False, compact: bool = False, max_chars: Optional[int] = None, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    doc, err = load_document(doc_path)
    if err:
        return err
    paras = doc.paragraphs
    # Fix: valid indices are 0 to len-1
    if start_idx < 0 or end_idx > len(paras) - 1:
        reason = f"Invalid range. Document contains {len(paras)} paragraphs (valid indices: 0–{len(paras)-1})."
        if json_mode:
            return {"status": "ERROR", "message": errors.err("reading", "read_range", reason)}
        return errors.err("reading", "read_range", reason)

    paragraphs_data = []
    lines = [f"=== PARAGRAPH RANGE [{start_idx}-{end_idx}] ==="]
    total_chars = 0
    truncated = False

    for i in range(start_idx, end_idx + 1):
        para = paras[i]
        if include_formatting:
            if compact:
                runs_info = []
                for r in para.runs:
                    fmt = _run_fmt_compact(r)
                    runs_info.append(f"[{fmt}]{r.text}" if fmt else r.text)
                line_content = "".join(runs_info)
            else:
                runs_info = []
                for r in para.runs:
                    fmt = _run_fmt_verbose(r)
                    runs_info.append(f'[{fmt}]"{r.text}"')
                line_content = "".join(runs_info)
            line = f"[P{i}] ({para.style.name}) {line_content}"
            paragraphs_data.append({"index": i, "style": para.style.name, "runs": line_content})
        else:
            line = f"[P{i}] ({para.style.name}) {para.text}"
            paragraphs_data.append({"index": i, "style": para.style.name, "text": para.text})

        total_chars += len(line)
        if max_chars and total_chars > max_chars:
            lines.append(f"[TRUNCATED at {max_chars} chars; use --start {i} to continue]")
            truncated = True
            break
        lines.append(line)

    if json_mode:
        return {"status": "SUCCESS", "range": [start_idx, end_idx], "paragraphs": paragraphs_data, "truncated": truncated}
    return "\n".join(lines)


def search_text(doc_path: str, query: str, context_lines: int = 1, compact: bool = False, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    # LXML fast path for search
    try:
        with zipfile.ZipFile(doc_path, 'r') as z:
            xml_data = z.read('word/document.xml')
        context = etree.iterparse(io.BytesIO(xml_data), events=('end',), tag='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        
        all_texts = []
        for _, elem in context:
            texts = elem.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            p_text = "".join([t.text for t in texts if t.text])
            all_texts.append(p_text)
            elem.clear()
            
        matches_data = []
        results = []
        for i, text in enumerate(all_texts):
            if re.search(query, text, re.IGNORECASE):
                ctx_start = max(0, i - context_lines)
                ctx_end = min(len(all_texts) - 1, i + context_lines)
                block = []
                context_before = []
                context_after = []
                for j in range(ctx_start, ctx_end + 1):
                    marker = ">>>" if j == i else "   "
                    block.append(f"{marker} [P{j}] {all_texts[j]}")
                    if j < i:
                        context_before.append({"index": j, "text": all_texts[j]})
                    elif j > i:
                        context_after.append({"index": j, "text": all_texts[j]})
                matches_data.append({
                    "para_index": i,
                    "text": text,
                    "context_before": context_before,
                    "context_after": context_after,
                })
                if not compact:
                    results.append("\n".join(block))
                else:
                    results.append(f"[P{i}] {text}")

        if not matches_data:
            if json_mode:
                return {"status": "SUCCESS", "query": query, "match_count": 0, "matches": []}
            return f"'{query}' not found."

        if json_mode:
            return {"status": "SUCCESS", "query": query, "match_count": len(matches_data), "matches": matches_data}

        sep = "\n---\n" if not compact else "\n"
        return (f"=== SEARCH: '{query}' ({len(results)} matches) ===\n\n"
                + sep.join(results))
    except Exception:
        pass # Fall back to python-docx

    doc, err = load_document(doc_path)
    if err:
        return err

    matches_data = []
    results = []
    paras = doc.paragraphs
    for i, para in enumerate(paras):
        if re.search(query, para.text, re.IGNORECASE):
            ctx_start = max(0, i - context_lines)
            ctx_end = min(len(paras) - 1, i + context_lines)
            block = []
            context_before = []
            context_after = []
            for j in range(ctx_start, ctx_end + 1):
                marker = ">>>" if j == i else "   "
                block.append(f"{marker} [P{j}] {paras[j].text}")
                if j < i:
                    context_before.append({"index": j, "text": paras[j].text})
                elif j > i:
                    context_after.append({"index": j, "text": paras[j].text})
            matches_data.append({
                "para_index": i,
                "text": para.text,
                "context_before": context_before,
                "context_after": context_after,
            })
            if not compact:
                results.append("\n".join(block))
            else:
                results.append(f"[P{i}] {para.text}")

    if not matches_data:
        if json_mode:
            return {"status": "SUCCESS", "query": query, "match_count": 0, "matches": []}
        return f"'{query}' not found."

    if json_mode:
        return {"status": "SUCCESS", "query": query, "match_count": len(matches_data), "matches": matches_data}

    sep = "\n---\n" if not compact else "\n"
    return (f"=== SEARCH: '{query}' ({len(results)} matches) ===\n\n"
            + sep.join(results))
