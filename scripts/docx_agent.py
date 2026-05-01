#!/usr/bin/env python3
"""
Ultimate DOCX Agent — Full MS Office capabilities for LLM agents.
A comprehensive command-line tool for reading, writing, and formatting .docx files.
"""
import argparse
import sys
import json
import os
import shutil
import hashlib
from datetime import datetime
from docx_engine import (
    core, reading, editing, formatting, tables, annotations,
    rendering, smart_features, batch_tools, extended, advanced
)

# Fix Unicode encoding on Windows consoles
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

TEMP_FILES = []

def register_temp(path):
    if path and os.path.exists(path):
        TEMP_FILES.append(path)

def cleanup():
    for path in TEMP_FILES:
        try:
            if os.path.isfile(path): os.remove(path)
            elif os.path.isdir(path): shutil.rmtree(path)
        except Exception:
            pass

def _str2bool(v):
    if isinstance(v, bool): return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'): return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'): return False
    else: raise argparse.ArgumentTypeError('Boolean value expected.')

def _parse_indices(s):
    """Parse comma-separated int indices like '1,2,3' into a list."""
    return [int(x.strip()) for x in s.split(',') if x.strip()]

def _file_hash(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        h.update(f.read())
    return h.hexdigest()

def _emit(result, use_json, command_name=None, file_path=None):
    """Emit result to stdout — either as JSON envelope or plain text."""
    if use_json:
        if isinstance(result, dict):
            # Already structured — ensure envelope fields exist
            out = result
            if "command" not in out:
                out["command"] = command_name or ""
            if "file" not in out:
                out["file"] = file_path or ""
        else:
            text = str(result)
            status = "ERROR" if text.startswith("ERROR") else ("WARNING" if text.startswith("WARNING") else "SUCCESS")
            out = {
                "status": status,
                "command": command_name or "",
                "file": file_path or "",
                "message": text,
            }
        print(json.dumps(out, indent=2, ensure_ascii=False))
    else:
        print(result)

def main():
    parser = argparse.ArgumentParser(
        description="Ultimate DOCX Agent — Full MS Office capabilities for LLMs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("file", help="Path to the .docx file to process.")
    parser.add_argument("--json", action="store_true", help="Output result in JSON format.")
    parser.add_argument("--backup", action="store_true", help="Create a backup before write operations.")
    parser.add_argument("--cleanup", action="store_true", help="Auto-delete temp files after execution.")
    parser.add_argument("--verify", action="store_true", help="After write ops, read back affected range to confirm.")

    sub = parser.add_subparsers(dest="action", help="Action to perform")

    # ===== DOCUMENT LIFECYCLE =====
    sub.add_parser("create", help="Create a new blank document.")
    sub.add_parser("info", help="Show document metadata and statistics.")

    p = sub.add_parser("set_metadata", help="Set document metadata fields.")
    p.add_argument("--title"); p.add_argument("--author"); p.add_argument("--subject")
    p.add_argument("--keywords"); p.add_argument("--category"); p.add_argument("--output")

    # ===== READING =====
    sub.add_parser("outline", help="Show heading hierarchy with [P_index] labels.")

    p = sub.add_parser("full_text", help="Read entire document text with paragraph indices.")
    p.add_argument("--formatting", action="store_true", help="Include run-level formatting detail.")
    p.add_argument("--compact", action="store_true", help="Skip empty paragraphs; short format codes.")
    p.add_argument("--max-chars", type=int, default=None, dest="max_chars", help="Truncate output at N characters.")

    p = sub.add_parser("read_section", help="Read all paragraphs under a specific heading.")
    p.add_argument("--heading", required=True)

    p = sub.add_parser("read_range", help="Read a range of paragraphs by index.")
    p.add_argument("--start", type=int, required=True)
    p.add_argument("--end", type=int, required=True)
    p.add_argument("--formatting", action="store_true")
    p.add_argument("--compact", action="store_true")
    p.add_argument("--max-chars", type=int, default=None, dest="max_chars")

    p = sub.add_parser("search", help="Search for text (regex-capable) with context lines.")
    p.add_argument("--query", required=True)
    p.add_argument("--context", type=int, default=1)
    p.add_argument("--compact", action="store_true")

    p = sub.add_parser("read_page", help="Read content from an approximate page number.")
    p.add_argument("--page", type=int, required=True)

    sub.add_parser("summary_map", help="Compact map of entire document — best first command for unknown docs.")
    sub.add_parser("word_count", help="Word count broken down by section/heading.")
    sub.add_parser("detect_language", help="Detect primary language of the document.")

    # ===== TABLES (READING) =====
    sub.add_parser("list_tables", help="List all tables with [T_index], size, and preview.")

    p = sub.add_parser("read_table", help="Read table content as Markdown + cell coordinates.")
    p.add_argument("--index", type=int, required=True)

    # ===== EDITING =====
    p = sub.add_parser("insert_paragraph", help="Insert a paragraph at an index or after a heading.")
    p.add_argument("--text", required=True); p.add_argument("--index", type=int)
    p.add_argument("--after-heading", dest="after_heading"); p.add_argument("--style"); p.add_argument("--output")

    p = sub.add_parser("insert_heading", help="Insert a heading at level 1–9.")
    p.add_argument("--text", required=True); p.add_argument("--level", type=int, default=1)
    p.add_argument("--index", type=int); p.add_argument("--output")

    p = sub.add_parser("delete_paragraphs", help="Delete paragraphs by comma-separated indices.")
    p.add_argument("--indices", required=True); p.add_argument("--output")

    p = sub.add_parser("replace_text", help="Surgical find-and-replace preserving XML formatting.")
    p.add_argument("--find", required=True); p.add_argument("--replace", required=True)
    p.add_argument("--regex", action="store_true"); p.add_argument("--output")

    p = sub.add_parser("append_text", help="Append text to the end of the document.")
    p.add_argument("--text", required=True); p.add_argument("--style"); p.add_argument("--output")

    # ===== FORMATTING =====
    p = sub.add_parser("format_text", help="Apply character formatting (bold, italic, font, color…).")
    p.add_argument("--match"); p.add_argument("--para-indices", dest="para_indices")
    p.add_argument("--bold", type=_str2bool); p.add_argument("--italic", type=_str2bool)
    p.add_argument("--underline", type=_str2bool); p.add_argument("--strike", type=_str2bool)
    p.add_argument("--font-name", dest="font_name"); p.add_argument("--font-size", type=float, dest="font_size")
    p.add_argument("--font-color", dest="font_color"); p.add_argument("--highlight")
    p.add_argument("--all-caps", type=_str2bool, dest="all_caps")
    p.add_argument("--small-caps", type=_str2bool, dest="small_caps")
    p.add_argument("--superscript", type=_str2bool); p.add_argument("--subscript", type=_str2bool)
    p.add_argument("--output")

    p = sub.add_parser("format_paragraph", help="Apply paragraph formatting (alignment, spacing, indent…).")
    p.add_argument("--para-indices", required=True, dest="para_indices")
    p.add_argument("--alignment"); p.add_argument("--line-spacing", type=float, dest="line_spacing")
    p.add_argument("--space-before", type=float, dest="space_before")
    p.add_argument("--space-after", type=float, dest="space_after")
    p.add_argument("--left-indent", type=float, dest="left_indent")
    p.add_argument("--right-indent", type=float, dest="right_indent")
    p.add_argument("--first-line-indent", type=float, dest="first_line_indent")
    p.add_argument("--keep-together", type=_str2bool, dest="keep_together")
    p.add_argument("--keep-with-next", type=_str2bool, dest="keep_with_next")
    p.add_argument("--page-break-before", type=_str2bool, dest="page_break_before")
    p.add_argument("--output")

    p = sub.add_parser("list_styles", help="List all paragraph/character/table styles.")
    p.add_argument("--type", dest="style_type", choices=["paragraph", "character", "table", "list"],
                   help="Filter by style type.")

    p = sub.add_parser("apply_style", help="Apply a named style to paragraphs by index.")
    p.add_argument("--para-indices", required=True, dest="para_indices")
    p.add_argument("--style", required=True)
    p.add_argument("--output")

    p = sub.add_parser("clone_format", help="Format Painter: copy formatting from one para to others.")
    p.add_argument("--source-para", type=int, required=True, dest="source_para")
    p.add_argument("--target-paras", required=True, dest="target_paras")
    p.add_argument("--output")

    # ===== TABLE OPERATIONS =====
    p = sub.add_parser("create_table", help="Create a table with headers and optional data rows.")
    p.add_argument("--headers", required=True, help="Comma-separated column headers.")
    p.add_argument("--rows-data", dest="rows_data", help="Semicolon-separated rows; columns comma-separated.")
    p.add_argument("--style"); p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    p = sub.add_parser("modify_cell", help="Set text in a specific table cell [row,col].")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--row", type=int, required=True); p.add_argument("--col", type=int, required=True)
    p.add_argument("--text", required=True); p.add_argument("--output")

    p = sub.add_parser("add_row", help="Add a row to a table.")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--values", help="Comma-separated cell values."); p.add_argument("--output")

    p = sub.add_parser("delete_row", help="Delete a row from a table.")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--row", type=int, required=True); p.add_argument("--output")

    p = sub.add_parser("add_column", help="Add a column to a table.")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--header", default=""); p.add_argument("--output")

    p = sub.add_parser("delete_column", help="Delete a column from a table.")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--col", type=int, required=True); p.add_argument("--output")

    p = sub.add_parser("merge_cells", help="Merge a rectangular range of cells in a table.")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--start-row", type=int, required=True, dest="start_row")
    p.add_argument("--start-col", type=int, required=True, dest="start_col")
    p.add_argument("--end-row", type=int, required=True, dest="end_row")
    p.add_argument("--end-col", type=int, required=True, dest="end_col")
    p.add_argument("--output")

    p = sub.add_parser("format_table_cell", help="Format a table cell (bg color, bold, alignment, font size).")
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--row", type=int, required=True); p.add_argument("--col", type=int, required=True)
    p.add_argument("--bg-color", dest="bg_color"); p.add_argument("--bold", type=_str2bool)
    p.add_argument("--alignment"); p.add_argument("--font-size", type=float, dest="font_size")
    p.add_argument("--output")

    # ===== ANNOTATIONS =====
    sub.add_parser("read_comments", help="Read all comments, authors, dates, and referenced text.")
    sub.add_parser("read_tracked_changes", help="List all insertions, deletions, and format changes.")
    sub.add_parser("accept_all_changes", help="Accept all tracked changes in the document.")

    p = sub.add_parser("reject_all_changes", help="Reject all tracked changes in the document.")
    p.add_argument("--output")

    p = sub.add_parser("add_comment", help="Add a review comment to a paragraph.")
    p.add_argument("--para", type=int, required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--author", default="LLM Agent"); p.add_argument("--output")

    p = sub.add_parser("delete_comment", help="Delete a comment by its ID.")
    p.add_argument("--id", type=int, required=True, dest="comment_id"); p.add_argument("--output")

    sub.add_parser("read_footnotes", help="Read all footnotes.")
    sub.add_parser("read_endnotes", help="Read all endnotes.")

    p = sub.add_parser("add_footnote", help="Add a footnote to a paragraph.")
    p.add_argument("--para", type=int, required=True)
    p.add_argument("--text", required=True); p.add_argument("--output")

    sub.add_parser("list_bookmarks", help="List all bookmarks in the document.")

    p = sub.add_parser("add_bookmark", help="Add a named bookmark to a paragraph.")
    p.add_argument("--para", type=int, required=True)
    p.add_argument("--name", required=True); p.add_argument("--output")

    sub.add_parser("read_textboxes", help="Read all text boxes and shapes.")

    # ===== ADVANCED / IMAGES / LAYOUT =====
    p = sub.add_parser("insert_image", help="Insert an image into the document.")
    p.add_argument("--image", required=True, dest="image_path")
    p.add_argument("--width", type=float); p.add_argument("--height", type=float)
    p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    sub.add_parser("list_images", help="List all images embedded in the document.")

    p = sub.add_parser("set_header", help="Set the header text for a section.")
    p.add_argument("--text", required=True)
    p.add_argument("--section", type=int, default=0, dest="section_index"); p.add_argument("--output")

    p = sub.add_parser("set_footer", help="Set the footer text (optionally with page number).")
    p.add_argument("--text", required=True)
    p.add_argument("--page-number", action="store_true", dest="add_page_number")
    p.add_argument("--section", type=int, default=0, dest="section_index"); p.add_argument("--output")

    p = sub.add_parser("read_header", help="Read header text from a section.")
    p.add_argument("--section", type=int, default=0, dest="section_index")

    p = sub.add_parser("read_footer", help="Read footer text from a section.")
    p.add_argument("--section", type=int, default=0, dest="section_index")

    p = sub.add_parser("set_margins", help="Set page margins (in cm).")
    p.add_argument("--top", type=float); p.add_argument("--bottom", type=float)
    p.add_argument("--left", type=float); p.add_argument("--right", type=float)
    p.add_argument("--section", type=int, default=None, dest="section_index"); p.add_argument("--output")

    p = sub.add_parser("set_orientation", help="Set page orientation (portrait or landscape).")
    p.add_argument("--orientation", required=True, choices=["portrait", "landscape"])
    p.add_argument("--section", type=int, default=None, dest="section_index"); p.add_argument("--output")

    p = sub.add_parser("set_page_size", help="Set page size by preset (a4, letter, legal…) or dimensions.")
    p.add_argument("--preset"); p.add_argument("--width", type=float); p.add_argument("--height", type=float)
    p.add_argument("--section", type=int, default=None, dest="section_index"); p.add_argument("--output")

    p = sub.add_parser("insert_page_break", help="Insert a page break after a paragraph.")
    p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    p = sub.add_parser("insert_section_break", help="Insert a section break.")
    p.add_argument("--type", default="new_page", dest="break_type",
                   choices=["new_page", "continuous", "even_page", "odd_page"])
    p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    p = sub.add_parser("insert_list", help="Insert a bullet or numbered list.")
    p.add_argument("--items", required=True, help="Semicolon-separated list items.")
    p.add_argument("--type", default="bullet", dest="list_type", choices=["bullet", "number"])
    p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    p = sub.add_parser("insert_hyperlink", help="Insert a hyperlink paragraph.")
    p.add_argument("--text", required=True); p.add_argument("--url", required=True)
    p.add_argument("--after-para", type=int, dest="after_para"); p.add_argument("--output")

    p = sub.add_parser("insert_toc", help="Insert a Table of Contents field.")
    p.add_argument("--title", default="Table of Contents"); p.add_argument("--output")

    # ===== SMART FEATURES =====
    p = sub.add_parser("from_template", help="Create a document from a template, replacing {{variables}}.")
    p.add_argument("--template", required=True, dest="template_path")
    p.add_argument("--output", required=True)
    p.add_argument("--vars", required=True, dest="variables", help="key1=val1,key2=val2")

    sub.add_parser("list_template_vars", help="Find all {{variable}} placeholders in the document.")

    # ===== EXTENDED / COMPLEX ELEMENTS =====
    sub.add_parser("read_smartart", help="Detect and describe SmartArt diagrams.")
    sub.add_parser("read_charts", help="Identify chart types, titles, and data labels.")
    sub.add_parser("list_embedded", help="List all embedded OLE objects.")
    sub.add_parser("read_content_controls", help="Read all content controls / form fields.")
    sub.add_parser("read_protection", help="Check document protection settings.")
    sub.add_parser("full_statistics", help="Comprehensive document statistics (style/font usage, etc.).")

    p = sub.add_parser("add_watermark", help="Add a diagonal text watermark to all pages.")
    p.add_argument("--text", required=True)
    p.add_argument("--font-size", type=int, default=72, dest="font_size")
    p.add_argument("--color", default="#C0C0C0"); p.add_argument("--output")

    p = sub.add_parser("remove_watermark", help="Remove watermark from the document.")
    p.add_argument("--output")

    # ===== LAYOUT & RENDERING =====
    p = sub.add_parser("render", help="Convert pages to PNG images (LibreOffice required).")
    p.add_argument("--pages", help="Comma-separated page numbers, e.g. '1,2,3'.")
    p.add_argument("--output-dir", dest="output_dir")
    p.add_argument("--dpi", type=int, default=200)

    sub.add_parser("describe_layout", help="Text-based map of margins, page size, and element locations.")

    # ===== BACKUP / UNDO =====
    sub.add_parser("undo", help="Restore document from the most recent backup.")
    sub.add_parser("list_backups", help="List all available backups for this document.")

    # ===== BATCH & DIFF =====
    p = sub.add_parser("batch", help="Execute multiple commands from a JSON file (atomic, in-process).")
    p.add_argument("--commands", required=True, dest="commands_json_path", help="Path to commands JSON file.")
    p.add_argument("--dry-run", action="store_true", dest="dry_run")

    p = sub.add_parser("diff", help="Compare this document to another DOCX file.")
    p.add_argument("--compare", required=True, dest="compare_path")

    args = parser.parse_args()

    if not args.action:
        parser.print_help()
        sys.exit(0)

    READ_ONLY_ACTIONS = {
        'info', 'outline', 'full_text', 'read_section', 'read_range', 'search',
        'list_tables', 'read_table', 'read_comments', 'read_tracked_changes',
        'read_footnotes', 'read_endnotes', 'list_bookmarks', 'read_textboxes',
        'list_images', 'read_header', 'read_footer', 'read_protection',
        'read_smartart', 'read_charts', 'list_embedded', 'read_content_controls',
        'full_statistics', 'describe_layout', 'summary_map', 'word_count',
        'detect_language', 'list_styles', 'list_backups', 'list_template_vars',
        'read_page', 'diff',
    }

    use_json = args.json
    file_path = args.file
    action = args.action

    # Handle Backup (lazy: only if file changed)
    pre_hash = None
    if args.backup and action not in READ_ONLY_ACTIONS:
        if os.path.exists(file_path):
            pre_hash = _file_hash(file_path)

    try:
        result = "No action performed."

        # LIFECYCLE
        if action == "create":
            result = core.create_document(file_path)
        elif action == "info":
            result = core.get_info(file_path, json_mode=use_json)
        elif action == "set_metadata":
            result = core.set_metadata(file_path, output_path=getattr(args, 'output', None),
                                       title=args.title, author=args.author, subject=args.subject,
                                       keywords=args.keywords, category=args.category)

        # READING
        elif action == "outline":
            result = reading.get_outline(file_path, json_mode=use_json)
        elif action == "full_text":
            result = reading.full_text(file_path, include_formatting=args.formatting,
                                       compact=args.compact, max_chars=args.max_chars,
                                       json_mode=use_json)
        elif action == "read_section":
            result = reading.read_section(file_path, args.heading, json_mode=use_json)
        elif action == "read_range":
            result = reading.read_range(file_path, args.start, args.end,
                                        include_formatting=args.formatting,
                                        compact=args.compact, max_chars=args.max_chars,
                                        json_mode=use_json)
        elif action == "search":
            result = reading.search_text(file_path, args.query, context_lines=args.context,
                                         compact=args.compact, json_mode=use_json)
        elif action == "read_page":
            result = smart_features.read_page(file_path, args.page)
        elif action == "summary_map":
            result = smart_features.summary_map(file_path)
        elif action == "word_count":
            result = smart_features.word_count(file_path)
        elif action == "detect_language":
            result = smart_features.detect_language(file_path)
        elif action == "list_template_vars":
            result = smart_features.list_template_variables(file_path)

        # TABLES (read)
        elif action == "list_tables":
            result = tables.list_tables(file_path, json_mode=use_json)
        elif action == "read_table":
            result = tables.read_table(file_path, args.index, json_mode=use_json)

        # EDITING
        elif action == "insert_paragraph":
            result = editing.insert_paragraph(file_path, args.text, index=args.index,
                                              after_heading=args.after_heading,
                                              style=args.style, output_path=getattr(args, 'output', None))
        elif action == "insert_heading":
            result = editing.insert_heading(file_path, args.text, level=args.level,
                                            index=args.index, output_path=getattr(args, 'output', None))
        elif action == "delete_paragraphs":
            result = editing.delete_paragraphs(file_path, args.indices,
                                               output_path=getattr(args, 'output', None))
        elif action == "replace_text":
            result = editing.replace_text(file_path, args.find, args.replace,
                                          use_regex=args.regex,
                                          output_path=getattr(args, 'output', None))
        elif action == "append_text":
            result = editing.append_text(file_path, args.text, style=args.style,
                                         output_path=getattr(args, 'output', None))

        # FORMATTING
        elif action == "format_text":
            idx = _parse_indices(args.para_indices) if args.para_indices else None
            result = formatting.format_text(file_path, match=args.match, para_indices=idx,
                                            bold=args.bold, italic=args.italic,
                                            underline=args.underline, strike=args.strike,
                                            font_name=args.font_name, font_size=args.font_size,
                                            font_color=args.font_color, highlight=args.highlight,
                                            all_caps=args.all_caps, small_caps=args.small_caps,
                                            superscript=args.superscript, subscript=args.subscript,
                                            output_path=getattr(args, 'output', None))
        elif action == "format_paragraph":
            idx = _parse_indices(args.para_indices)
            result = formatting.format_paragraph(file_path, idx,
                                                 alignment=args.alignment,
                                                 line_spacing=args.line_spacing,
                                                 space_before=args.space_before,
                                                 space_after=args.space_after,
                                                 left_indent=args.left_indent,
                                                 right_indent=args.right_indent,
                                                 first_line_indent=args.first_line_indent,
                                                 keep_together=args.keep_together,
                                                 keep_with_next=args.keep_with_next,
                                                 page_break_before=args.page_break_before,
                                                 output_path=getattr(args, 'output', None))
        elif action == "list_styles":
            result = formatting.list_styles(file_path, style_type=args.style_type, json_mode=use_json)
        elif action == "apply_style":
            idx = _parse_indices(args.para_indices)
            result = formatting.apply_style(file_path, idx, args.style,
                                            output_path=getattr(args, 'output', None))
        elif action == "clone_format":
            target_idx = _parse_indices(args.target_paras)
            result = smart_features.clone_format(file_path, args.source_para, target_idx,
                                                 output_path=getattr(args, 'output', None))

        # TABLE OPERATIONS (write)
        elif action == "create_table":
            result = tables.create_table(file_path, args.headers,
                                         rows_data=args.rows_data, style=args.style,
                                         after_para=args.after_para,
                                         output_path=getattr(args, 'output', None))
        elif action == "modify_cell":
            result = tables.modify_cell(file_path, args.index, args.row, args.col, args.text,
                                        output_path=getattr(args, 'output', None))
        elif action == "add_row":
            result = tables.add_row(file_path, args.index, values=args.values,
                                    output_path=getattr(args, 'output', None))
        elif action == "delete_row":
            result = tables.delete_row(file_path, args.index, args.row,
                                       output_path=getattr(args, 'output', None))
        elif action == "add_column":
            result = tables.add_column(file_path, args.index, header=args.header,
                                       output_path=getattr(args, 'output', None))
        elif action == "delete_column":
            result = tables.delete_column(file_path, args.index, args.col,
                                          output_path=getattr(args, 'output', None))
        elif action == "merge_cells":
            result = tables.merge_cells(file_path, args.index,
                                        args.start_row, args.start_col,
                                        args.end_row, args.end_col,
                                        output_path=getattr(args, 'output', None))
        elif action == "format_table_cell":
            result = tables.format_table_cell(file_path, args.index, args.row, args.col,
                                              bg_color=args.bg_color, bold=args.bold,
                                              alignment=args.alignment, font_size=args.font_size,
                                              output_path=getattr(args, 'output', None))

        # ANNOTATIONS
        elif action == "read_comments":
            result = annotations.read_comments(file_path)
        elif action == "add_comment":
            result = annotations.add_comment(file_path, args.para, args.text,
                                             author=args.author,
                                             output_path=getattr(args, 'output', None))
        elif action == "delete_comment":
            result = annotations.delete_comment(file_path, args.comment_id,
                                                output_path=getattr(args, 'output', None))
        elif action == "read_tracked_changes":
            result = annotations.read_tracked_changes(file_path)
        elif action == "accept_all_changes":
            result = annotations.accept_all_changes(file_path,
                                                     output_path=getattr(args, 'output', None))
        elif action == "reject_all_changes":
            result = annotations.reject_all_changes(file_path,
                                                     output_path=getattr(args, 'output', None))
        elif action == "read_footnotes":
            result = annotations.read_footnotes(file_path)
        elif action == "read_endnotes":
            result = annotations.read_endnotes(file_path)
        elif action == "add_footnote":
            result = annotations.add_footnote(file_path, args.para, args.text,
                                              output_path=getattr(args, 'output', None))
        elif action == "list_bookmarks":
            result = annotations.list_bookmarks(file_path)
        elif action == "add_bookmark":
            result = annotations.add_bookmark(file_path, args.para, args.name,
                                              output_path=getattr(args, 'output', None))
        elif action == "read_textboxes":
            result = annotations.read_textboxes(file_path)

        # ADVANCED
        elif action == "insert_image":
            result = advanced.insert_image(file_path, args.image_path,
                                           width=args.width, height=args.height,
                                           after_para=args.after_para,
                                           output_path=getattr(args, 'output', None))
        elif action == "list_images":
            result = advanced.list_images(file_path)
        elif action == "set_header":
            result = advanced.set_header(file_path, args.text, section_index=args.section_index,
                                         output_path=getattr(args, 'output', None))
        elif action == "set_footer":
            result = advanced.set_footer(file_path, args.text,
                                         add_page_number=args.add_page_number,
                                         section_index=args.section_index,
                                         output_path=getattr(args, 'output', None))
        elif action == "read_header":
            result = advanced.read_header(file_path, section_index=args.section_index)
        elif action == "read_footer":
            result = advanced.read_footer(file_path, section_index=args.section_index)
        elif action == "set_margins":
            result = advanced.set_margins(file_path, top=args.top, bottom=args.bottom,
                                          left=args.left, right=args.right,
                                          section_index=args.section_index,
                                          output_path=getattr(args, 'output', None))
        elif action == "set_orientation":
            result = advanced.set_orientation(file_path, args.orientation,
                                              section_index=args.section_index,
                                              output_path=getattr(args, 'output', None))
        elif action == "set_page_size":
            result = advanced.set_page_size(file_path, preset=args.preset,
                                            width=args.width, height=args.height,
                                            section_index=args.section_index,
                                            output_path=getattr(args, 'output', None))
        elif action == "insert_page_break":
            result = advanced.insert_page_break(file_path, after_para=args.after_para,
                                                output_path=getattr(args, 'output', None))
        elif action == "insert_section_break":
            result = advanced.insert_section_break(file_path, break_type=args.break_type,
                                                    after_para=args.after_para,
                                                    output_path=getattr(args, 'output', None))
        elif action == "insert_list":
            result = advanced.insert_list(file_path, args.items, list_type=args.list_type,
                                          after_para=args.after_para,
                                          output_path=getattr(args, 'output', None))
        elif action == "insert_hyperlink":
            result = advanced.insert_hyperlink(file_path, args.text, args.url,
                                               after_para=args.after_para,
                                               output_path=getattr(args, 'output', None))
        elif action == "insert_toc":
            result = advanced.insert_toc(file_path, title=args.title,
                                         output_path=getattr(args, 'output', None))

        # SMART FEATURES
        elif action == "from_template":
            result = smart_features.from_template(args.template_path, args.output,
                                                   variables=args.variables)

        # EXTENDED
        elif action == "read_smartart":
            result = extended.read_smartart(file_path)
        elif action == "read_charts":
            result = extended.read_charts(file_path)
        elif action == "list_embedded":
            result = extended.list_embedded_objects(file_path)
        elif action == "read_content_controls":
            result = extended.read_content_controls(file_path)
        elif action == "read_protection":
            result = extended.read_protection(file_path)
        elif action == "full_statistics":
            result = extended.full_statistics(file_path)
        elif action == "add_watermark":
            result = extended.add_watermark(file_path, args.text,
                                            font_size=args.font_size, color=args.color,
                                            output_path=getattr(args, 'output', None))
        elif action == "remove_watermark":
            result = extended.remove_watermark(file_path,
                                               output_path=getattr(args, 'output', None))

        # LAYOUT & RENDERING
        elif action == "render":
            pages = [int(p) for p in args.pages.split(',')] if args.pages else None
            result = rendering.render_pages(file_path, pages=pages,
                                            output_dir=args.output_dir, dpi=args.dpi)
        elif action == "describe_layout":
            result = rendering.describe_layout(file_path)

        # BACKUP / UNDO
        elif action == "undo":
            result = batch_tools.undo(file_path)
        elif action == "list_backups":
            result = batch_tools.list_backups(file_path)

        # BATCH & DIFF
        elif action == "batch":
            result = batch_tools.batch_execute(file_path, args.commands_json_path,
                                               dry_run=args.dry_run, json_mode=use_json)
        elif action == "diff":
            result = batch_tools.diff_documents(file_path, args.compare_path, json_mode=use_json)

        # Handle lazy backup (only write if content actually changed)
        if args.backup and action not in READ_ONLY_ACTIONS and pre_hash is not None:
            if os.path.exists(file_path):
                post_hash = _file_hash(file_path)
                if pre_hash != post_hash:
                    batch_tools.create_backup(file_path)

        # --verify: read back affected range
        if args.verify and action not in READ_ONLY_ACTIONS:
            doc_info = core.get_info(file_path, json_mode=True)
            if isinstance(doc_info, dict):
                n_paras = doc_info.get("info", {}).get("paragraph_count", "?")
                verify_note = f"\n[VERIFY] Document now has {n_paras} paragraphs."
                if isinstance(result, str):
                    result += verify_note
                elif isinstance(result, dict):
                    result["verify_para_count"] = n_paras

        _emit(result, use_json, command_name=action, file_path=file_path)

    except Exception as e:
        import traceback
        err_msg = f"ERROR: Unexpected exception in '{action}': {e}"
        if use_json:
            print(json.dumps({"status": "ERROR", "command": action, "message": err_msg,
                               "traceback": traceback.format_exc()}, indent=2))
        else:
            print(err_msg)
    finally:
        if args.cleanup:
            cleanup()


if __name__ == "__main__":
    main()
