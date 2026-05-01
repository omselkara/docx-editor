"""Batch processing, diff/undo, document comparison — all in-process (no subprocess)."""
import os
import json
import time
import shutil
import datetime
from typing import List, Optional, Dict, Any, Union
from docx_engine.core import load_document
from docx_engine.constants import LCS_PARAGRAPH_CAP, BACKUP_SUFFIXES, MAX_BACKUP_COUNT
from docx_engine import errors

# ===================== BACKUP / UNDO =====================

def create_backup(doc_path: str) -> Optional[str]:
    """Create a rolling backup (up to 3 snapshots)."""
    if not os.path.exists(doc_path):
        return None
    backup_path = doc_path + ".bak"
    for i in range(3, 0, -1):
        old = f"{doc_path}.bak{i}"
        newer = f"{doc_path}.bak{i+1}" if i < 3 else None
        if os.path.exists(old):
            if newer:
                shutil.copy2(old, newer)
            else:
                os.remove(old)
    if os.path.exists(backup_path):
        shutil.copy2(backup_path, f"{doc_path}.bak1")
    shutil.copy2(doc_path, backup_path)
    return backup_path


def undo(doc_path: str) -> str:
    """Restore the most recent backup."""
    backup_path = doc_path + ".bak"
    if not os.path.exists(backup_path):
        return errors.err("batch_tools", "undo", "Backup file not found. Cannot undo.")
    shutil.copy2(backup_path, doc_path)
    for i in range(1, 4):
        old = f"{doc_path}.bak{i}"
        if os.path.exists(old):
            shutil.copy2(old, backup_path)
            os.remove(old)
            break
    return errors.ok(f"Document restored: {doc_path}")


def list_backups(doc_path: str) -> str:
    """List all available backups for a document."""
    backups = []
    for suffix in BACKUP_SUFFIXES:
        path = doc_path + suffix
        if os.path.exists(path):
            stat = os.stat(path)
            mod_time = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            backups.append({"path": path, "size_bytes": stat.st_size, "modified": mod_time})
    if not backups:
        return errors.err("batch_tools", "list_backups", "No backup files found.")
    lines = ["=== BACKUPS ==="]
    for b in backups:
        lines.append(f"  {b['path']} ({b['size_bytes']} bytes, {b['modified']})")
    return "\n".join(lines)


# ===================== DIFF (Document Comparison) =====================

def diff_documents(doc_path1: str, doc_path2: str, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    """Compare two DOCX documents paragraph by paragraph using a proper LCS diff."""
    doc1, err1 = load_document(doc_path1)
    if err1:
        return err1
    doc2, err2 = load_document(doc_path2)
    if err2:
        return err2

    paras1 = [(p.text, p.style.name) for p in doc1.paragraphs]
    paras2 = [(p.text, p.style.name) for p in doc2.paragraphs]

    changes = _lcs_diff(paras1, paras2)

    if json_mode:
        return {
            "status": "SUCCESS",
            "file1": doc_path1,
            "file2": doc_path2,
            "para_count_1": len(paras1),
            "para_count_2": len(paras2),
            "diff_count": len(changes),
            "changes": changes,
        }

    if not changes:
        return errors.ok("Documents are identical (no paragraph-level differences found).")

    lines = ["=== DOCUMENT COMPARISON ==="]
    lines.append(f"  File 1: {doc_path1} ({len(paras1)} paragraphs)")
    lines.append(f"  File 2: {doc_path2} ({len(paras2)} paragraphs)")
    lines.append(f"  Total differences: {len(changes)}")
    lines.append(f"\n  [+] Added | [-] Deleted | [~] Style changed\n")
    for c in changes:
        lines.append(f"  [{c['type']}] {c.get('info', '')}")
    return "\n".join(lines)


def _lcs_diff(paras1: List[tuple], paras2: List[tuple]) -> List[Dict[str, Any]]:
    """Produce a diff using standard LCS on paragraph text."""
    texts1 = [p[0] for p in paras1]
    texts2 = [p[0] for p in paras2]
    n, m = len(texts1), len(texts2)

    # Build LCS table (space-optimized for large docs)
    cap = LCS_PARAGRAPH_CAP
    t1 = texts1[:cap]
    t2 = texts2[:cap]
    n2, m2 = len(t1), len(t2)

    dp = [[0] * (m2 + 1) for _ in range(n2 + 1)]
    for i in range(1, n2 + 1):
        for j in range(1, m2 + 1):
            if t1[i-1] == t2[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
            else:
                dp[i][j] = max(dp[i-1][j], dp[i][j-1])

    # Backtrack
    changes = []
    i, j = n2, m2
    while i > 0 or j > 0:
        if i > 0 and j > 0 and t1[i-1] == t2[j-1]:
            # Same text — check style change
            if paras1[i-1][1] != paras2[j-1][1]:
                changes.append({"type": "~", "para_1": i-1, "para_2": j-1,
                                 "info": f"P{i-1}: style {paras1[i-1][1]} → {paras2[j-1][1]}"})
            i -= 1
            j -= 1
        elif j > 0 and (i == 0 or dp[i][j-1] >= dp[i-1][j]):
            changes.append({"type": "+", "para_2": j-1,
                             "info": f"P{j-1} added: {t2[j-1][:80]}"})
            j -= 1
        else:
            changes.append({"type": "-", "para_1": i-1,
                             "info": f"P{i-1} deleted: {t1[i-1][:80]}"})
            i -= 1

    changes.reverse()
    return changes


# ===================== IN-PROCESS BATCH EXECUTOR =====================

def _build_dispatch() -> Dict[str, Any]:
    """Lazily build dispatch table to avoid circular imports at module load."""
    from docx_engine import editing, formatting, tables, annotations, advanced, smart_features, extended

    return {
        # Editing
        "insert_paragraph": editing.insert_paragraph,
        "insert_heading": editing.insert_heading,
        "delete_paragraphs": editing.delete_paragraphs,
        "replace_text": editing.replace_text,
        "append_text": editing.append_text,
        # Formatting
        "format_text": formatting.format_text,
        "format_paragraph": formatting.format_paragraph,
        "apply_style": formatting.apply_style,
        # Tables
        "create_table": tables.create_table,
        "modify_cell": tables.modify_cell,
        "add_row": tables.add_row,
        "delete_row": tables.delete_row,
        "add_column": tables.add_column,
        "delete_column": tables.delete_column,
        "merge_cells": tables.merge_cells,
        "format_table_cell": tables.format_table_cell,
        # Annotations
        "add_comment": annotations.add_comment,
        "delete_comment": annotations.delete_comment,
        "accept_all_changes": annotations.accept_all_changes,
        "reject_all_changes": annotations.reject_all_changes,
        "add_bookmark": annotations.add_bookmark,
        "add_footnote": annotations.add_footnote,
        # Advanced
        "insert_image": advanced.insert_image,
        "set_header": advanced.set_header,
        "set_footer": advanced.set_footer,
        "set_margins": advanced.set_margins,
        "set_orientation": advanced.set_orientation,
        "set_page_size": advanced.set_page_size,
        "insert_page_break": advanced.insert_page_break,
        "insert_section_break": advanced.insert_section_break,
        "insert_list": advanced.insert_list,
        "insert_hyperlink": advanced.insert_hyperlink,
        "insert_toc": advanced.insert_toc,
        # Smart features
        "clone_format": smart_features.clone_format,
        "from_template": smart_features.from_template,
        # Extended
        "add_watermark": extended.add_watermark,
        "remove_watermark": extended.remove_watermark,
        "set_line_numbering": extended.set_line_numbering,
    }


def batch_execute(doc_path: str, commands_json_path: str, dry_run: bool = False, json_mode: bool = False) -> Union[str, Dict[str, Any]]:
    """Execute multiple commands from a JSON file — in-process, atomic, with rollback.

    JSON format:
    {
        "on_error": "stop",   // or "continue" — default: stop
        "commands": [
            {"action": "replace_text", "args": {"find": "Old", "replace": "New"}},
            {"action": "insert_paragraph", "args": {"text": "Hello"}, "on_error": "continue"}
        ]
    }
    """
    if not os.path.exists(commands_json_path):
        return errors.err("batch_tools", "batch_execute", f"Command file not found: {commands_json_path}")

    with open(commands_json_path, 'r', encoding='utf-8') as f:
        payload = json.load(f)

    # Support both bare list and envelope format
    if isinstance(payload, list):
        commands = payload
        global_on_error = "stop"
    elif isinstance(payload, dict):
        commands = payload.get("commands", [])
        global_on_error = payload.get("on_error", "stop")
    else:
        return errors.err("batch_tools", "batch_execute", "JSON must be a list of commands or an object with 'commands' key.")

    if not commands:
        return errors.warn("batch_tools", "batch_execute", "No commands to execute.")

    dispatch = _build_dispatch()

    # Auto-backup before batch (unless dry run)
    backup_path = None
    if not dry_run and os.path.exists(doc_path):
        backup_path = create_backup(doc_path)

    results = []
    succeeded = 0
    failed = 0
    aborted = False

    for i, cmd in enumerate(commands):
        action = cmd.get("action", "")
        args = cmd.get("args", {})
        cmd_on_error = cmd.get("on_error", global_on_error)
        elapsed_ms = 0

        if action not in dispatch:
            status = "ERROR"
            msg = errors.err("batch_tools", "execute", f"Unknown action: '{action}'.")
            failed += 1
        else:
            fn = dispatch[action]
            # Always pass doc_path as first arg
            call_args = {k.replace("-", "_"): v for k, v in args.items()}
            call_args["doc_path"] = doc_path

            t0 = time.monotonic()
            try:
                output = fn(**call_args) if not dry_run else f"[DRY-RUN] Would call {action}({call_args})"
                elapsed_ms = int((time.monotonic() - t0) * 1000)
                status = "ERROR" if isinstance(output, str) and output.startswith("ERROR") else "SUCCESS"
                msg = str(output)[:300]
                if status == "SUCCESS":
                    succeeded += 1
                else:
                    failed += 1
            except TypeError as e:
                elapsed_ms = int((time.monotonic() - t0) * 1000)
                status = "ERROR"
                msg = errors.err("batch_tools", "execute", f"Argument error in '{action}': {e}")
                failed += 1
            except Exception as e:
                elapsed_ms = int((time.monotonic() - t0) * 1000)
                status = "ERROR"
                msg = errors.err("batch_tools", "execute", f"Exception in '{action}': {e}")
                failed += 1

        results.append({
            "step": i + 1,
            "action": action,
            "status": status,
            "message": msg,
            "ms": elapsed_ms,
        })

        if status == "ERROR" and cmd_on_error == "stop":
            # Rollback
            if backup_path and not dry_run:
                shutil.copy2(backup_path, doc_path)
                results.append({"step": "ROLLBACK", "action": "restore_backup",
                                 "status": "SUCCESS", "message": f"Rolled back to {backup_path}", "ms": 0})
            aborted = True
            break

    total_ms = sum(r.get("ms", 0) for r in results)
    summary = {
        "status": "SUCCESS" if failed == 0 else ("PARTIAL" if succeeded > 0 else "ERROR"),
        "total": len(commands),
        "succeeded": succeeded,
        "failed": failed,
        "aborted": aborted,
        "dry_run": dry_run,
        "total_ms": total_ms,
        "commands": results,
    }

    if json_mode:
        return summary

    status_str = summary["status"]
    lines = [f"=== BATCH {'(DRY RUN) ' if dry_run else ''}RESULTS — {status_str} ({total_ms}ms) ==="]
    lines.append(f"  {succeeded}/{len(commands)} succeeded | {failed} failed | aborted: {aborted}")
    for r in results:
        lines.append(f"  [{r['step']}] {r['status']}: {r['action']} ({r['ms']}ms) → {r['message'][:100]}")
    return "\n".join(lines)


# ===================== JSON OUTPUT WRAPPER (legacy compat) =====================

def to_json(output_text: str) -> str:
    """Legacy converter: wrap any text output in a simple JSON envelope."""
    return json.dumps({"status": "SUCCESS", "raw": output_text}, ensure_ascii=False, separators=(',', ':'))
