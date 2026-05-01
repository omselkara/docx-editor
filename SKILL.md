---
name: docx-editor
description: Ultimate MS Office DOCX manipulation — full reading, surgical editing, formatting, tables, annotations, layout, rendering, and batch processing. Optimized for LLM agents with compact output, structured JSON, and atomic batch operations.
---

# Ultimate DOCX Editor Skill

The `scripts/docx_agent.py` tool provides **full Microsoft Word (DOCX) capabilities** via the command line.
Use `python scripts/docx_agent.py <file> <command> [options]`.

---

## Global Flags

| Flag | What it does |
|---|---|
| `--json` | Emit structured JSON (`{"status","command","data",...}`). Use for all programmatic consumption. |
| `--backup` | Lazy backup: only writes `.bak` if the document actually changed. |
| `--cleanup` | Delete all temp files after execution. |
| `--verify` | After any write op, reads back paragraph count so you can confirm the change. |

**Response prefixes (plain-text mode):**
- `SUCCESS:` — operation completed
- `ERROR:` — operation failed; details follow
- `WARNING:` — operation completed with caveats

---

## ⚡ Token Cost Guide

> Use this to pick the most token-efficient command for your task.

| Command | Typical output tokens | Notes |
|---|---|---|
| `info` | ~80 | Always safe to run first |
| `outline` | ~50–400 | Proportional to heading count |
| `summary_map` | ~300–800 | **Best first read for any unknown document** |
| `search --compact` | ~20–200 | Most efficient targeted read |
| `read_section` | ~100–2 000 | Section-scoped read |
| `read_range` | ~50–500 | Use instead of full_text on large docs |
| `full_text` | ~500–50 000 | ⚠️ Avoid on long documents; use `--max-chars` |
| `full_text --compact` | ~300–15 000 | Skips empty lines, shorter format codes |
| `list_tables` | ~30–200 | — |
| `read_table` | ~50–1 000 | Markdown + coordinates |

---

## References

- **Agent Workflow Recipes**: See [references/workflows.md](references/workflows.md) for sequential editing steps, large document strategies, and batch building patterns.
- **Command Reference**: See [references/commands.md](references/commands.md) for the complete list of 50+ specific CLI commands and their arguments.

---

## Error Decision Tree

```
ERROR: File '...' not found
  → Use absolute paths. Check the file exists first.

ERROR: Invalid index / Invalid range
  → Run `outline` or `full_text --compact` to see valid [P_index] values.
  → Valid range: 0 to (paragraph_count - 1).

ERROR: Heading '...' not found
  → Run `outline` to see exact heading text. Matching is case-insensitive substring.

WARNING: No text found to format
  → The --match pattern found no runs. Try search first to confirm text exists.

ERROR: Style '...' not found
  → Run `list_styles --type paragraph` to see available style names.

ERROR: No backup files found (undo)
  → Always pass --backup before write operations. No backup = no undo.

ERROR: Unknown action
  → Run `python scripts/docx_agent.py --help` for the full command list.
```
