---
name: docx-editor
description: Full MS Office DOCX capabilities (read, edit, format, tables, rendering). Use for surgical manipulation or batch operations on .docx files.
---

# DOCX Editor Skill

Entry: `python scripts/docx_agent.py <file> <command> [options]`

## Global Flags
- `--json`: Mandatory for programmatic use. Minified output.
- `--backup`: Automatic .bak creation before changes.
- `--cleanup`: Auto-delete temp files.
- `--verify`: Immediate validation after write.

## Knowledge Base
- [Commands](references/commands.md): Full CLI command list.
- [Workflows](references/workflows.md): Optimization patterns for large docs/tables.

## Precision Rules
1. Never read full_text first. Use info -> outline -> summary_map.
2. Targeted Reads: Use search --compact, read_range, or read_section.
3. Verify Every Edit: Use --verify or read_range follow-up.
4. Batch Operations: Use batch for multiple edits to minimize I/O.
