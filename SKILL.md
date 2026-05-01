---
name: docx-editor
description: Full MS Office DOCX capabilities (read, edit, format, tables, rendering). Use when user needs to manipulate .docx files with surgical precision or batch operations.
---

# DOCX Editor Skill

Access via: `python scripts/docx_agent.py <file> <command> [options]`

## 🛠 Global Flags
- `--json`: Mandatory for programmatic use. Minified JSON output.
- `--backup`: Automated `.bak` creation before changes.
- `--cleanup`: Auto-delete temp files.
- `--verify`: Immediate read-back validation after write operations.

## 📚 Deep Knowledge base
- **[Full Command List](references/commands.md)**: 50+ CLI commands with examples.
- **[Optimization Workflows](references/workflows.md)**: Efficient editing patterns for large docs and tables.

## ⚡ Precision Rules
1. **Never read `full_text` first.** Use `info` -> `outline` -> `summary_map` to map the doc structure.
2. **Targeted Reads**: Use `search --compact`, `read_range`, or `read_section` to save tokens.
3. **Verify Every Edit**: Always use `--verify` or follow up with a `read_range` to confirm success.
4. **Batch Operations**: Use `batch` with a JSON plan for multiple sequential edits to minimize disk I/O.
