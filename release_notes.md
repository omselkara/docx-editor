## What's in this release

### Core Features
- **75 CLI subcommands** covering read, edit, format, tables, annotations, advanced layout, batch operations
- **Surgical text replacement** — multi-run formatting preserved across paragraph boundaries
- **Atomic batch execution** — auto-rollback on failure, rolling backups (up to 3 snapshots)
- **LCS document diff** — paragraph-level comparison between two DOCX files
- **LXML fast-path** — 10x faster text extraction and search on large documents

### New in v2.0.0
- **Chart data extraction** — series names, numeric values, and category labels from embedded charts
- **Token efficiency flags** — `--max-depth` for outline, `--strip-format` for full_text/read_range
- **Robust document loading** — graceful handling of encrypted, corrupted, and invalid ZIP files
- **Type hints** across all modules, `errors.py` standardization, `constants.py` centralization
- **142-test suite** with CI on Python 3.9, 3.11, 3.12
- **Progressive disclosure** architecture — SKILL.md → references/ for token-efficient agent routing

### Designed For
LLM agents (Gemini CLI) that need to read, write, and manipulate Microsoft Word documents with surgical precision and minimal token overhead.

### Installation
```bash
gemini skills install https://github.com/omselkara/docx-editor.git
```
Or as a Python library:
```bash
git clone https://github.com/omselkara/docx-editor.git
cd docx-editor
pip install -e .
```
