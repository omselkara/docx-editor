# DOCX Editor Skill for Gemini CLI

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/omselkara/docx-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/omselkara/docx-editor/actions)

**Ultimate MS Office DOCX manipulation for LLM Agents.**

A specialized [Gemini CLI Skill](https://github.com/google/gemini-cli) that empowers AI agents to read, write, format, and manipulate Microsoft Word (`.docx`) files with surgical precision. Designed with **token efficiency** and **performance** in mind: LXML XML fast-paths, JSON-minified outputs, and atomic batch operations.

## Features

- **Read & Analyze:** Outlines, dense text, tables, document maps, page-based reading
- **Surgical Editing:** Formatting-preserving multi-run text replacement, paragraph insertion/deletion
- **Formatting:** Styles, bold/italic/underline, colors, alignment, spacing
- **Tables:** Full CRUD — create, read, modify cells, add/delete rows and columns, merge
- **Annotations:** Comments, tracked changes (accept/reject), footnotes, bookmarks
- **Advanced:** Headers/footers, margins, page layout, TOC, hyperlinks, images
- **Smart Features:** Document summary map, template variable substitution, language detection
- **Batch Operations:** Atomic execution with auto-rollback on failure, document diffing, rolling backups
- **Agent-Optimized:** Structured JSON designed for LLM programmatic consumption

## Requirements

- **Python >= 3.9**
- `python-docx >= 1.1.0`
- `lxml >= 4.9.0`
- `Pillow >= 10.0`
- `pdf2image >= 1.16` *(optional — only for visual page rendering)*
- **LibreOffice** *(optional — only for DOCX→PDF→PNG rendering)*

## Installation

### As a Gemini CLI Skill

```bash
gemini skills install https://github.com/omselkara/docx-editor.git
```

Then reload skills in Gemini CLI:
```
/skills reload
```

### As a Python Library

```bash
git clone https://github.com/omselkara/docx-editor.git
cd docx-editor
pip install -e .
# For development (tests, linting):
pip install -e ".[dev]"
```

## Quick Start

```bash
# Read document structure
python scripts/docx_agent.py report.docx outline

# Replace text while preserving formatting
python scripts/docx_agent.py contract.docx replace_text --find DRAFT --replace FINAL

# Read a specific section
python scripts/docx_agent.py thesis.docx read_section --heading "Introduction"

# Create a table
python scripts/docx_agent.py invoice.docx create_table --headers "Item,Qty,Price"

# Batch operations with auto-rollback
python scripts/docx_agent.py doc.docx batch_execute --commands commands.json

# Get JSON output (for agent use)
python scripts/docx_agent.py report.docx full_text --json
```

## Usage (For Agents)

Example prompts with Gemini CLI:
- *"Read the outline of report.docx and tell me the main sections."*
- *"Find the word 'DRAFT' in contract.docx and replace it with 'FINAL'."*
- *"Create a new document called invoice.docx and add a table with 3 columns: Item, Qty, Price."*
- *"Batch-replace all dates in Q4-report.docx using commands.json."*

## Development

```bash
# Run tests
pytest tests/

# Run tests with coverage
pytest tests/ --cov=scripts/docx_engine

# Run performance benchmarks
pytest tests/test_performance.py -v
pytest tests/test_performance.py --benchmark-save=baseline
pytest tests/test_performance.py --benchmark-compare=baseline

# Lint
ruff check scripts/
```

## Architecture

Progressive Disclosure pattern — expensive context loaded only when needed:

```
SKILL.md            ← Agent trigger + routing
scripts/
  docx_agent.py     ← CLI entry point (70+ subcommands)
  docx_engine/
    constants.py    ← Centralized magic numbers
    errors.py       ← Standardized error message builders
    core.py         ← Document lifecycle (create, load, save, info)
    reading.py      ← Outline, full_text (LXML fast-path), search
    editing.py      ← Surgical text replace, insert/delete paragraphs
    formatting.py   ← Styles, bold/italic/colors, paragraph format
    tables.py       ← Full table CRUD
    annotations.py  ← Comments, tracked changes, footnotes, bookmarks
    advanced.py     ← Images, headers/footers, page layout, TOC
    extended.py     ← SmartArt, charts, protection, watermarks
    smart_features.py ← Document map, templates, language detection
    batch_tools.py  ← Atomic batch execution, LCS diff, backup/undo
    rendering.py    ← DOCX→PDF→PNG pipeline
references/         ← Detailed command reference (loaded on-demand)
```

## Troubleshooting

**LibreOffice not found (visual rendering)**
Install LibreOffice and ensure `soffice` is in your PATH, or pass the full path via the `--libreoffice-path` flag. The skill falls back to text-based layout rendering automatically.

**`pdf2image` errors**
Install Poppler: `winget install poppler` (Windows) or `brew install poppler` (macOS).

**Unicode issues on Windows**
The agent auto-configures the console for UTF-8. If you still see encoding errors, set `PYTHONIOENCODING=utf-8` in your environment.

## License

[MIT License](LICENSE)
