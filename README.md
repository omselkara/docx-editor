# DOCX Editor Skill for Gemini CLI

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Ultimate MS Office DOCX manipulation for LLM Agents.**

This repository contains a specialized [Gemini CLI Skill](https://github.com/google/gemini-cli) that empowers AI agents to read, write, format, and manipulate Microsoft Word (`.docx`) files with surgical precision. 

It is designed with **token efficiency** and **performance** in mind, featuring custom XML fast-paths, JSON-minified outputs, and batch operation support to keep context windows lean and execution times low.

## Features

- **Read & Analyze:** Extracts outlines, dense text, tables, and document maps.
- **Surgical Editing:** Precise text replacement, paragraph insertion, and deletion.
- **Formatting:** Apply styles, bold, italic, colors, and alignment.
- **Tables:** Create, read, and modify tables cell-by-cell.
- **Advanced:** Handle comments, tracked changes, footnotes, headers/footers, and templates.
- **Agent-Optimized:** Outputs structured JSON designed for LLM programmatic consumption.

## Installation

You can install this skill directly into your Gemini CLI environment using the repository URL:

```bash
gemini skills install https://github.com/omselkara/docx-editor.git
```

After installation, reload your skills in an active Gemini CLI session:
```
/skills reload
```

## Usage (For Agents)

Once installed, the skill activates automatically when you ask the agent to interact with a Word document. 

Example prompts you can use with Gemini CLI:
- *"Read the outline of report.docx and tell me the main sections."*
- *"Find the word 'DRAFT' in contract.docx and replace it with 'FINAL'."*
- *"Create a new document called invoice.docx and add a table with 3 columns: Item, Qty, Price."*
- *"Summarize the changes made in the latest revision of thesis.docx."*

## Architecture

This skill follows the Progressive Disclosure architectural pattern:
- `SKILL.md`: The core trigger and high-level routing file.
- `scripts/`: Contains the highly-optimized Python engine (`docx_agent.py` and `docx_engine/`).
- `references/`: Contains detailed command references and workflow recipes, loaded into the LLM's context only when needed.

## License

This project is licensed under the [MIT License](LICENSE).
