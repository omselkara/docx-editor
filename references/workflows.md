# Agent Workflow Recipes

### Safe Edit Workflow (recommended for all writes)
```
1. python scripts/docx_agent.py doc.docx info               # understand the doc
2. python scripts/docx_agent.py doc.docx outline            # find target heading/index
3. python scripts/docx_agent.py doc.docx search --query "target text"  # confirm location
4. python scripts/docx_agent.py doc.docx --backup replace_text --find "old" --replace "new"
5. python scripts/docx_agent.py doc.docx read_range --start N --end N  # verify change
```

### Large Document Workflow (>200 paragraphs)
```
1. summary_map    → get page-aware structure overview
2. read_section   → drill into specific section
3. search --compact --query "keyword"  → locate exact paragraph
4. read_range --start X --end Y  → inspect before editing
```

### Table Editing Workflow
```
1. list_tables    → find table index [T0], [T1]…
2. read_table --index 0   → see full content + [row,col] coordinates
3. modify_cell --index 0 --row 1 --col 2 --text "new value"
4. read_table --index 0   → confirm
```

### Batch Document Building
```json
{
  "on_error": "stop",
  "commands": [
    {"action": "insert_heading", "args": {"text": "Introduction", "level": 1}},
    {"action": "insert_paragraph", "args": {"text": "This document covers...", "after-heading": "Introduction"}},
    {"action": "create_table", "args": {"headers": "Name,Role,Status", "rows-data": "Alice,Dev,Active;Bob,PM,Active"}}
  ]
}
```
```
python scripts/docx_agent.py doc.docx batch --commands plan.json --dry-run  # validate first
python scripts/docx_agent.py doc.docx batch --commands plan.json            # execute
```