### Safe Edit
```
1. python scripts/docx_agent.py doc.docx info
2. python scripts/docx_agent.py doc.docx outline
3. python scripts/docx_agent.py doc.docx search --query "text"
4. python scripts/docx_agent.py doc.docx --backup replace_text --find "old" --replace "new"
5. python scripts/docx_agent.py doc.docx read_range --start N --end N
```

### Large Docs
```
1. summary_map
2. read_section
3. search --compact --query "keyword"
4. read_range --start X --end Y
```

### Tables
```
1. list_tables
2. read_table --index 0
3. modify_cell --index 0 --row 1 --col 2 --text "val"
4. read_table --index 0
```

### Batch
```json
{
  "on_error": "stop",
  "commands": [
    {"action": "insert_heading", "args": {"text": "Title", "level": 1}},
    {"action": "insert_paragraph", "args": {"text": "Text...", "after-heading": "Title"}},
    {"action": "create_table", "args": {"headers": "A,B", "rows-data": "1,2"}}
  ]
}
```
```
python scripts/docx_agent.py doc.docx batch --commands plan.json --dry-run
python scripts/docx_agent.py doc.docx batch --commands plan.json
```