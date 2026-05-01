# Command Reference

### 1. Document Lifecycle

| Command | Example |
|---|---|
| `create` | `python scripts/docx_agent.py new.docx create` |
| `info` | `python scripts/docx_agent.py doc.docx info --json` |
| `set_metadata` | `python scripts/docx_agent.py doc.docx set_metadata --title "Report" --author "Alice"` |

### 2. Reading & Analysis

| Command | Example |
|---|---|
| `outline` | `python scripts/docx_agent.py doc.docx outline` |
| `summary_map` | `python scripts/docx_agent.py doc.docx summary_map` |
| `full_text` | `python scripts/docx_agent.py doc.docx full_text --compact --max-chars 5000` |
| `full_text --formatting` | `python scripts/docx_agent.py doc.docx full_text --formatting --compact` |
| `read_section` | `python scripts/docx_agent.py doc.docx read_section --heading "Methods"` |
| `read_range` | `python scripts/docx_agent.py doc.docx read_range --start 10 --end 20` |
| `read_page` | `python scripts/docx_agent.py doc.docx read_page --page 3` |
| `search` | `python scripts/docx_agent.py doc.docx search --query "revenue" --context 2 --compact` |
| `word_count` | `python scripts/docx_agent.py doc.docx word_count` |
| `detect_language` | `python scripts/docx_agent.py doc.docx detect_language` |
| `full_statistics` | `python scripts/docx_agent.py doc.docx full_statistics` |

### 3. High-Precision Editing

| Command | Example |
|---|---|
| `replace_text` | `python scripts/docx_agent.py doc.docx --backup replace_text --find "Q1" --replace "Q2"` |
| `replace_text --regex` | `python scripts/docx_agent.py doc.docx replace_text --find "\d{4}" --replace "XXXX" --regex` |
| `insert_paragraph` | `python scripts/docx_agent.py doc.docx insert_paragraph --text "New content." --after-heading "Intro"` |
| `insert_heading` | `python scripts/docx_agent.py doc.docx insert_heading --text "New Section" --level 2 --index 5` |
| `delete_paragraphs` | `python scripts/docx_agent.py doc.docx --backup delete_paragraphs --indices "5,6,7"` |
| `append_text` | `python scripts/docx_agent.py doc.docx append_text --text "Conclusion." --style "Normal"` |

### 4. Formatting & Styles

| Command | Example |
|---|---|
| `format_text` | `python scripts/docx_agent.py doc.docx format_text --match "IMPORTANT" --bold true --font-color "#FF0000"` |
| `format_paragraph` | `python scripts/docx_agent.py doc.docx format_paragraph --para-indices "0,1" --alignment center` |
| `list_styles` | `python scripts/docx_agent.py doc.docx list_styles --type paragraph` |
| `apply_style` | `python scripts/docx_agent.py doc.docx apply_style --para-indices "0" --style "Heading 1"` |
| `clone_format` | `python scripts/docx_agent.py doc.docx clone_format --source-para 5 --target-paras "10,11"` |

### 5. Tables

| Command | Example |
|---|---|
| `list_tables` | `python scripts/docx_agent.py doc.docx list_tables` |
| `read_table` | `python scripts/docx_agent.py doc.docx read_table --index 0 --json` |
| `create_table` | `python scripts/docx_agent.py doc.docx create_table --headers "Name,Score" --rows-data "Alice,95;Bob,87"` |
| `modify_cell` | `python scripts/docx_agent.py doc.docx modify_cell --index 0 --row 1 --col 2 --text "Updated"` |
| `add_row` | `python scripts/docx_agent.py doc.docx add_row --index 0 --values "Carol,91"` |
| `delete_row` | `python scripts/docx_agent.py doc.docx delete_row --index 0 --row 3` |
| `add_column` | `python scripts/docx_agent.py doc.docx add_column --index 0 --header "Grade"` |
| `delete_column` | `python scripts/docx_agent.py doc.docx delete_column --index 0 --col 2` |
| `merge_cells` | `python scripts/docx_agent.py doc.docx merge_cells --index 0 --start-row 0 --start-col 0 --end-row 0 --end-col 2` |
| `format_table_cell` | `python scripts/docx_agent.py doc.docx format_table_cell --index 0 --row 0 --col 0 --bg-color "4472C4" --bold true` |

### 6. Annotations & Review

| Command | Example |
|---|---|
| `read_comments` | `python scripts/docx_agent.py doc.docx read_comments` |
| `add_comment` | `python scripts/docx_agent.py doc.docx add_comment --para 5 --text "Please review." --author "Claude"` |
| `delete_comment` | `python scripts/docx_agent.py doc.docx delete_comment --id 3` |
| `read_tracked_changes` | `python scripts/docx_agent.py doc.docx read_tracked_changes` |
| `accept_all_changes` | `python scripts/docx_agent.py doc.docx --backup accept_all_changes` |
| `reject_all_changes` | `python scripts/docx_agent.py doc.docx --backup reject_all_changes` |
| `read_footnotes` | `python scripts/docx_agent.py doc.docx read_footnotes` |
| `read_endnotes` | `python scripts/docx_agent.py doc.docx read_endnotes` |
| `add_footnote` | `python scripts/docx_agent.py doc.docx add_footnote --para 10 --text "See also chapter 3."` |
| `list_bookmarks` | `python scripts/docx_agent.py doc.docx list_bookmarks` |
| `add_bookmark` | `python scripts/docx_agent.py doc.docx add_bookmark --para 5 --name "key_finding"` |
| `read_textboxes` | `python scripts/docx_agent.py doc.docx read_textboxes` |

### 7. Layout & Page Setup

| Command | Example |
|---|---|
| `set_header` | `python scripts/docx_agent.py doc.docx set_header --text "CONFIDENTIAL"` |
| `set_footer` | `python scripts/docx_agent.py doc.docx set_footer --text "Page " --page-number` |
| `read_header` | `python scripts/docx_agent.py doc.docx read_header` |
| `read_footer` | `python scripts/docx_agent.py doc.docx read_footer` |
| `set_margins` | `python scripts/docx_agent.py doc.docx set_margins --top 2.5 --bottom 2.5 --left 3 --right 3` |
| `set_orientation` | `python scripts/docx_agent.py doc.docx set_orientation --orientation landscape` |
| `set_page_size` | `python scripts/docx_agent.py doc.docx set_page_size --preset a4` |
| `insert_page_break` | `python scripts/docx_agent.py doc.docx insert_page_break --after-para 15` |
| `insert_section_break` | `python scripts/docx_agent.py doc.docx insert_section_break --type continuous` |
| `describe_layout` | `python scripts/docx_agent.py doc.docx describe_layout` |

### 8. Content Elements

| Command | Example |
|---|---|
| `insert_image` | `python scripts/docx_agent.py doc.docx insert_image --image chart.png --width 4.5 --after-para 10` |
| `list_images` | `python scripts/docx_agent.py doc.docx list_images` |
| `insert_list` | `python scripts/docx_agent.py doc.docx insert_list --items "Item 1;Item 2;Item 3" --type bullet` |
| `insert_hyperlink` | `python scripts/docx_agent.py doc.docx insert_hyperlink --text "Visit site" --url "https://example.com"` |
| `insert_toc` | `python scripts/docx_agent.py doc.docx insert_toc --title "Contents"` |
| `add_watermark` | `python scripts/docx_agent.py doc.docx add_watermark --text "DRAFT"` |
| `remove_watermark` | `python scripts/docx_agent.py doc.docx remove_watermark` |
| `set_line_numbering` | (via batch or direct API call) |

### 9. Advanced / Complex Elements

| Command | Example |
|---|---|
| `read_smartart` | `python scripts/docx_agent.py doc.docx read_smartart` |
| `read_charts` | `python scripts/docx_agent.py doc.docx read_charts` |
| `list_embedded` | `python scripts/docx_agent.py doc.docx list_embedded` |
| `read_content_controls` | `python scripts/docx_agent.py doc.docx read_content_controls` |
| `read_protection` | `python scripts/docx_agent.py doc.docx read_protection` |

### 10. Templates

| Command | Example |
|---|---|
| `list_template_vars` | `python scripts/docx_agent.py doc.docx list_template_vars` |
| `from_template` | `python scripts/docx_agent.py doc.docx from_template --output out.docx --vars "name=Alice,date=2024-01-01"` |

### 11. Visual Rendering

| Command | Example |
|---|---|
| `render` | `python scripts/docx_agent.py doc.docx render --pages 1,2 --dpi 150` |

### 12. Batch & Comparison

| Command | Example |
|---|---|
| `batch` | `python scripts/docx_agent.py doc.docx batch --commands edits.json` |
| `batch --dry-run` | `python scripts/docx_agent.py doc.docx batch --commands edits.json --dry-run` |
| `diff` | `python scripts/docx_agent.py v1.docx diff --compare v2.docx --json` |

### 13. Backup & Undo

| Command | Example |
|---|---|
| `undo` | `python scripts/docx_agent.py doc.docx undo` |
| `list_backups` | `python scripts/docx_agent.py doc.docx list_backups` |