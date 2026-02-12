# word-mcp-live

MCP server for Microsoft Word with live COM automation on Windows.

![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)
![Tools: 99](https://img.shields.io/badge/tools-99-orange)

78 cross-platform tools (python-docx) work everywhere. 21 Windows-only live tools use COM automation to edit documents **while they're open in Word** — no file locking issues, real-time tracked changes, comments, and layout control.

## What's New vs the Original

This project builds on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (54 tools) and adds 45 new tools:

| Category | New Tools | What They Do |
|----------|-----------|--------------|
| Live editing (COM) | `word_live_insert_text`, `word_live_delete_text`, `word_live_format_text`, `word_live_add_table` | Edit documents open in Word — no lock conflicts |
| Live reading (COM) | `word_live_get_text`, `word_live_get_info`, `word_live_find_text` | Read from open documents |
| Tracked changes | `track_replace`, `track_insert`, `track_delete`, `list_tracked_changes`, `accept_tracked_changes`, `reject_tracked_changes` | Full revision tracking via OOXML manipulation |
| Live revisions (COM) | `word_live_list_revisions`, `word_live_accept_revisions`, `word_live_reject_revisions` | Manage tracked changes in open documents |
| Comments | `add_comment`, `word_live_add_comment`, `word_live_get_comments` | Write and read comments (both file-based and COM) |
| Hyperlinks | `manage_hyperlinks` | Add, list, remove, and update hyperlinks |
| Layout (COM) | `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark` | Full document layout control in open documents |
| Layout (file-based) | `set_page_layout`, `add_header_footer`, `add_page_numbers`, `add_section_break`, `set_paragraph_spacing`, `add_bookmark`, `add_watermark` | Same layout tools for file-based editing |
| Footnotes | 10 tools | Add, delete, validate, customize footnotes and endnotes |
| Protection | `protect_document`, `unprotect_document`, `add_restricted_editing`, `add_digital_signature`, `verify_document` | Document protection and signatures |
| Screen capture | `word_screen_capture` | Screenshot of the Word window |

### Configurable author name

All tools that write author metadata (tracked changes, comments) read from the `MCP_AUTHOR` environment variable. Set it in your MCP config to use your own name:

```json
{
  "env": {
    "MCP_AUTHOR": "Your Name",
    "MCP_AUTHOR_INITIALS": "YN"
  }
}
```

## Tool List

### Cross-Platform Tools (78)

These work on Windows, macOS, and Linux using python-docx.

<details>
<summary>Document Management (7)</summary>

| Tool | Description |
|------|-------------|
| `create_document` | Create a new Word document with optional metadata |
| `copy_document` | Create a copy of a Word document |
| `get_document_info` | Get document properties and statistics |
| `get_document_text` | Extract all text from a document |
| `get_document_outline` | Get document heading structure |
| `list_available_documents` | List .docx files in a directory |
| `get_document_xml` | Get raw OOXML structure |

</details>

<details>
<summary>Content (14)</summary>

| Tool | Description |
|------|-------------|
| `add_paragraph` | Add a paragraph with optional formatting |
| `add_heading` | Add a heading (levels 1-9) with formatting |
| `add_table` | Add a table with custom data |
| `add_picture` | Add an image with proportional scaling |
| `add_page_break` | Insert a page break |
| `delete_paragraph` | Delete a paragraph by index |
| `search_and_replace` | Find and replace text |
| `add_table_of_contents` | Add a TOC based on heading styles |
| `insert_header_near_text` | Insert a heading before/after target text |
| `insert_line_or_paragraph_near_text` | Insert a paragraph before/after target text |
| `insert_numbered_list_near_text` | Insert a bulleted or numbered list |
| `replace_paragraph_block_below_header` | Replace content under a heading |
| `replace_block_between_manual_anchors` | Replace content between anchor texts |
| `merge_documents` | Merge multiple documents into one |

</details>

<details>
<summary>Formatting (19)</summary>

| Tool | Description |
|------|-------------|
| `format_text` | Format text (bold, italic, color, font, size) |
| `create_custom_style` | Create a custom document style |
| `format_table` | Format table borders and structure |
| `set_table_cell_shading` | Set cell background color |
| `apply_table_alternating_rows` | Alternating row colors |
| `highlight_table_header` | Highlight header row |
| `merge_table_cells` | Merge a rectangular cell area |
| `merge_table_cells_horizontal` | Merge cells in a row |
| `merge_table_cells_vertical` | Merge cells in a column |
| `set_table_cell_alignment` | Set cell text alignment |
| `set_table_alignment_all` | Set alignment for all cells |
| `set_table_column_width` | Set a column's width |
| `set_table_column_widths` | Set multiple column widths |
| `set_table_width` | Set overall table width |
| `auto_fit_table_columns` | Auto-fit columns to content |
| `format_table_cell_text` | Format text in a specific cell |
| `set_table_cell_padding` | Set cell padding |

</details>

<details>
<summary>Comments (4)</summary>

| Tool | Description |
|------|-------------|
| `get_all_comments` | Extract all comments |
| `get_comments_by_author` | Filter comments by author |
| `get_comments_for_paragraph` | Get comments for a specific paragraph |
| `add_comment` | Add a comment anchored to text |

</details>

<details>
<summary>Tracked Changes (6)</summary>

| Tool | Description |
|------|-------------|
| `track_replace` | Replace text as a tracked change |
| `track_insert` | Insert text as a tracked change |
| `track_delete` | Delete text as a tracked change |
| `list_tracked_changes` | List all tracked changes |
| `accept_tracked_changes` | Accept all tracked changes |
| `reject_tracked_changes` | Reject all tracked changes |

</details>

<details>
<summary>Hyperlinks (1)</summary>

| Tool | Description |
|------|-------------|
| `manage_hyperlinks` | Add, list, remove, and update hyperlinks |

</details>

<details>
<summary>Layout (9)</summary>

| Tool | Description |
|------|-------------|
| `set_page_layout` | Set orientation, size, and margins |
| `add_header_footer` | Add header/footer text |
| `add_page_numbers` | Add page numbers |
| `add_section_break` | Add section break (new page, continuous, etc.) |
| `set_paragraph_spacing` | Set paragraph spacing |
| `add_bookmark` | Add a named bookmark |
| `add_watermark` | Add a diagonal text watermark |

</details>

<details>
<summary>Footnotes (10)</summary>

| Tool | Description |
|------|-------------|
| `add_footnote_to_document` | Add a footnote to a paragraph |
| `add_footnote_after_text` | Add footnote after specific text |
| `add_footnote_before_text` | Add footnote before specific text |
| `add_footnote_enhanced` | Enhanced footnote with superscript |
| `add_footnote_robust` | Robust footnote with validation |
| `add_endnote_to_document` | Add an endnote |
| `customize_footnote_style` | Customize footnote numbering |
| `delete_footnote_from_document` | Delete a footnote |
| `delete_footnote_robust` | Delete with cleanup |
| `validate_document_footnotes` | Validate all footnotes |

</details>

<details>
<summary>Protection (5)</summary>

| Tool | Description |
|------|-------------|
| `protect_document` | Add password protection |
| `unprotect_document` | Remove protection |
| `add_restricted_editing` | Restrict editing to specific sections |
| `add_digital_signature` | Add a digital signature |
| `verify_document` | Verify protection and signatures |

</details>

<details>
<summary>Extraction (4)</summary>

| Tool | Description |
|------|-------------|
| `get_paragraph_text_from_document` | Get text from a specific paragraph |
| `find_text_in_document` | Find text occurrences |
| `get_highlighted_text` | Extract highlighted/colored text |
| `convert_to_pdf` | Convert to PDF |

</details>

### Windows Live Tools (21)

These require Windows with Microsoft Word installed. They operate on documents **currently open in Word** via COM automation — no file locking issues.

| Tool | Description |
|------|-------------|
| **Screen Capture** | |
| `word_screen_capture` | Screenshot of the Word window |
| **Editing** | |
| `word_live_insert_text` | Insert text (with optional tracked changes) |
| `word_live_delete_text` | Delete text |
| `word_live_format_text` | Format text (bold, italic, font, highlight, etc.) |
| `word_live_add_table` | Add a table |
| **Reading** | |
| `word_live_get_text` | Get all text paragraph by paragraph |
| `word_live_get_info` | Get document metadata (pages, words, sections) |
| `word_live_find_text` | Find text with context |
| **Comments & Revisions** | |
| `word_live_get_comments` | Get all comments |
| `word_live_add_comment` | Add a comment |
| `word_live_list_revisions` | List tracked changes |
| `word_live_accept_revisions` | Accept tracked changes |
| `word_live_reject_revisions` | Reject tracked changes |
| **Layout** | |
| `word_live_set_page_layout` | Set orientation, size, margins |
| `word_live_add_header_footer` | Add header/footer |
| `word_live_add_page_numbers` | Add page numbers |
| `word_live_add_section_break` | Add section break |
| `word_live_set_paragraph_spacing` | Set paragraph spacing |
| `word_live_add_bookmark` | Add a bookmark |
| `word_live_add_watermark` | Add a watermark |

## Quick Start

### Claude Code (`.mcp.json`)

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["path/to/word_mcp_server.py"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["--from", "word-mcp-live", "word_mcp_server"]
    }
  }
}
```

### From source

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -r requirements.txt
python word_mcp_server.py
```

## Requirements

- Python 3.11+
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (see `requirements.txt`)
- **Windows Live tools only:** Windows 10/11 + Microsoft Word + `pywin32`

## Acknowledgments

Built on top of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (MIT License).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/), [FastMCP](https://github.com/modelcontextprotocol/python-sdk), [pywin32](https://github.com/mhammond/pywin32).

## License

MIT License — see [LICENSE](LICENSE) for details.
