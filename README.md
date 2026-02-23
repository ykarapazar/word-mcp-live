<div align="center">

# word-mcp-live

**The MCP server that gives AI full control of Microsoft Word**

`105 tools` &middot; `Dual-mode architecture` &middot; `Live COM editing` &middot; `Cross-platform`

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools: 105](https://img.shields.io/badge/tools-105-orange)]()
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

## What is word-mcp-live?

An [MCP](https://modelcontextprotocol.io/) server that lets AI assistants create, read, and edit Microsoft Word documents. 78 cross-platform tools work on any OS using python-docx, while 27 Windows-exclusive tools use COM automation to edit documents **while they're open in Word** — with real-time tracked changes, per-operation undo, and zero file-locking issues.

|  | Cross-Platform Mode | Windows Live Mode |
|---|---|---|
| **Engine** | python-docx | COM automation (pywin32) |
| **Tools** | 78 | 27 |
| **Platform** | Windows, macOS, Linux | Windows only |
| **File state** | Must be closed | Must be open in Word |
| **Undo** | N/A (file-level saves) | Per-operation Ctrl+Z |
| **Tracked changes** | OOXML manipulation | Native Word revisions |

## Key Features

- **Live editing** — insert, delete, format, find & replace in documents open in Word; no save-close-reopen cycle
- **Per-operation undo** — every tool call is a single Ctrl+Z entry in Word's undo stack via `UndoRecord`
- **Tracked changes** — both OOXML-based (cross-platform) and native COM revisions (Windows)
- **Comments** — add and read comments anchored to specific text ranges
- **Layout diagnostics** — detect keep_with_next chains, heading style misuse, PageBreakBefore problems
- **Tables** — create, format, merge cells, auto-fit, alternating rows, cell shading
- **Footnotes & endnotes** — add, delete, validate, customize numbering styles
- **Multiple transports** — stdio (default), SSE, and streamable-http for remote deployment
- **PyPI packaging** — install with `pip install word-mcp-live` or run with `uvx`

## Architecture

```
┌─────────────────────────────────────────────────┐
│                 FastMCP Server                   │
│              (stdio / SSE / HTTP)                │
├────────────────────┬────────────────────────────┤
│  Cross-Platform    │     Windows Live            │
│  (python-docx)     │     (COM / pywin32)         │
│                    │                             │
│  78 tools          │     27 tools                │
│  File must be      │     File must be            │
│  CLOSED            │     OPEN in Word            │
│                    │                             │
│  ✓ Any OS          │     ✓ Per-op undo           │
│  ✓ File-based I/O  │     ✓ Native revisions      │
│  ✓ OOXML direct    │     ✓ Real-time updates      │
└────────────────────┴────────────────────────────┘
```

## Quick Start

### Install from PyPI

```bash
# Run directly (no install needed)
uvx --from word-mcp-live word_mcp_server

# Or install globally
pip install word-mcp-live
word_mcp_server
```

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["--from", "word-mcp-live", "word_mcp_server"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

### Claude Code

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["--from", "word-mcp-live", "word_mcp_server"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

### From Source

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -e .
python word_mcp_server.py
```

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_AUTHOR` | `"AI Assistant"` | Author name for tracked changes and comments |
| `MCP_AUTHOR_INITIALS` | `"AI"` | Author initials for comments |
| `MCP_TRANSPORT` | `stdio` | Transport type: `stdio`, `sse`, or `streamable-http` |
| `MCP_HOST` | `0.0.0.0` | Host to bind (for SSE/HTTP transports) |
| `MCP_PORT` | `8000` | Port to bind (for SSE/HTTP transports) |
| `FASTMCP_LOG_LEVEL` | `INFO` | Log level for FastMCP |

For remote deployment (e.g., Render), see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Tool Reference

### Summary

| Category | Cross-Platform | Windows Live | Total |
|----------|:-:|:-:|:-:|
| Document Management | 7 | — | 7 |
| Content | 14 | — | 14 |
| Formatting | 17 | 3 | 20 |
| Comments | 4 | 2 | 6 |
| Tracked Changes | 6 | 3 | 9 |
| Hyperlinks | 1 | — | 1 |
| Layout | 7 | 7 | 14 |
| Footnotes | 10 | — | 10 |
| Protection | 5 | — | 5 |
| Extraction | 4 | — | 4 |
| Reading | — | 6 | 6 |
| Editing | — | 5 | 5 |
| Undo | — | 2 | 2 |
| Screen Capture | — | 1 | 1 |
| Diagnostics | — | 1 | 1 |
| **Total** | **75** | **30** | **105** |

### Cross-Platform Tools

These work on Windows, macOS, and Linux using python-docx. The document file must be **closed** (not open in Word).

<details>
<summary><b>Document Management (7)</b></summary>

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
<summary><b>Content (14)</b></summary>

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
<summary><b>Formatting (17)</b></summary>

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
<summary><b>Comments (4)</b></summary>

| Tool | Description |
|------|-------------|
| `get_all_comments` | Extract all comments |
| `get_comments_by_author` | Filter comments by author |
| `get_comments_for_paragraph` | Get comments for a specific paragraph |
| `add_comment` | Add a comment anchored to text |

</details>

<details>
<summary><b>Tracked Changes (6)</b></summary>

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
<summary><b>Hyperlinks (1)</b></summary>

| Tool | Description |
|------|-------------|
| `manage_hyperlinks` | Add, list, remove, and update hyperlinks |

</details>

<details>
<summary><b>Layout (7)</b></summary>

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
<summary><b>Footnotes (10)</b></summary>

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
<summary><b>Protection (5)</b></summary>

| Tool | Description |
|------|-------------|
| `protect_document` | Add password protection |
| `unprotect_document` | Remove protection |
| `add_restricted_editing` | Restrict editing to specific sections |
| `add_digital_signature` | Add a digital signature |
| `verify_document` | Verify protection and signatures |

</details>

<details>
<summary><b>Extraction (4)</b></summary>

| Tool | Description |
|------|-------------|
| `get_paragraph_text_from_document` | Get text from a specific paragraph |
| `find_text_in_document` | Find text occurrences |
| `get_highlighted_text` | Extract highlighted/colored text |
| `convert_to_pdf` | Convert to PDF |

</details>

### Windows Live Tools

These require Windows with Microsoft Word installed. They operate on documents **currently open in Word** via COM automation. All destructive tools are wrapped with `UndoRecord` — each tool call appears as a single Ctrl+Z entry in Word's undo stack.

<details open>
<summary><b>Editing (8)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_insert_text` | Insert text at a position (with optional tracked changes) |
| `word_live_delete_text` | Delete a character range |
| `word_live_replace_text` | Find & replace via COM — works across tracked change boundaries; supports wildcards |
| `word_live_format_text` | Format text (bold, italic, font, highlight, paragraph alignment, page break before) |
| `word_live_add_table` | Insert a table |
| `word_live_format_table` | Format an existing table |
| `word_live_apply_list` | Apply bullet, numbered, or multilevel list formatting |
| `word_live_setup_heading_numbering` | Auto-numbered headings (1. / 1.1) with configurable style |

</details>

<details open>
<summary><b>Reading (7)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_text` | Get all text paragraph by paragraph |
| `word_live_get_page_text` | Get text from specific page(s) with char offsets for chaining |
| `word_live_get_paragraph_format` | Inspect paragraph formatting (font, spacing, alignment, list info, per-run detail) |
| `word_live_get_info` | Get document metadata (pages, words, sections) |
| `word_live_find_text` | Find text with context; supports wildcards (`^m`, `^t`, `^p`) |
| `word_live_get_undo_history` | List undo stack entries |
| `word_live_diagnose_layout` | Scan for layout problems (keep_with_next chains, style misuse, break issues) |

</details>

<details open>
<summary><b>Comments & Revisions (5)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_comments` | Get all comments |
| `word_live_add_comment` | Add a comment anchored to text |
| `word_live_list_revisions` | List tracked changes |
| `word_live_accept_revisions` | Accept tracked changes (all or by author/type) |
| `word_live_reject_revisions` | Reject tracked changes (all or by author/type) |

</details>

<details open>
<summary><b>Layout (7)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_set_page_layout` | Set orientation, size, and margins |
| `word_live_add_header_footer` | Add header/footer text |
| `word_live_add_page_numbers` | Add page numbers |
| `word_live_add_section_break` | Add section break |
| `word_live_set_paragraph_spacing` | Set paragraph spacing (line_spacing in **points**: 1.15× = 13.8pt) |
| `word_live_add_bookmark` | Add a named bookmark |
| `word_live_add_watermark` | Add a text watermark |

</details>

<details open>
<summary><b>Undo & Screen Capture (2)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_undo` | Undo last N operations (each tool call = one undo entry) |
| `word_screen_capture` | Screenshot of the Word window |

</details>

## Compared to the Original

This project builds on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (54 tools) with the following additions:

- **27 Windows Live tools** — COM automation for editing documents open in Word
- **21 new cross-platform tools** — tracked changes, comments, hyperlinks, layout, footnotes, protection
- **Per-operation undo** — `UndoRecord` wrapping on all destructive live tools
- **Layout diagnostics** — `word_live_diagnose_layout` and `word_live_get_paragraph_format`
- **Multiple transports** — stdio, SSE, and streamable-http (for remote deployment)
- **Configurable author** — `MCP_AUTHOR` environment variable for tracked changes and comments
- **PyPI packaging** — `pip install word-mcp-live` / `uvx --from word-mcp-live word_mcp_server`

## Requirements

- **Python 3.11+**
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (installed automatically via pip)
- **Windows Live tools only:** Windows 10/11 + Microsoft Word + `pywin32`

> **Note:** The 78 cross-platform tools work without Word installed — only python-docx is needed.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, code style, and how to add new tools.

Found a bug? [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md).
Have an idea? [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md).

## Acknowledgments

Built on top of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) by GongRzhe (MIT License).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/) &middot; [FastMCP](https://github.com/modelcontextprotocol/python-sdk) &middot; [pywin32](https://github.com/mhammond/pywin32)

## License

MIT License — see [LICENSE](LICENSE) for details.
