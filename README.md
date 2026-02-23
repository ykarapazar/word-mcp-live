<div align="center">

# word-mcp-live

**The MCP server that gives AI full control of Microsoft Word**

`Live editing` &middot; `Tracked changes` &middot; `Per-action undo` &middot; `Cross-platform`

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

## What Can It Do?

word-mcp-live lets any AI assistant that supports [MCP](https://modelcontextprotocol.io/) work with Microsoft Word like a human would. Open a document, tell the AI what you need, and watch it happen — formatting, tracked changes, comments, and all.

Just tell the AI what you want in plain language:

```
"Draft a contract with tracked changes so my colleague can review"
"Format all headings as Cambria 13pt bold and add automatic numbering"
"Add a comment on paragraph 3 asking about the deadline"
"Find every mention of 'ABC Corp' and replace with 'XYZ Ltd' as a tracked change"
"Set the page to A4 landscape with 2cm margins"
"Insert a table of contents based on the document headings"
"Add page numbers in the footer and our company name in the header"
```

## Key Capabilities

- **Edit documents while they're open** — no save-close-reopen cycle
- **Track every change** — insertions, deletions, and replacements all appear as tracked changes your team can review
- **Undo anything** — every AI action is a single Ctrl+Z in Word
- **Add comments** — anchored to specific text, just like a human reviewer
- **Format professionally** — fonts, styles, headings, automatic numbering, tables, watermarks
- **Read any part** — extract text by page, search with context, get document structure
- **Full layout control** — margins, orientation, headers, footers, page numbers, section breaks
- **Works on any OS** — core features work on Windows, macOS, and Linux

## Two Modes

|  | Works everywhere | Windows with Word open |
|---|---|---|
| **What it does** | Create and edit saved .docx files | Edit documents live while you work in Word |
| **Platform** | Windows, macOS, Linux | Windows only |
| **Undo** | File-level saves | Per-action Ctrl+Z |
| **Best for** | Batch processing, document generation | Interactive editing, formatting, review |

Both modes work together. The AI picks the right one for the task.

## Getting Started

### Install

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -e .
```

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["/absolute/path/to/word-mcp-live/word_mcp_server.py"],
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
      "command": "python",
      "args": ["/absolute/path/to/word-mcp-live/word_mcp_server.py"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

> **About `MCP_AUTHOR`:** Set this to your name so that tracked changes and comments show your name instead of "AI Assistant". `MCP_AUTHOR_INITIALS` sets the initials shown on comments.

### Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_AUTHOR` | `"AI Assistant"` | Author name for tracked changes and comments |
| `MCP_AUTHOR_INITIALS` | `"AI"` | Author initials for comments |
| `MCP_TRANSPORT` | `stdio` | Transport type: `stdio`, `sse`, or `streamable-http` |
| `MCP_HOST` | `0.0.0.0` | Host to bind (for SSE/HTTP transports) |
| `MCP_PORT` | `8000` | Port to bind (for SSE/HTTP transports) |

For remote deployment, see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Tool Reference

See the [complete tool reference](TOOLS.md) for all capabilities, organized by category.

## Compared to the Original

This project builds on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) with significant additions:

- **Windows Live editing** — COM automation for editing documents open in Word
- **Per-operation undo** — every tool call is a single Ctrl+Z entry in Word's undo stack
- **Tracked changes** — both OOXML-based (cross-platform) and native Word revisions (Windows)
- **Comments** — add and read comments anchored to specific text
- **Layout diagnostics** — detect formatting problems like keep_with_next chains and heading misuse
- **New cross-platform tools** — hyperlinks, footnotes, endnotes, protection, digital signatures
- **Multiple transports** — stdio, SSE, and streamable-http for remote deployment
- **Configurable author** — `MCP_AUTHOR` for tracked changes and comments

## Requirements

- **Python 3.11+**
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (installed automatically by `pip install -e .`)
- **Windows Live tools only:** Windows 10/11 + Microsoft Word + `pywin32`

> The cross-platform tools work without Word installed — only python-docx is needed.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, code style, and how to add new tools.

Found a bug? [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md).
Have an idea? [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md).

## Acknowledgments

Built on top of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) by GongRzhe (MIT License).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/) &middot; [FastMCP](https://github.com/modelcontextprotocol/python-sdk) &middot; [pywin32](https://github.com/mhammond/pywin32)

## License

MIT License — see [LICENSE](LICENSE) for details.
