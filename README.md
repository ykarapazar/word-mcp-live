[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ykarapazar-word-mcp-live-badge.png)](https://mseep.ai/app/ykarapazar-word-mcp-live)

<div align="center">

[![Install in Cursor](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/en/install-mcp?name=word&config=eyJjb21tYW5kIjoidXZ4IiwiYXJncyI6WyJ3b3JkLW1jcC1saXZlIl19)

# word-mcp-live

**The only MCP server that edits Word documents while they're open**

`Live editing` &middot; `Tracked changes` &middot; `Per-action undo` &middot; `124 tools` &middot; `Cross-platform`

[![PyPI](https://img.shields.io/pypi/v/word-mcp-live?color=blue)](https://pypi.org/project/word-mcp-live/)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

word-mcp-live gives any AI assistant that supports [MCP](https://modelcontextprotocol.io/) full control of Microsoft Word. Open a document, tell the AI what you need, and watch it happen — formatting, tracked changes, comments, and all. Changes appear live in your open document.

<table>
<tr>
<td width="50%">

### Without word-mcp-live

- AI can discuss your document but can't touch it
- You copy-paste between AI and Word, losing formatting
- Track changes? You do those manually after the fact
- Every edit means save → close → process → reopen

</td>
<td width="50%">

### With word-mcp-live

- "Add a tracked change replacing ABC Corp with XYZ Ltd" — done
- Changes appear live in your open Word document
- Every AI edit is one Ctrl+Z away
- Real tracked changes with your name, not XML hacks

</td>
</tr>
</table>

### See it in action

https://github.com/user-attachments/assets/fbb09af4-1e25-4e49-94d0-45b363278810

## What Sets This Apart

- **Live editing** — Edit documents while they're open in Word. No save-close-reopen cycle.
- **Full undo** — Every AI action is a single Ctrl+Z. Made a mistake? Just undo it.
- **Native tracked changes** — Real Word revisions with your name, not XML hacks.
- **Threaded comments** — Add, reply, resolve, and delete comments like a human reviewer.
- **Layout diagnostics** — Detects formatting problems before they become print disasters.
- **Equations & cross-references** — Insert math formulas and auto-updating references.
- **124 tools** — The most comprehensive Word MCP server available.

## Quick Start

```bash
pip install word-mcp-live
```

Or install from source:

```bash
git clone https://github.com/ykarapazar/word-mcp-live.git
cd word-mcp-live
pip install -e .
```

## Client Installation

<details open>
<summary><b>Claude Desktop</b></summary>

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Claude Code</b></summary>

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Cursor</b></summary>

**One-click:** Click the install button at the top of this page.

**Manual:** Add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>VS Code / Copilot</b></summary>

**One-click:** [Install in VS Code](vscode:mcp/install?%7B%22name%22%3A%20%22word%22%2C%20%22command%22%3A%20%22uvx%22%2C%20%22args%22%3A%20%5B%22word-mcp-live%22%5D%7D)

**Manual:** Add to your VS Code `settings.json`:

```json
{
  "mcp": {
    "servers": {
      "word": {
        "command": "uvx",
        "args": ["word-mcp-live"],
        "env": {
          "MCP_AUTHOR": "Your Name",
          "MCP_AUTHOR_INITIALS": "YN"
        }
      }
    }
  }
}
```

</details>

<details>
<summary><b>Windsurf</b></summary>

Add to `~/.codeium/windsurf/mcp_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Docker</b></summary>

```json
{
  "mcpServers": {
    "word": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "ghcr.io/ykarapazar/word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

> Note: Docker mode supports cross-platform tools only. Live editing requires a native Windows install.

</details>

> **`MCP_AUTHOR`** sets your name on tracked changes and comments (default: `"Author"`). **`MCP_AUTHOR_INITIALS`** sets comment initials.

## Two Modes

|  | Works everywhere | Live editing (Word open) |
|---|---|---|
| **What it does** | Create and edit saved .docx files | Edit documents live while you work in Word |
| **Platform** | Windows, macOS, Linux | Windows (COM) and macOS (JXA) |
| **Undo** | File-level saves | Per-action Ctrl+Z (Windows); per-operation undo (macOS) |
| **Best for** | Batch processing, document generation | Interactive editing, formatting, review |

Both modes work together. The AI picks the right one for the task.

### macOS Live Editing (New in v1.5.0)

Live tools now work on macOS via JavaScript for Automation (JXA). Same tool names, same parameters — the server detects your platform and uses the right automation backend.

| Feature | Windows | macOS |
|---------|---------|-------|
| Text read/write/find/replace | COM | JXA |
| Formatting (bold, font, style) | COM | JXA |
| Track changes & revisions | COM | JXA |
| Comments (add, delete, list) | COM | JXA |
| Tables (read, write, add rows) | COM | JXA |
| Page layout, headers, bookmarks | COM | JXA |
| Equations, cross-references | COM | JXA |
| Threaded comment replies | COM | Not available |
| Comment resolve/unresolve | COM | Not available |
| Undo history inspection | COM | Not available |
| Watermarks | COM | Not available |

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_AUTHOR` | `"Author"` | Author name for tracked changes and comments |
| `MCP_AUTHOR_INITIALS` | `""` | Author initials for comments |
| `MCP_TRANSPORT` | `stdio` | Transport type: `stdio`, `sse`, or `streamable-http` |
| `MCP_HOST` | `0.0.0.0` | Host to bind (for SSE/HTTP transports) |
| `MCP_PORT` | `8000` | Port to bind (for SSE/HTTP transports) |

For remote deployment, see [RENDER_DEPLOYMENT.md](RENDER_DEPLOYMENT.md).

## Example Prompts

Just tell the AI what you want in plain language:

```
"Draft a contract with tracked changes so my colleague can review"
"Format all headings as Cambria 13pt bold and add automatic numbering"
"Add a comment on paragraph 3 asking about the deadline"
"Find every mention of 'ABC Corp' and replace with 'XYZ Ltd' as a tracked change"
"Set the page to A4 landscape with 2cm margins"
"Insert a table of contents based on the document headings"
"Add page numbers in the footer and our company name in the header"
"Insert a cross-reference to Heading 2 in paragraph 5"
```

## Usage Examples

### Example 1: Read a document (cross-platform)

**Tool call:** `get_document_text`
```json
{ "filename": "C:/Documents/report.docx" }
```
**Expected output:**
```json
{
  "status": "success",
  "paragraphs": [
    {"index": 0, "text": "Quarterly Report", "style": "Heading 1"},
    {"index": 1, "text": "Revenue increased by 15% compared to Q3.", "style": "Normal"},
    {"index": 2, "text": "Key Metrics", "style": "Heading 2"}
  ],
  "total_paragraphs": 3
}
```

### Example 2: Live editing with tracked changes (Windows)

**Tool call:** `word_live_replace_text`
```json
{
  "filename": "report.docx",
  "find_text": "ABC Corporation",
  "replace_text": "XYZ Ltd",
  "match_case": true,
  "replace_all": true,
  "track_changes": true
}
```
**Expected output:**
```json
{
  "status": "success",
  "replacements": 4,
  "message": "Replaced 4 occurrences (tracked changes enabled)"
}
```
The replacements appear as tracked changes in Word with strikethrough on "ABC Corporation" and underline on "XYZ Ltd".

### Example 3: Add a comment anchored to text (cross-platform)

**Tool call:** `add_comment`
```json
{
  "filename": "C:/Documents/contract.docx",
  "target_text": "payment within 30 days",
  "comment_text": "Should we extend this to 45 days?",
  "author": "Jane Smith"
}
```
**Expected output:**
```json
{
  "status": "success",
  "message": "Comment added by Jane Smith on 'payment within 30 days'"
}
```
The comment appears in Word's Review panel, anchored to the specified text.

## Tool Reference

**124 tools** across two modes — see the [complete tool reference](TOOLS.md) for details.

| Category | Count |
|----------|-------|
| Cross-platform (python-docx) | 80 |
| Windows Live (COM automation) | 44 |
| macOS Live (JXA automation) | 40 (of the 44 live tools) |

## Requirements

- **Python 3.11+**
- `python-docx`, `fastmcp`, `msoffcrypto-tool` (installed automatically)
- **Windows Live tools:** Windows 10/11 + Microsoft Word + `pywin32` (installed automatically)
- **macOS Live tools:** macOS + Microsoft Word for Mac (uses built-in JXA — no extra dependencies)

> The cross-platform tools work without Word installed — only python-docx is needed.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, code style, and how to add new tools.

Found a bug? [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md).
Have an idea? [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md).

## Acknowledgments

Built on top of [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) by GongRzhe (MIT License).

Additional libraries: [python-docx](https://python-docx.readthedocs.io/) &middot; [FastMCP](https://github.com/modelcontextprotocol/python-sdk) &middot; [pywin32](https://github.com/mhammond/pywin32)

## Privacy

This server runs entirely on your local machine. No data is collected, transmitted, or stored. See the full [Privacy Policy](PRIVACY.md).

## Support

- **Bug reports:** [Open an issue](https://github.com/ykarapazar/word-mcp-live/issues/new?template=bug_report.md)
- **Feature requests:** [Request a feature](https://github.com/ykarapazar/word-mcp-live/issues/new?template=feature_request.md)
- **Discussions:** [GitHub Discussions](https://github.com/ykarapazar/word-mcp-live/discussions)

## License

MIT License — see [LICENSE](LICENSE) for details.

## Star History

<a href="https://star-history.com/#ykarapazar/word-mcp-live&Date">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date&theme=dark" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date" />
   <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=ykarapazar/word-mcp-live&type=Date" />
 </picture>
</a>

<!-- mcp-name: io.github.ykarapazar/word-mcp-live -->
