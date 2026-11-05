# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.6.22] - Unreleased

### Fixed
- **`uvx word-mcp-live` failed** ("An executable named `word-mcp-live` is not provided by package `word-mcp-live`") because the only console script was `word_mcp_server`. Added a `word-mcp-live` console script (same entry point); `word_mcp_server` is kept for compatibility.
- **MCP `serverInfo.version` reported the fastmcp version (4.x)** instead of the package version. It now reports the package version.
- **macOS: wrong document** — `word_live_find_text`, `word_live_replace_text` and `word_live_insert_text` with `position="cursor"` used the front window's selection, so targeting a background document searched or edited the front one. They now activate the target and restore the previously active document. `word_live_add_comment` always used the active document and ignored `author`; it now targets the named document and sets the comment author.
- **macOS control-character guard** — the `\x07`/control-byte check added in 1.6.0 for `word_live_insert_text`, `word_live_replace_text` and `word_live_find_text` only ran on Windows because the macOS branch returned first. It now runs before the platform branch.
- **`word_live_insert_text` (macOS)** — a numeric `position` was interpolated unescaped into the JXA script; it is now parsed as an integer in Python, and non-numeric values return an error.
- `SyntaxWarning: invalid escape sequence` in the `word_live_setup_heading_numbering` docstring (now a raw string; description text unchanged).
- `test_formatting.py` now runs under pytest (it was an `async def` test with no async plugin).
- **`word_live_setup_heading_numbering` (macOS)** failed with `NameError: name 're' is not defined`; the missing import is added.
- **macOS: two open documents with the same name** (from different folders) — Word for Mac binds ranges to the document name, so edits and finds could hit the other file even when a full path was given. Tools now refuse with a clear error until only one is open.
- **macOS: Word not responding** — an osascript timeout surfaced as a raw `TimeoutExpired` traceback. It now returns a clear message (Word may be showing a dialog or still busy with the previous request).
- **Error results now set `isError`.** Live tools that returned `{"error": ...}` JSON reported success at the MCP level; the error JSON is now raised as a tool error (same text).
- **Windows: one Ctrl+Z per tool call even when the target document is not the active one.** Word only groups edits (`UndoRecord`) on the active document's undo stack, so a call on a background document could take a dozen undos (12 in our test, now 1). The target is activated for the call and the previous document restored.
- **Windows: `MCP_AUTHOR` was ignored.** Tracked changes and comments carried the signed-in Office name. They now use `MCP_AUTHOR` (default `Author`), and comment results report the author Word stored. `reject_revisions`/`accept_revisions` by that author now match (rejected 0 → 3 in our test). Note: users who never set `MCP_AUTHOR` now get `Author` instead of their Office name.
- **`MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` from `.env` were ignored**: the defaults were read before `.env` was loaded.
- **Windows: two open documents with the same file name** — a basename silently picked one of them, and a full path could still return the other. A full path now selects the exact document; an ambiguous basename returns an error listing the paths.
- **Windows: clearer errors** — a protected document (Restrict Editing) returned a raw COM tuple; an open dialog box returned `<unknown>.Name`; Word closing mid-session returned an RPC error. Each now returns a plain-language message. `word_live_save` on a document opened read-only returns an error at once; before, the call hung until timeout (300 s in testing) and blocked later calls.
- **Windows: page counts** (`get_info`, `list_open`, `get_text`, `get_page_text`) reported the pages laid out so far (223 of 290 on a 5,000-paragraph document); layout is now completed first.
- **Control characters** are now also rejected in `word_live_insert_paragraphs`, `word_live_add_comment` and `word_live_reply_to_comment` (a `\x07` comment was stored as-is).
- **Tool annotations**: `word_screen_capture` (writes or overwrites `output_path`) and `word_live_take_snapshot` / `word_live_get_diff` (store snapshot state) were marked read-only; they no longer are.
- **macOS: `word_live_apply_list`** (bullet/number) failed with `undefined is not an object (evaluating 'gallery.listTemplates')`; it now uses the list galleries Word for Mac exposes.
- **macOS: `word_live_add_section_break`** failed with -1708, and `break_type` values (`new_page`, `continuous`, `even_page`, `odd_page`) were passed through unmapped. Both fixed.
- **macOS: `word_live_add_table`** failed with -1708 whenever another document was in front (Word for Mac only creates tables in the active document), and with -1728 when the document already ended with a table. The target is activated for the call, and a paragraph separates the new table from a trailing one. Offsets inside an existing table are refused, as on Windows.
- **macOS: `word_live_get_paragraph_format`** returned no `style` and ignored `include_runs`; both now work (runs in the Windows shape).
- **macOS: read-only documents** — edits and save returned a bare -1708; they now say the document is open read-only.

### Changed
- **Windows speed** (same output): reading paragraphs is O(n) instead of O(n²) COM calls, and `word_live_get_page_text` walks only the requested page. On a 5,000-paragraph document: `take_snapshot` 133 s → 0.14 s, `get_diff` 137 s → 0.14 s, `get_page_text` for the last page 140 s → 0.14 s; `take_snapshot` on 1,000 paragraphs 9.3 s → 0.06 s.
- **Live tools run on one dedicated worker thread** instead of blocking the server's event loop, so pings and cancellation are answered while a long Word call runs.
- **macOS speed** (same output): `word_live_get_paragraph_format` and `word_live_diagnose_layout` no longer resolve the whole paragraph collection (about 3.5× and 2.6× faster on a 400-paragraph document; a 200-paragraph range of a 5,000-paragraph document now takes about 17 s where it used to time out); `word_live_get_page_text` locates pages with `goToNext` instead of a binary search (page 50 of a 5,000-paragraph document now takes 1.6 s where it used to time out); `word_live_find_text` sets its find options once per call.
- Tool descriptions: 32 live tools that work on macOS were labelled "[Windows only]"; they now say "[Requires Word open — Windows or macOS]". The 13 live tools with no macOS path are labelled "[Windows only]". No tool names or parameters changed.
- Packaging: `pytest` moved from runtime dependencies to the `dev` dependency group; `python-dotenv` (imported by `main.py`) is now declared; `requirements.txt` mirrors the `pyproject.toml` runtime dependencies.
- `manifest.json`: version, tool count, `compatibility.platforms` (`darwin`, `win32`) and `privacy_policies`. `server.json`: version, tool count, and env var default descriptions matching the code (`MCP_AUTHOR` defaults to `Author`, initials empty).
- `word_document_server.__version__` was still `1.2.0`; now matches the package version.
- README / TOOLS.md: tool counts corrected to 120 (75 cross-platform + 44 `word_live_*` + `word_screen_capture`; 32 of the 45 live tools on macOS), macOS feature table corrected, missing `word_live_set_core_properties` added to TOOLS.md, OpenAI Codex install instructions added.

## [1.6.21] - 2026-10-06

### Fixed
- **`word_live_format_text` (macOS)** — `highlight_color=0` was skipped as falsy and integer colours were passed as invalid JXA enum strings (silently clearing highlight). Values 0–16 now map to WdColorIndex names; invalid values are rejected.
- macOS parity: `word_live_modify_table` `add_row` / `delete_row` / `set_row` and multi-paragraph `set_cell`; `word_live_add_table` via `createRange` + `textObject`; `NameError` in `word_live_set_page_layout`; `word_live_add_page_numbers` via `createNewField`.

### Changed
- Dependencies: `fastmcp>=4.0.11`, `python-docx>=1.2.0`, `msoffcrypto-tool>=6.0.0`; `uv.lock` refreshed.

## [1.6.20] - 2026-05-29

### Fixed
- `word_live_setup_heading_numbering` (macOS) — applies the list template to each heading paragraph explicitly, since `linkedStyle` assignment is unreliable in JXA.

## [1.6.19] - 2026-05-29

### Added
- `outline_numbered=false` option on `word_live_apply_list`, for lists that must stay independent of the heading outline numbering.

## [1.6.17] - 2026-05-28

### Added
- `word_live_setup_heading_numbering` accepts `heading_map` (`{text: level}`, same shape as `word_live_apply_list`).

### Changed
- List-template levels are linked to `Heading 1` / `Heading 2` styles instead of `Normal`; Level 2 restarts after each Level 1.

## [1.6.16] - 2026-05-28

### Fixed
- VBA escaping for the macOS header/footer list cleanup.

## [1.6.15] - 2026-05-28

### Fixed
- macOS header/footer list cleanup now uses `do Visual Basic`, because JXA header/footer accessors are unreliable on Word for Mac.

## [1.6.14] - 2026-05-28

### Fixed
- `word_live_apply_list` (macOS): no longer resets the Level 1 counter (duplicate numbers); clears runaway list formatting from headers/footers; removes numbering with one range call instead of one per paragraph.

### Changed
- `word_live_get_text` (macOS) reads the whole document in one call (much faster on large or OneDrive-backed documents).

Versions 1.6.7–1.6.13 and 1.6.18 were never published.

## [1.6.6] - 2026-05-28

### Changed
- `word_live_apply_list` (macOS) reads paragraph texts in one batched call; JXA timeout raised to 600 s.

## [1.6.5] - 2026-05-28

### Fixed
- Smart quote in the macOS `apply_list` JXA code that caused "Invalid character" errors.

## [1.6.4] - 2026-05-28

### Fixed
- `font_color` was not exposed in the `word_live_apply_list` MCP schema.

## [1.6.3] - 2026-05-28

### Added
- `word_live_apply_list` (macOS): multi-level lists with `level_map` text-to-level matching, `font_color` override, automatic `(a)`/`(b)` Level 3 detection and manual-prefix stripping.

### Fixed
- macOS: JXA save-as-PDF for `convert_to_pdf`, auto-truncation for large documents in `word_live_get_text`, and `word_live_get_page_text` reliability fixes.

## [1.6.2] - 2026-04-29

### Fixed
- Added `Pillow` as a Windows runtime dependency for `word_screen_capture` (failed in isolated `uvx` environments).

## [1.6.1] - 2026-04-29

### Fixed
- `word_screen_capture` raised `NameError: _MAC_AVAILABLE` on Windows.

## [1.6.0] - 2026-04-29

### Added
- **`word_live_set_core_properties`** — set Word document Title, Subject, Author, Keywords, Comments, Category, Manager, Company, Last Author via `Document.BuiltInDocumentProperties`. Wrapped in `undo_record` so a single Ctrl+Z reverts every property in one call. Equivalent to File > Info > Properties in the Word UI.
- New `word_document_server/utils/text_safety.py` — shared `reject_control_chars()` validator for Find/Replace/Insert text inputs.
- `scrub_orphans` parameter on `word_live_modify_table` `delete_table` operation (default `True`) — cleans stranded `\x07` cell-separator bytes the Word COM `Table.Delete()` occasionally leaves behind.

### Fixed
- **`word_live_replace_text` data-loss vector** — passing `\x07` (cell separator) as `find_text` previously matched across cell boundaries and could delete entire documents. Control bytes (U+0000–U+001F except `\t`, `\n`, `\r`) now rejected with a descriptive error before Find.Execute is reached.
- **`word_live_find_text`** — same control-byte protection applied to `search_text`.
- **`word_live_insert_text`** — same control-byte protection applied to `text` (prevents inserting orphan cell separators outside a real table).
- **`word_live_modify_table` `delete_table`** — leftover `\x07` separators after Word's native `Table.Delete()` now scrubbed by default (configurable via `scrub_orphans=False`).
- **`word_live_add_table`** — rejects `position` offset that falls inside an existing table's range or sits immediately after an orphan cell separator (would otherwise silently merge new content into existing/residual table structure).
- **`word_live_setup_heading_numbering`** — paragraphs that previously kept a custom template style (e.g. `Font Style30/31`) after a forced heading reassignment now (a) get explicit per-paragraph style assignment with try/except, (b) receive the same font/size/bold/color customizations as direct formatting so visual output matches even when the underlying style refuses to swap, (c) report any failed reassignments under a new `restyle_failures` field in the response.
- **`word_live_modify_table`** — re-reads `Tables.Count` per call and validates `table_index` with a helpful message ("table_index N out of range. Document has K table(s)…") instead of throwing "Document has no tables" when a stale index is passed after a prior delete.
- **`word_live_list_open`** — defensive per-document property access; one document in a broken COM proxy state no longer aborts the whole call. Each document entry now includes `index`, `track_revisions`, and per-property `errors` array.
- **`word_live_find_text`** — defensive `Range.Text` / `Range.Start` / `Range.End` / `Document.Name` access via internal `_safe_attr` helper. Transient COM marshalling failures after MCP reconnect now produce partial matches with `<unreadable>` placeholders and a `partial_errors` array, rather than aborting the call.

## [1.5.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — infinite loop when wildcard pattern matches zero-length strings (e.g., `*` alone); now skips forward on zero-length matches and enforces 50K replacement safety ceiling

## [1.5.0] - 2026-04-08

### Added
- **macOS live editing support** via JavaScript for Automation (JXA) — 33 of 41 `word_live_*` tools now work on macOS with Word for Mac
- New module `word_document_server/core/word_mac.py` — JXA bridge with 30+ functions for Word for Mac automation
- Platform auto-detection: same tool names and parameters on both Windows and macOS
- `pywin32` as conditional dependency (Windows only) in `pyproject.toml`

### Changed
- All `print()` calls in `main.py` redirected to stderr — fixes MCP stdio protocol corruption that prevented the server from loading in some clients
- All live tool functions now dispatch to macOS JXA implementations when `sys.platform == "darwin"`
- Updated tool count: 76 cross-platform + 41 Windows Live + 33 macOS Live

### Not Available on macOS
These 4 tools require Windows COM APIs with no AppleScript/JXA equivalent:
- `word_live_get_undo_history` — undo stack inspection not exposed in Word for Mac's scripting dictionary
- `word_live_reply_to_comment` — threaded comment replies not in AppleScript dictionary
- `word_live_resolve_comment` — comment Done property not in AppleScript dictionary
- `word_live_add_watermark` — requires VBA `Shapes.AddTextEffect` (VBA bridge killed by Apple sandboxing in Word 365)

## [1.4.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — `^s` (non-breaking space) now converted to `\u00a0` in replacement text (#4)

## [1.4.0] - 2026-04-08

### Added
- `word_live_insert_paragraphs` — insert multiple paragraphs near a target (by text or index) in a single undo record
- `word_live_take_snapshot` — store paragraph baseline for efficient change detection
- `word_live_get_diff` — compare current document against snapshot, returns only changed paragraphs
- `word_live_snapshot_status` — check snapshot existence and age
- `word_live_modify_table` — new `set_row` and `set_range` operations for bulk cell updates

### Fixed
- `word_live_replace_text` — infinite loop when document has TrackRevisions enabled independently of `track_changes` parameter (#7)
- All destructive tools now unconditionally restore `doc.TrackRevisions` in `finally` block

### Credits
- Snapshot/diff tools, `insert_paragraphs`, and bulk table operations adapted from PR #5 by @FarhadGSRX

## [1.3.0] - 2026-02-28

### Added
- `word_live_modify_table` — table operations via COM: get info, set cell, add/delete rows/columns, merge cells, autofit, delete table
- `word_live_save` — save document in place or save-as (docx, pdf, rtf, txt)
- `word_live_toggle_track_changes` — toggle or explicitly set track changes mode on/off
- `word_live_insert_image` — insert image with sizing, alignment, wrapping, and optional border
- `word_live_insert_cross_reference` — insert live cross-references to headings, bookmarks, figures, tables, equations, footnotes, endnotes
- `word_live_list_cross_reference_items` — list available cross-reference targets with their indices
- `word_live_insert_equation` — insert mathematical equations using UnicodeMath syntax
- `word_live_reply_to_comment` — threaded comment replies (Word 2016+)
- `word_live_resolve_comment` — mark comments as resolved/unresolved (Word 2016+)
- `word_live_delete_comment` — permanently delete a comment
- Total tool count now **114** (75 cross-platform + 39 Windows Live)

### Changed
- `word_live_delete_text` — now table-aware: deletes table objects within range before text deletion
- `word_live_insert_text` — auto-chunks text >30K chars to avoid COM 32K limit
- `word_live_setup_heading_numbering` — handles inflated paragraph ranges from comment anchors
- `word_live_modify_table` set_cell operation now accepts tracked changes before writing to prevent layered content

## [1.2.0] - 2025-02-15

### Added
- `word_live_replace_text` — find & replace via COM that works across tracked change boundaries; supports wildcards (`^m`, `^t`, `^p`) and tracked changes mode
- `word_live_diagnose_layout` — read-only scan for layout problems: keep_with_next chains, heading styles on body text, PageBreakBefore misuse, manual breaks
- `word_live_get_paragraph_format` — inspect paragraph formatting (font, spacing, alignment, list info, style); `include_runs=True` for per-run detail
- `word_live_get_page_text` — read text from specific page(s) with char offsets for chaining into format/edit tools
- `word_live_get_undo_history` — list undo stack entries
- `word_live_apply_list` — apply bullet, numbered, or multilevel list formatting
- `word_live_setup_heading_numbering` — auto-numbered headings (1. / 1.1) via multilevel list linked to Heading styles; configurable style params (font, size, color, spacing)

### Changed
- `word_live_format_text` — added `paragraph_alignment`, `page_break_before`, paragraph-index addressing (`start_paragraph`/`end_paragraph`), `preserve_direct_formatting` for style changes
- `word_live_find_text` — added `use_wildcards` for `^m`/`^t`/`^p`/Word wildcard syntax; `context_chars` now configurable (default 60, was 30)
- `word_live_set_paragraph_spacing` — clarified that `line_spacing` is in points (1.15 lines = 13.8pt)

## [1.1.0] - 2025-01-10

### Added
- 27 Windows Live tools (`word_live_*`) using COM automation for editing documents open in Word
- Per-operation undo system — all destructive tools wrapped with `UndoRecord`; each tool call = one Ctrl+Z entry
- `word_live_undo` — programmatic undo of last N operations
- Live editing tools: `word_live_insert_text`, `word_live_delete_text`, `word_live_format_text`, `word_live_add_table`
- Live reading tools: `word_live_get_text`, `word_live_get_info`, `word_live_find_text`
- Live comment & revision tools: `word_live_add_comment`, `word_live_get_comments`, `word_live_list_revisions`, `word_live_accept_revisions`, `word_live_reject_revisions`
- Live layout tools: `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark`
- `word_screen_capture` — screenshot of the Word window
- Cross-platform tracked changes: `track_replace`, `track_insert`, `track_delete`, `list_tracked_changes`, `accept_tracked_changes`, `reject_tracked_changes`
- Cross-platform comments: `add_comment` anchored to text
- Cross-platform hyperlinks: `manage_hyperlinks` (add, list, remove, update)
- Cross-platform layout tools: `set_page_layout`, `add_header_footer`, `add_page_numbers`, `add_section_break`, `set_paragraph_spacing`, `add_bookmark`, `add_watermark`
- Cross-platform footnote tools (10): add, delete, validate, customize footnotes and endnotes
- Cross-platform protection tools: `protect_document`, `unprotect_document`, `add_restricted_editing`, `add_digital_signature`, `verify_document`
- Multiple transport support: stdio (default), SSE, streamable-http
- `MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` environment variables for author metadata
- PyPI packaging as `word-mcp-live`

## [1.0.0] - 2024-12-01

### Added
- Initial release based on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
- 54 cross-platform tools using python-docx
- Document management, content editing, formatting, tables, extraction
- FastMCP server with stdio transport

[1.5.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.1...v1.5.0
[1.4.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ykarapazar/word-mcp-live/releases/tag/v1.0.0
