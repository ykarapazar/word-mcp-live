# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[1.4.1]: https://github.com/ykarapazar/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/ykarapazar/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ykarapazar/word-mcp-live/releases/tag/v1.0.0
