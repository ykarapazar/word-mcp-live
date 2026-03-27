# Tool Reference

Complete list of all 115 tools provided by word-mcp-live.

---

## Cross-Platform Tools

These work on Windows, macOS, and Linux using python-docx. The document file must be **closed** (not open in Word).

<details>
<summary><b>Document Management</b></summary>

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
<summary><b>Content</b></summary>

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
<summary><b>Formatting</b></summary>

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
<summary><b>Comments</b></summary>

| Tool | Description |
|------|-------------|
| `get_all_comments` | Extract all comments |
| `get_comments_by_author` | Filter comments by author |
| `get_comments_for_paragraph` | Get comments for a specific paragraph |
| `add_comment` | Add a comment anchored to text |

</details>

<details>
<summary><b>Tracked Changes</b></summary>

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
<summary><b>Hyperlinks</b></summary>

| Tool | Description |
|------|-------------|
| `manage_hyperlinks` | Add, list, remove, and update hyperlinks |

</details>

<details>
<summary><b>Layout</b></summary>

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
<summary><b>Footnotes</b></summary>

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
<summary><b>Protection</b></summary>

| Tool | Description |
|------|-------------|
| `protect_document` | Add password protection |
| `unprotect_document` | Remove protection |
| `add_restricted_editing` | Restrict editing to specific sections |
| `add_digital_signature` | Add a digital signature |
| `verify_document` | Verify protection and signatures |

</details>

<details>
<summary><b>Extraction</b></summary>

| Tool | Description |
|------|-------------|
| `get_paragraph_text_from_document` | Get text from a specific paragraph |
| `find_text_in_document` | Find text occurrences |
| `get_highlighted_text` | Extract highlighted/colored text |
| `convert_to_pdf` | Convert to PDF |

</details>

---

## Windows Live Tools

These require Windows with Microsoft Word installed. They operate on documents **currently open in Word** via COM automation. Every tool call is a single Ctrl+Z entry in Word's undo stack.

<details open>
<summary><b>Editing</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_insert_text` | Insert text at a position (with optional tracked changes) |
| `word_live_insert_paragraphs` | Insert multiple paragraphs near a target (by text match or paragraph index) in one call |
| `word_live_delete_text` | Delete a character range |
| `word_live_replace_text` | Find & replace via COM — works across tracked change boundaries; supports wildcards |
| `word_live_format_text` | Format text (bold, italic, font, highlight, paragraph alignment, page break before) |
| `word_live_add_table` | Insert a table |
| `word_live_format_table` | Format an existing table |
| `word_live_apply_list` | Apply bullet, numbered, or multilevel list formatting |
| `word_live_setup_heading_numbering` | Auto-numbered headings (1. / 1.1) with configurable style |
| `word_live_modify_table` | Modify table structure: get info, set cell, set row, set range (bulk fill), add/delete rows/columns, merge cells, autofit, or delete table |
| `word_live_save` | Save document in place or save-as to a new path (docx, pdf, rtf, txt) |
| `word_live_toggle_track_changes` | Toggle or explicitly set track changes mode on/off |
| `word_live_insert_image` | Insert an image with sizing, alignment, wrapping, and optional border |
| `word_live_insert_cross_reference` | Insert a live cross-reference to headings, bookmarks, figures, tables, equations, footnotes, or endnotes |
| `word_live_insert_equation` | Insert a mathematical equation using UnicodeMath syntax |

</details>

<details open>
<summary><b>Reading</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_list_open` | List all documents currently open in Word with name, path, pages, and saved status |
| `word_live_get_text` | Get all text paragraph by paragraph |
| `word_live_get_page_text` | Get text from specific page(s) with char offsets for chaining |
| `word_live_get_paragraph_format` | Inspect paragraph formatting (font, spacing, alignment, list info, per-run detail) |
| `word_live_get_info` | Get document metadata (pages, words, sections) |
| `word_live_find_text` | Find text with context; supports wildcards |
| `word_live_get_undo_history` | List undo stack entries |
| `word_live_list_cross_reference_items` | List available cross-reference targets (headings, bookmarks, figures, tables) with indices |
| `word_live_diagnose_layout` | Scan for layout problems (keep_with_next chains, style misuse, break issues) |

</details>

<details open>
<summary><b>Comments & Revisions</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_comments` | Get all comments |
| `word_live_add_comment` | Add a comment anchored to text |
| `word_live_list_revisions` | List tracked changes |
| `word_live_reply_to_comment` | Add a threaded reply to an existing comment (Word 2016+) |
| `word_live_resolve_comment` | Mark a comment as resolved or unresolve it (Word 2016+) |
| `word_live_delete_comment` | Permanently delete a comment from the document |
| `word_live_accept_revisions` | Accept tracked changes (all or by author/type) |
| `word_live_reject_revisions` | Reject tracked changes (all or by author/type) |

</details>

<details open>
<summary><b>Layout</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_set_page_layout` | Set orientation, size, and margins |
| `word_live_add_header_footer` | Add header/footer text |
| `word_live_add_page_numbers` | Add page numbers |
| `word_live_add_section_break` | Add section break |
| `word_live_set_paragraph_spacing` | Set paragraph spacing |
| `word_live_add_bookmark` | Add a named bookmark |
| `word_live_add_watermark` | Add a text watermark |

</details>

<details open>
<summary><b>Undo & Screen Capture</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_undo` | Undo last N operations (each tool call = one undo entry) |
| `word_screen_capture` | Screenshot of the Word window |

</details>
