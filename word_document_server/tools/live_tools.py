"""Live editing tools for Microsoft Word via COM automation.

These tools operate on documents that are currently open in Word,
providing real-time editing capabilities with optional tracked changes.
"""

import json
import os
import sys

from word_document_server.defaults import DEFAULT_AUTHOR

# Word COM constants
WD_STORY = 6


async def word_live_insert_text(
    filename: str = None,
    text: str = "",
    position: str = "end",
    bookmark: str = None,
    track_changes: bool = False,
) -> str:
    """Insert text into an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Text to insert.
        position: "start", "end", "cursor", or character offset as string.
        bookmark: Insert after a named bookmark (overrides position).
        track_changes: Track the insertion as a revision.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Insert Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                if bookmark:
                    if not doc.Bookmarks.Exists(bookmark):
                        return json.dumps({"error": f"Bookmark '{bookmark}' not found"})
                    doc.Bookmarks(bookmark).Range.InsertAfter(text)
                elif position == "start":
                    doc.Range(0, 0).InsertBefore(text)
                elif position == "end":
                    end_pos = doc.Content.End - 1
                    doc.Range(end_pos, end_pos).InsertAfter(text)
                elif position == "cursor":
                    app.Selection.TypeText(text)
                else:
                    try:
                        offset = int(position)
                        doc.Range(offset, offset).InsertBefore(text)
                    except ValueError:
                        return json.dumps(
                            {
                                "error": f"Invalid position: {position}. "
                                "Use 'start', 'end', 'cursor', or a character offset."
                            }
                        )
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "text_length": len(text),
                "position": position,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_format_text(
    filename: str = None,
    start: int = None,
    end: int = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    highlight_color: int = None,
    style_name: str = None,
    paragraph_alignment: str = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Format text in an open Word document: font, color, highlight, style, alignment.
    Use this tool for any visual/formatting change that does NOT alter the text content itself.

    Args:
        filename: Document name or path (None = active document).
        start: Start character position (required). Use word_live_find_text to get positions.
        end: End character position (required).
        bold: Set bold (True/False).
        italic: Set italic (True/False).
        underline: Set underline (True/False).
        font_name: Font family (e.g., "Arial", "Times New Roman").
        font_size: Font size in points (e.g., 12).
        font_color: Text color as "#RRGGBB" hex (e.g., "#FF0000" for red).
        highlight_color: Text highlight background color index.
            0 = remove highlight, 1 = black, 2 = blue, 3 = turquoise,
            4 = bright green, 5 = pink, 6 = red, 7 = yellow,
            8 = white, 9 = dark blue, 10 = teal, 11 = green,
            12 = violet, 13 = dark red, 14 = dark yellow, 15 = gray, 16 = light gray.
            Common: 7=yellow (add), 0=none (remove).
        style_name: Apply a named Word style (e.g., "Heading 1", "Normal").
        paragraph_alignment: Paragraph alignment — "left" (0), "center" (1), "right" (2), "justify" (3).
            Applies to ALL paragraphs in the selected range.
        track_changes: Track formatting changes as revisions.

    Note: To change formatting without changing text (e.g., remove highlight,
    change font), use this tool. For text content changes, use track_replace
    or word_live_insert_text instead.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start is None or end is None:
        return json.dumps(
            {"error": "Both 'start' and 'end' character positions are required"}
        )

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)

        with undo_record(app, "MCP: Format Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                if bold is not None:
                    rng.Font.Bold = bold
                if italic is not None:
                    rng.Font.Italic = italic
                if underline is not None:
                    rng.Font.Underline = 1 if underline else 0
                if font_name is not None:
                    rng.Font.Name = font_name
                if font_size is not None:
                    rng.Font.Size = font_size
                if font_color is not None:
                    c = font_color.lstrip("#")
                    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                    rng.Font.Color = r + (g << 8) + (b << 16)
                if highlight_color is not None:
                    rng.HighlightColorIndex = highlight_color
                if style_name is not None:
                    rng.Style = style_name
                if paragraph_alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    al = align_map.get(paragraph_alignment.lower())
                    if al is None:
                        return json.dumps({"error": f"Invalid alignment: {paragraph_alignment}. Use: left, center, right, justify"})
                    for para in rng.Paragraphs:
                        para.Format.Alignment = al
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = rng.Text
        if len(preview) > 50:
            preview = preview[:50] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "range": f"{start}-{end}",
                "text_preview": preview,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_apply_list(
    filename: str = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    list_type: str = "bullet",
    level: int = 0,
    remove: bool = False,
    continue_previous: bool = False,
    track_changes: bool = False,
) -> str:
    """[Windows only] Apply or remove bullet/numbered list formatting on paragraphs in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph to format (1-indexed, required).
        end_paragraph: Last paragraph to format (1-indexed, defaults to start_paragraph).
        list_type: "bullet" for bullet list, "number" for numbered list.
        level: Indentation level (0 = first level, 1 = second level, etc.).
        remove: If True, removes list formatting from the range.
        continue_previous: If True, continues numbering from a previous list above
            (useful when bullets interrupt a numbered list, e.g. items 1-4, then bullets,
            then item 5 should continue as 5 not restart at 1).
        track_changes: Track changes as revisions.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start_paragraph is None:
        return json.dumps({"error": "start_paragraph is required (1-indexed)"})

    if end_paragraph is None:
        end_paragraph = start_paragraph

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count
        if start_paragraph < 1 or end_paragraph > total_paras:
            return json.dumps({
                "error": f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
            })

        with undo_record(app, "MCP: Apply List"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # Word COM ListGallery constants:
                # wdBulletGallery = 1, wdNumberGallery = 2, wdOutlineNumberGallery = 3
                gallery_map = {"bullet": 1, "number": 2}
                formatted = 0

                for i in range(start_paragraph, end_paragraph + 1):
                    para = doc.Paragraphs(i)
                    if remove:
                        para.Range.ListFormat.RemoveNumbers()
                    else:
                        gallery_idx = gallery_map.get(list_type, 1)
                        template = doc.Application.ListGalleries(gallery_idx).ListTemplates(1)
                        should_continue = (i > start_paragraph) or continue_previous
                        para.Range.ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate=template,
                            ContinuePreviousList=should_continue,
                            DefaultListBehavior=1,  # wdWord2003
                        )
                        if level > 0:
                            para.Range.ListFormat.ListLevelNumber = level + 1
                    formatted += 1
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        action = "removed" if remove else f"applied {list_type}"
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "action": action,
            "paragraphs": f"{start_paragraph}-{end_paragraph}",
            "count": formatted,
            "level": level,
            "tracked": track_changes,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_setup_heading_numbering(
    filename: str = None,
    h1_paragraphs: list = None,
    h2_paragraphs: list = None,
    strip_manual_numbers: bool = True,
    font_name: str = None,
    h1_size: float = None,
    h2_size: float = None,
    bold: bool = None,
    alignment: str = None,
    font_color: str = None,
    h1_space_before: float = None,
    h1_space_after: float = None,
    h2_space_before: float = None,
    h2_space_after: float = None,
    line_spacing: float = None,
) -> str:
    """[Windows only] Set up auto-numbered headings with multilevel list (1. / 1.1).

    Creates a multilevel list template: Level 1 = "1." linked to Heading 1,
    Level 2 = "1.1" linked to Heading 2. Applies styles and numbering to the
    specified paragraphs, then optionally strips manual number prefixes
    (regex: ^\d+(\.\d+)*\.?\s+ — matches "1. ", "2.3 ", "10. ", etc.).

    If any style parameter is provided, Heading 1 and Heading 2 styles are
    customized before applying. If no style params are given, only numbering
    is applied (existing styles are preserved).

    Args:
        filename: Document name or path (None = active document).
        h1_paragraphs: List of 1-indexed paragraph numbers for Heading 1 (main sections).
        h2_paragraphs: List of 1-indexed paragraph numbers for Heading 2 (sub-sections).
        strip_manual_numbers: Remove leading "N." or "N.N" text from headings (default True).
        font_name: Font family for both heading styles (e.g., "Cambria").
        h1_size: Font size in points for Heading 1 (e.g., 13).
        h2_size: Font size in points for Heading 2 (e.g., 11).
        bold: Set bold on both heading styles (True/False).
        alignment: Paragraph alignment — "left", "center", "right", "justify".
        font_color: Text color as "#RRGGBB" hex (e.g., "#0D0D0D").
        h1_space_before: Space before Heading 1 in points (e.g., 18).
        h1_space_after: Space after Heading 1 in points (e.g., 6).
        h2_space_before: Space before Heading 2 in points (e.g., 12).
        h2_space_after: Space after Heading 2 in points (e.g., 6).
        line_spacing: Line spacing in points for both heading styles (e.g., 13.8 for 1.15x).

    Returns:
        JSON with h1_applied, h2_applied, and stripped counts.
    """
    import re

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools only on Windows"})

    if not h1_paragraphs and not h2_paragraphs:
        return json.dumps({"error": "Provide h1_paragraphs and/or h2_paragraphs"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Setup Heading Numbering"):
            # --- Optionally customize heading styles ---
            has_style_params = any(p is not None for p in [
                font_name, h1_size, h2_size, bold, alignment, font_color,
                h1_space_before, h1_space_after, h2_space_before, h2_space_after,
                line_spacing,
            ])

            if has_style_params:
                align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                align_val = align_map.get(alignment.lower()) if alignment else None

                color_int = None
                if font_color:
                    c = font_color.lstrip("#")
                    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                    color_int = r + (g << 8) + (b << 16)

                for style_id, size, sp_before, sp_after in [
                    (-2, h1_size, h1_space_before, h1_space_after),
                    (-3, h2_size, h2_space_before, h2_space_after),
                ]:
                    s = doc.Styles(style_id)
                    if font_name is not None:
                        s.Font.Name = font_name
                    if size is not None:
                        s.Font.Size = size
                    if bold is not None:
                        s.Font.Bold = bold
                        s.Font.Italic = False
                    if color_int is not None:
                        s.Font.Color = color_int
                    if align_val is not None:
                        s.ParagraphFormat.Alignment = align_val
                    if sp_before is not None:
                        s.ParagraphFormat.SpaceBefore = sp_before
                    if sp_after is not None:
                        s.ParagraphFormat.SpaceAfter = sp_after
                    if line_spacing is not None:
                        s.ParagraphFormat.LineSpacingRule = 5  # multiple
                        s.ParagraphFormat.LineSpacing = line_spacing
                    s.ParagraphFormat.KeepWithNext = True
                    s.ParagraphFormat.KeepTogether = False

            # --- Create multilevel list template ---
            lt = doc.ListTemplates.Add(OutlineNumbered=True)

            # Level 1: "1." linked to Heading 1
            lv1 = lt.ListLevels(1)
            lv1.NumberFormat = "%1."
            lv1.NumberStyle = 0  # wdListNumberStyleArabic
            lv1.StartAt = 1
            lv1.Alignment = 0  # left
            lv1.NumberPosition = 0
            lv1.TextPosition = 28  # ~1cm indent for text after number
            lv1.TabPosition = 28
            lv1.LinkedStyle = "Heading 1"

            # Level 2: "1.1" linked to Heading 2
            lv2 = lt.ListLevels(2)
            lv2.NumberFormat = "%1.%2"
            lv2.NumberStyle = 0
            lv2.StartAt = 1
            lv2.Alignment = 0
            lv2.NumberPosition = 0
            lv2.TextPosition = 28
            lv2.TabPosition = 28
            lv2.LinkedStyle = "Heading 2"

            # --- Apply styles to paragraphs ---
            h1_applied = 0
            h2_applied = 0

            all_heading_paras = []
            for idx in (h1_paragraphs or []):
                all_heading_paras.append((idx, -2))  # wdStyleHeading1
            for idx in (h2_paragraphs or []):
                all_heading_paras.append((idx, -3))  # wdStyleHeading2
            all_heading_paras.sort(key=lambda x: x[0])

            for para_idx, style_id in all_heading_paras:
                if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                    continue
                para = doc.Paragraphs(para_idx)
                para.Style = doc.Styles(style_id)
                if style_id == -2:
                    h1_applied += 1
                else:
                    h2_applied += 1

            # --- Apply list template to all heading paragraphs ---
            for para_idx, _ in all_heading_paras:
                if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                    continue
                para = doc.Paragraphs(para_idx)
                para.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=lt,
                    ContinuePreviousList=True,
                    DefaultListBehavior=1,
                )

            # --- Strip manual numbers ---
            stripped = 0
            if strip_manual_numbers:
                for para_idx, _ in all_heading_paras:
                    if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                        continue
                    para = doc.Paragraphs(para_idx)
                    text = para.Range.Text.rstrip("\r\x07")
                    # Match patterns: "1. ", "1.1 ", "2.3 ", "10. ", "10.2 ", etc.
                    m = re.match(r"^\d+(\.\d+)*\.?\s+", text)
                    if m:
                        # Delete the matched prefix
                        prefix_len = len(m.group(0))
                        rng = doc.Range(para.Range.Start, para.Range.Start + prefix_len)
                        rng.Delete()
                        stripped += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "h1_applied": h1_applied,
            "h2_applied": h2_applied,
            "stripped": stripped,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_table(
    filename: str = None,
    rows: int = 2,
    cols: int = 2,
    position: str = "end",
    data: list = None,
    track_changes: bool = False,
) -> str:
    """Add a table to an open Word document.

    Args:
        filename: Document name or path.
        rows: Number of rows.
        cols: Number of columns.
        position: "start", "end", or character offset.
        data: Optional 2D list of cell data.
        track_changes: Track as revision.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
        else:
            try:
                offset = int(position)
                rng = doc.Range(offset, offset)
            except ValueError:
                return json.dumps({"error": f"Invalid position: {position}"})

        with undo_record(app, "MCP: Add Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                table = doc.Tables.Add(rng, rows, cols)
                if data:
                    for r_idx, row_data in enumerate(data):
                        if r_idx >= rows:
                            break
                        for c_idx, cell_val in enumerate(row_data):
                            if c_idx >= cols:
                                break
                            table.Cell(r_idx + 1, c_idx + 1).Range.Text = str(cell_val)
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "rows": rows,
                "cols": cols,
                "position": position,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_text(
    filename: str = None,
    start: int = None,
    end: int = None,
    track_changes: bool = False,
) -> str:
    """Delete text from an open Word document.

    Args:
        filename: Document name or path.
        start: Start character position.
        end: End character position.
        track_changes: Track deletion as a revision.

    Returns:
        JSON with deleted text info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start is None or end is None:
        return json.dumps(
            {"error": "Both 'start' and 'end' character positions are required"}
        )

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)
        deleted_text = rng.Text

        with undo_record(app, "MCP: Delete Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                rng.Delete()
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = deleted_text
        if len(preview) > 100:
            preview = preview[:100] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "deleted_text": preview,
                "range": f"{start}-{end}",
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_undo(
    filename: str = None,
    times: int = 1,
) -> str:
    """[Windows only] Undo the last N operations in an open Word document.

    Each MCP destructive tool call is grouped as a single undo entry (e.g.,
    "MCP: Insert Text"). Calling undo(times=1) reverts the last MCP operation;
    undo(times=3) reverts the last three.

    Args:
        filename: Document name or path (None = active document).
        times: Number of undo steps (default 1).

    Returns:
        JSON with success status and number of undone steps.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if times < 1:
        return json.dumps({"error": "times must be >= 1"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = doc.Undo(times)

        return json.dumps({
            "success": bool(result),
            "document": doc.Name,
            "times_requested": times,
            "undo_result": bool(result),
        })

    except Exception as e:
        return json.dumps({"error": str(e)})
