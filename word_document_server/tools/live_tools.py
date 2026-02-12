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
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

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
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)

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
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count
        if start_paragraph < 1 or end_paragraph > total_paras:
            return json.dumps({
                "error": f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
            })

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
) -> str:
    """[Windows only] Set up auto-numbered headings with multilevel list (1. / 1.1).

    Customizes Heading 1 and Heading 2 styles to Karapazar house style before applying:
    - Heading 1: Cambria 13pt Bold, #0D0D0D, justify, 18pt before, 6pt after, 1.15 line spacing
    - Heading 2: Cambria 11pt Bold, #0D0D0D, justify, 12pt before, 6pt after, 1.15 line spacing

    Creates a multilevel list template: Level 1 = "1." linked to Heading 1,
    Level 2 = "1.1" linked to Heading 2. Applies styles and numbering to the
    specified paragraphs, then optionally strips manual number prefixes
    (regex: ^\d+(\.\d+)*\.?\s+ — matches "1. ", "2.3 ", "10. ", etc.).

    Args:
        filename: Document name or path (None = active document).
        h1_paragraphs: List of 1-indexed paragraph numbers for Heading 1 (main sections).
        h2_paragraphs: List of 1-indexed paragraph numbers for Heading 2 (sub-sections).
        strip_manual_numbers: Remove leading "N." or "N.N" text from headings (default True).

    Returns:
        JSON with h1_applied, h2_applied, and stripped counts.
    """
    import re

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools only on Windows"})

    if not h1_paragraphs and not h2_paragraphs:
        return json.dumps({"error": "Provide h1_paragraphs and/or h2_paragraphs"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        color_int = 0x0D + (0x0D << 8) + (0x0D << 16)  # #0D0D0D in Word RGB

        # --- Customize Heading 1 style ---
        s1 = doc.Styles(-2)  # wdStyleHeading1
        s1.Font.Name = "Cambria"
        s1.Font.Size = 13
        s1.Font.Bold = True
        s1.Font.Italic = False
        s1.Font.Color = color_int
        s1.ParagraphFormat.Alignment = 3  # justify
        s1.ParagraphFormat.SpaceBefore = 18
        s1.ParagraphFormat.SpaceAfter = 6
        s1.ParagraphFormat.LineSpacingRule = 5  # multiple
        s1.ParagraphFormat.LineSpacing = 13.8
        s1.ParagraphFormat.KeepWithNext = True
        s1.ParagraphFormat.KeepTogether = False

        # --- Customize Heading 2 style ---
        s2 = doc.Styles(-3)  # wdStyleHeading2
        s2.Font.Name = "Cambria"
        s2.Font.Size = 11
        s2.Font.Bold = True
        s2.Font.Italic = False
        s2.Font.Color = color_int
        s2.ParagraphFormat.Alignment = 3
        s2.ParagraphFormat.SpaceBefore = 12
        s2.ParagraphFormat.SpaceAfter = 6
        s2.ParagraphFormat.LineSpacingRule = 5
        s2.ParagraphFormat.LineSpacing = 13.8
        s2.ParagraphFormat.KeepWithNext = True
        s2.ParagraphFormat.KeepTogether = False

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
        from word_document_server.core.word_com import get_word_app, find_document

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
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)
        rng = doc.Range(start, end)
        deleted_text = rng.Text

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
