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
    start_paragraph: int = None,
    end_paragraph: int = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    highlight_color: int = None,
    style_name: str = None,
    paragraph_alignment: str = None,
    page_break_before: bool = None,
    preserve_direct_formatting: bool = False,
    track_changes: bool = False,
) -> str:
    """[Windows only] Format text in an open Word document: font, color, highlight, style, alignment, page breaks.
    Use this tool for any visual/formatting change that does NOT alter the text content itself.

    Two addressing modes (provide one):
    - start/end: Character positions (from word_live_find_text or word_live_get_page_text).
    - start_paragraph/end_paragraph: 1-indexed paragraph range (from word_live_get_text etc.).

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        start_paragraph: First paragraph index (1-indexed). Alternative to start/end.
        end_paragraph: Last paragraph index (1-indexed, defaults to start_paragraph).
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
        page_break_before: Set or clear PageBreakBefore on paragraphs in range (True/False).
        preserve_direct_formatting: When True and style_name is set, saves font/size/bold/italic/
            alignment/spacing before applying the style and restores them after. Useful for changing
            a paragraph's style (e.g., Heading 5 → Normal) without losing its visual formatting.
        track_changes: Track formatting changes as revisions.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Resolve addressing mode
        if start_paragraph is not None:
            if end_paragraph is None:
                end_paragraph = start_paragraph
            total_paras = doc.Paragraphs.Count
            if start_paragraph < 1 or end_paragraph > total_paras:
                return json.dumps({
                    "error": f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
                })
            p_start = doc.Paragraphs(start_paragraph).Range.Start
            p_end = doc.Paragraphs(end_paragraph).Range.End
            rng = doc.Range(p_start, p_end)
            range_label = f"para {start_paragraph}-{end_paragraph}"
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
            range_label = f"{start}-{end}"
        else:
            return json.dumps(
                {"error": "Provide start/end character positions OR start_paragraph/end_paragraph"}
            )

        with undo_record(app, "MCP: Format Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                # Save direct formatting before style change if requested
                saved_formats = []
                if preserve_direct_formatting and style_name is not None:
                    for para in rng.Paragraphs:
                        pr = para.Range
                        pf = para.Format
                        saved_formats.append({
                            "para": para,
                            "font_name": str(pr.Font.Name) if pr.Font.Name and pr.Font.Name != 9999999 else None,
                            "font_size": pr.Font.Size if pr.Font.Size and pr.Font.Size != 9999999 else None,
                            "bold": pr.Font.Bold if pr.Font.Bold != 9999999 else None,
                            "italic": pr.Font.Italic if pr.Font.Italic != 9999999 else None,
                            "alignment": pf.Alignment,
                            "space_before": pf.SpaceBefore,
                            "space_after": pf.SpaceAfter,
                            "line_spacing": pf.LineSpacing,
                            "line_spacing_rule": pf.LineSpacingRule,
                        })

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
                    if preserve_direct_formatting:
                        # Apply style per-paragraph and restore formatting
                        for sf in saved_formats:
                            p = sf["para"]
                            p.Style = doc.Styles(style_name)
                            pr = p.Range
                            pf = p.Format
                            if sf["font_name"] is not None:
                                pr.Font.Name = sf["font_name"]
                            if sf["font_size"] is not None:
                                pr.Font.Size = sf["font_size"]
                            if sf["bold"] is not None:
                                pr.Font.Bold = sf["bold"]
                            if sf["italic"] is not None:
                                pr.Font.Italic = sf["italic"]
                            pf.Alignment = sf["alignment"]
                            pf.SpaceBefore = sf["space_before"]
                            pf.SpaceAfter = sf["space_after"]
                            pf.LineSpacingRule = sf["line_spacing_rule"]
                            pf.LineSpacing = sf["line_spacing"]
                    else:
                        rng.Style = style_name
                if paragraph_alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    al = align_map.get(paragraph_alignment.lower())
                    if al is None:
                        return json.dumps({"error": f"Invalid alignment: {paragraph_alignment}. Use: left, center, right, justify"})
                    for para in rng.Paragraphs:
                        para.Format.Alignment = al
                if page_break_before is not None:
                    for para in rng.Paragraphs:
                        para.Format.PageBreakBefore = page_break_before
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
                "range": range_label,
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
    number_format: dict = None,
    start_at: dict = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Apply or remove bullet/numbered/multilevel list formatting on paragraphs.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph to format (1-indexed, required).
        end_paragraph: Last paragraph to format (1-indexed, defaults to start_paragraph).
        list_type: "bullet", "number", or "multilevel" (outline numbered).
        level: Indentation level (0 = first level, 1 = second level, etc.).
            For multilevel, this sets the list level per paragraph.
        remove: If True, removes list formatting from the range.
        continue_previous: If True, continues numbering from a previous list above.
        number_format: (multilevel only) Dict mapping level (int) to format string.
            Example: {1: "%1.", 2: "%1.%2."} → "5.", "5.1."
            Keys are 1-indexed levels. If not provided, defaults to {1: "%1.", 2: "%1.%2."}.
        start_at: (multilevel only) Dict mapping level (int) to starting number.
            Example: {1: 5} → numbering starts at 5.
            If not provided, starts at 1.
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
                formatted = 0

                if remove:
                    for i in range(start_paragraph, end_paragraph + 1):
                        doc.Paragraphs(i).Range.ListFormat.RemoveNumbers()
                        formatted += 1
                elif list_type == "multilevel":
                    # Create custom multilevel list template (OutlineNumbered gallery)
                    lt = doc.ListTemplates.Add(OutlineNumbered=True)
                    # Normalize dict keys to int (JSON sends string keys)
                    nf = {int(k): v for k, v in (number_format or {1: "%1.", 2: "%1.%2."}).items()}
                    sa = {int(k): v for k, v in (start_at or {}).items()}
                    for lvl_num, fmt_str in nf.items():
                        lv = lt.ListLevels(int(lvl_num))
                        lv.NumberFormat = fmt_str
                        lv.NumberStyle = 0  # wdListNumberStyleArabic
                        lv.StartAt = sa.get(int(lvl_num), 1)
                        lv.Alignment = 0  # left
                        lv.NumberPosition = 0
                        lv.TextPosition = 28
                        lv.TabPosition = 28
                        # Do NOT set LinkedStyle — avoids Heading style side effects

                    for i in range(start_paragraph, end_paragraph + 1):
                        para = doc.Paragraphs(i)
                        should_continue = (i > start_paragraph) or continue_previous
                        para.Range.ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate=lt,
                            ContinuePreviousList=should_continue,
                            DefaultListBehavior=1,
                        )
                        if level > 0:
                            para.Range.ListFormat.ListLevelNumber = level + 1
                        formatted += 1
                else:
                    # bullet or number (original logic)
                    gallery_map = {"bullet": 1, "number": 2}
                    gallery_idx = gallery_map.get(list_type, 1)
                    template = doc.Application.ListGalleries(gallery_idx).ListTemplates(1)
                    for i in range(start_paragraph, end_paragraph + 1):
                        para = doc.Paragraphs(i)
                        should_continue = (i > start_paragraph) or continue_previous
                        para.Range.ListFormat.ApplyListTemplateWithLevel(
                            ListTemplate=template,
                            ContinuePreviousList=should_continue,
                            DefaultListBehavior=1,
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
    h1_number_format: str = None,
    h2_number_format: str = None,
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

    Creates a multilevel list template linked to Heading 1 and Heading 2 styles.
    Default formats: Level 1 = "%1." (produces "1."), Level 2 = "%1.%2" (produces "1.1").
    Custom formats supported — e.g., h1_number_format="MADDE %1 – " produces "MADDE 1 – ".

    Applies styles and numbering to the specified paragraphs, then optionally
    strips manual number prefixes. Recognizes two patterns:
    - Numeric: "1. ", "6.2. ", "10.3 " (regex: ^\d+(\.\d+)*\.?\s+)
    - MADDE: "MADDE 6 – ", "MADDE 10 - " (regex: ^MADDE\s+\d+\s*[–-]\s*)

    If any style parameter is provided, Heading 1 and Heading 2 styles are
    customized before applying. If no style params are given, only numbering
    is applied (existing styles are preserved).

    Args:
        filename: Document name or path (None = active document).
        h1_paragraphs: List of 1-indexed paragraph numbers for Heading 1 (main sections).
        h2_paragraphs: List of 1-indexed paragraph numbers for Heading 2 (sub-sections).
        strip_manual_numbers: Remove leading number/MADDE prefix from headings (default True).
        h1_number_format: Custom Level 1 format (default "%1."). Use %1 for the number.
            Example: "MADDE %1 – " produces "MADDE 1 – ", "MADDE 2 – ", etc.
        h2_number_format: Custom Level 2 format (default "%1.%2"). Use %1 and %2.
            Example: "%1.%2." produces "1.1.", "1.2.", etc.
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

        def _find_para_text(doc, text):
            """Find paragraph text in doc body, return range or None."""
            search = text[:60] if len(text) > 60 else text
            if not search:
                return None
            rng = doc.Content.Duplicate
            rng.Find.ClearFormatting()
            rng.Find.Execute(
                FindText=search, Forward=True,
                MatchCase=True, MatchWholeWord=False, Wrap=0,
            )
            return rng if rng.Find.Found else None

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
                    # H1 keeps with next (heading stays with first body para).
                    # H2 does NOT — sub-clauses are often full paragraphs;
                    # chaining keep_with_next across them breaks page layout.
                    s.ParagraphFormat.KeepWithNext = (style_id == -2)
                    s.ParagraphFormat.KeepTogether = False

            # --- Create multilevel list template ---
            lt = doc.ListTemplates.Add(OutlineNumbered=True)
            h1_fmt = h1_number_format or "%1."
            h2_fmt = h2_number_format or "%1.%2"

            # Level 1 linked to Heading 1
            lv1 = lt.ListLevels(1)
            lv1.NumberFormat = h1_fmt
            lv1.NumberStyle = 0  # wdListNumberStyleArabic
            lv1.StartAt = 1
            lv1.Alignment = 0  # left
            lv1.NumberPosition = 0
            if len(h1_fmt) > 5:
                # Long format (e.g., "MADDE %1 – ") — text follows number directly
                lv1.TextPosition = 0
                lv1.TabPosition = 0
            else:
                lv1.TextPosition = 28  # ~1cm indent for text after number
                lv1.TabPosition = 28
            lv1.LinkedStyle = "Heading 1"

            # Level 2 linked to Heading 2
            lv2 = lt.ListLevels(2)
            lv2.NumberFormat = h2_fmt
            lv2.NumberStyle = 0
            lv2.StartAt = 1
            lv2.Alignment = 0
            lv2.NumberPosition = 0
            if len(h2_fmt) > 5:
                lv2.TextPosition = 0
                lv2.TabPosition = 0
            else:
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

            target_style = {-2: doc.Styles(-2), -3: doc.Styles(-3)}

            for para_idx, style_id in all_heading_paras:
                if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                    continue
                para = doc.Paragraphs(para_idx)
                text = para.Range.Text.rstrip("\r\x07")
                range_len = para.Range.End - para.Range.Start
                text_len = len(text)
                inflated = (range_len > text_len + 5)

                if not inflated:
                    # Normal paragraph — direct style assignment works.
                    try:
                        para.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    para.Style = target_style[style_id]
                else:
                    # Inflated Range (comments/fields extend it beyond text).
                    # Use Find to locate text, then Expand to full paragraph
                    # so the paragraph mark gets the style.
                    found = _find_para_text(doc, text)
                    if found:
                        try:
                            found.ListFormat.RemoveNumbers()
                        except Exception:
                            pass
                        # Expand found range to full paragraph (includes \r mark)
                        found.Expand(Unit=4)  # wdParagraph
                        found.Style = target_style[style_id]

                if style_id == -2:
                    h1_applied += 1
                else:
                    h2_applied += 1

            # --- Apply list template via LinkedStyle propagation ---
            # Apply list to FIRST H1 paragraph only. Because the template
            # has LinkedStyle for Heading 1 and Heading 2, Word auto-applies
            # the correct list level to ALL paragraphs with those styles.
            # This avoids per-paragraph Range issues (inflated Range.End
            # from comments/fields/bookmarks breaks per-paragraph approach).
            if h1_paragraphs:
                first_h1 = doc.Paragraphs(sorted(h1_paragraphs)[0])
                first_h1.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=lt,
                    ContinuePreviousList=False,
                    DefaultListBehavior=1,
                )

            # --- Strip manual numbers ---
            stripped = 0
            if strip_manual_numbers:
                strip_patterns = [
                    r"^MADDE\s+\d+\s*[–\-]\s*",  # "MADDE 6 – ", "MADDE 10 - "
                    r"^\d+(\.\d+)*\.?\s+",         # "1. ", "6.2. ", "10.3 "
                ]
                for para_idx, _ in all_heading_paras:
                    if para_idx < 1 or para_idx > doc.Paragraphs.Count:
                        continue
                    para = doc.Paragraphs(para_idx)
                    text = para.Range.Text.rstrip("\r\x07")
                    for pattern in strip_patterns:
                        m = re.match(pattern, text)
                        if m:
                            prefix_len = len(m.group(0))
                            range_len = para.Range.End - para.Range.Start
                            if range_len <= len(text) + 5:
                                # Normal — para.Range.Start is reliable
                                rng = doc.Range(
                                    para.Range.Start,
                                    para.Range.Start + prefix_len,
                                )
                            else:
                                # Inflated — find text to get real position
                                found = _find_para_text(doc, text)
                                if not found:
                                    break
                                rng = doc.Range(
                                    found.Start,
                                    found.Start + prefix_len,
                                )
                            rng.Delete()
                            stripped += 1
                            break

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "h1_applied": h1_applied,
            "h2_applied": h2_applied,
            "stripped": stripped,
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_replace_text(
    filename: str = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    match_whole_word: bool = False,
    use_wildcards: bool = False,
    replace_all: bool = True,
    track_changes: bool = False,
) -> str:
    """[Windows only] Find and replace text in an open Word document.

    Uses Word's native Find & Replace, which works across tracked change boundaries
    (unlike manual delete+insert). Supports Word special characters when use_wildcards=True:
    ^m (manual page break), ^t (tab), ^p (paragraph mark), and Word wildcard syntax.

    Args:
        filename: Document name or path (None = active document).
        find_text: Text to find. With use_wildcards=True, supports ^m, ^t, ^p and Word wildcards.
        replace_text: Replacement text. Use "" to delete matches.
        match_case: Case-sensitive search.
        match_whole_word: Match whole words only (ignored when use_wildcards=True).
        use_wildcards: Enable Word wildcards and special characters.
        replace_all: Replace all occurrences (True) or just the first one (False).
        track_changes: Track replacements as revisions.

    Returns:
        JSON with count of replacements made.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not find_text:
        return json.dumps({"error": "find_text is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Replace Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                count = 0
                rng = doc.Content.Duplicate
                rng.Find.ClearFormatting()
                rng.Find.Replacement.ClearFormatting()

                # Loop with wdReplaceOne (=1) to count replacements
                while True:
                    found = rng.Find.Execute(
                        FindText=find_text,
                        ReplaceWith=replace_text,
                        MatchCase=match_case,
                        MatchWholeWord=match_whole_word if not use_wildcards else False,
                        MatchWildcards=use_wildcards,
                        Forward=True,
                        Wrap=0,  # wdFindStop
                        Replace=1,  # wdReplaceOne
                    )
                    if not found:
                        break
                    count += 1
                    if not replace_all:
                        break
                    # Reset range to search remaining document
                    rng = doc.Content.Duplicate
                    rng.Find.ClearFormatting()
                    rng.Find.Replacement.ClearFormatting()
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements": count,
            "replace_all": replace_all,
            "tracked": track_changes,
        }, ensure_ascii=False)

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
