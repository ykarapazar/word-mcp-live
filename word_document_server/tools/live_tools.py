"""Live editing tools for Microsoft Word via COM automation.

These tools operate on documents that are currently open in Word,
providing real-time editing capabilities with optional tracked changes.
"""

import json
import os
import re
import sys

from word_document_server.defaults import DEFAULT_AUTHOR

# Word COM constants
WD_STORY = 6

# Word COM InsertBefore/InsertAfter limit (~32K chars).
# We use 30000 as safe margin below 2^15-1 = 32767.
_INSERT_CHUNK_SIZE = 30000


async def word_live_insert_text(
    filename: str = None,
    text: str = "",
    position: str = "end",
    bookmark: str = None,
    track_changes: bool = False,
) -> str:
    """Insert text into an open Word document.

    Automatically chunks large text (>30K chars) to avoid Word COM limits.

    Args:
        filename: Document name or path (None = active document).
        text: Text to insert (no length limit — auto-chunked if needed).
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

        # Convert literal escape sequences to actual characters.
        # MCP/JSON sends backslash-r as 2 chars; Word COM needs chr(13) for paragraph marks.
        text = text.replace("\\r\\n", "\r").replace("\\r", "\r").replace("\\n", "\r")

        with undo_record(app, "MCP: Insert Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                chunks = [text[i:i + _INSERT_CHUNK_SIZE]
                          for i in range(0, max(len(text), 1), _INSERT_CHUNK_SIZE)]

                if bookmark:
                    if not doc.Bookmarks.Exists(bookmark):
                        return json.dumps({"error": f"Bookmark '{bookmark}' not found"})
                    rng = doc.Bookmarks(bookmark).Range
                    for chunk in chunks:
                        rng.InsertAfter(chunk)
                        rng.Collapse(0)  # wdCollapseEnd
                elif position == "start":
                    # InsertBefore: reverse order so first chunk ends up first
                    for chunk in reversed(chunks):
                        doc.Range(0, 0).InsertBefore(chunk)
                elif position == "end":
                    for chunk in chunks:
                        end_pos = doc.Content.End - 1
                        rng = doc.Range(end_pos, end_pos)
                        rng.InsertAfter(chunk)
                elif position == "cursor":
                    for chunk in chunks:
                        app.Selection.TypeText(chunk)
                else:
                    try:
                        offset = int(position)
                    except ValueError:
                        return json.dumps(
                            {
                                "error": f"Invalid position: {position}. "
                                "Use 'start', 'end', 'cursor', or a character offset."
                            }
                        )
                    # InsertBefore at offset: reverse order so first chunk ends up at offset
                    for chunk in reversed(chunks):
                        doc.Range(offset, offset).InsertBefore(chunk)
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        result = {
            "success": True,
            "document": doc.Name,
            "text_length": len(text),
            "position": position,
            "tracked": track_changes,
        }
        if len(chunks) > 1:
            result["chunks_used"] = len(chunks)
        return json.dumps(result)

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
    strikethrough: bool = None,
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
        strikethrough: Set strikethrough (True/False).
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
                            "strikethrough": pr.Font.StrikeThrough if pr.Font.StrikeThrough != 9999999 else None,
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
                if strikethrough is not None:
                    rng.Font.StrikeThrough = strikethrough
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
                            if sf["strikethrough"] is not None:
                                pr.Font.StrikeThrough = sf["strikethrough"]
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
    number_style: dict = None,
    start_at: dict = None,
    level_map: dict = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Apply or remove bullet/numbered/multilevel list formatting on paragraphs.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph to format (1-indexed, required).
        end_paragraph: Last paragraph to format (1-indexed, defaults to start_paragraph).
        list_type: "bullet", "number", or "multilevel" (outline numbered).
        level: Indentation level (0 = first level, 1 = second level, etc.).
            For multilevel, this sets the default list level for all paragraphs.
            Use level_map instead for per-paragraph level control.
        remove: If True, removes list formatting from the range.
        continue_previous: If True, continues numbering from a previous list above.
        number_format: (multilevel only) Dict mapping level (int) to format string.
            Example: {1: "4.%1.", 2: "(%2)", 3: "(%3)"} → "4.1.", "(a)", "(i)"
            Keys are 1-indexed levels. If not provided, defaults to {1: "%1.", 2: "%1.%2."}.
        number_style: (multilevel only) Dict mapping level (int) to numbering style string.
            Styles: "arabic" (1,2,3), "lowercase_letter" (a,b,c), "uppercase_letter" (A,B,C),
            "lowercase_roman" (i,ii,iii), "uppercase_roman" (I,II,III).
            Example: {1: "arabic", 2: "lowercase_letter", 3: "lowercase_roman"}
            If a string is given instead of dict, applies same style to all levels.
            Default: "arabic" for all levels.
        start_at: (multilevel only) Dict mapping level (int) to starting number.
            Example: {1: 5} → numbering starts at 5.
            If not provided, starts at 1.
        level_map: (multilevel only) Dict mapping paragraph index (int) to list level (int, 1-indexed).
            Example: {29: 2, 30: 2, 37: 3} → para 29 at level 2, para 37 at level 3.
            Paragraphs not in the map stay at level 1 (or the value of `level + 1`).
            Applied AFTER the list template, so the template covers the full range.
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
                    lm = {int(k): int(v) for k, v in (level_map or {}).items()}
                    # Map number_style string to wdListNumberStyle constant
                    style_map = {
                        "arabic": 0, "lowercase_letter": 4, "uppercase_letter": 3,
                        "lowercase_roman": 2, "uppercase_roman": 1,
                    }
                    # number_style can be a string (same for all) or dict (per-level)
                    if isinstance(number_style, dict):
                        ns_map = {int(k): style_map.get(v, 0) for k, v in number_style.items()}
                    elif isinstance(number_style, str):
                        ns_map = {lvl: style_map.get(number_style, 0) for lvl in nf}
                    else:
                        ns_map = {}
                    for lvl_num, fmt_str in nf.items():
                        lv = lt.ListLevels(int(lvl_num))
                        lv.NumberFormat = fmt_str
                        lv.NumberStyle = ns_map.get(int(lvl_num), 0)
                        lv.StartAt = sa.get(int(lvl_num), 1)
                        lv.Alignment = 0  # left
                        lv.NumberPosition = 0
                        lv.TextPosition = 28
                        lv.TabPosition = 28
                        # Do NOT set LinkedStyle — avoids Heading style side effects

                    # Apply template to the full range at once (not per-paragraph)
                    rng = doc.Range(
                        doc.Paragraphs(start_paragraph).Range.Start,
                        doc.Paragraphs(end_paragraph).Range.End,
                    )
                    rng.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate=lt,
                        ContinuePreviousList=continue_previous,
                        ApplyTo=2,  # wdListApplyToSelection
                        DefaultListBehavior=0,
                    )
                    formatted = end_paragraph - start_paragraph + 1

                    # Set per-paragraph levels from level_map
                    default_lvl = level + 1 if level > 0 else 1
                    for i in range(start_paragraph, end_paragraph + 1):
                        target_lvl = lm.get(i, default_lvl)
                        if target_lvl != 1:
                            doc.Paragraphs(i).Range.ListFormat.ListLevelNumber = target_lvl
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

    if len(find_text) > 255:
        return json.dumps({
            "error": f"find_text is {len(find_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        })
    if len(replace_text) > 255:
        return json.dumps({
            "error": f"replace_text is {len(replace_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        })

    if replace_all and track_changes:
        return json.dumps({
            "error": "replace_all=True with track_changes=True causes an infinite loop "
            "(tracked deletions stay visible to Find, triggering endless re-replacement). "
            "Use replace_all=False — each unique text only needs one replacement."
        })

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
            elif replace_all and prev_tracking:
                # Issue #7: document has TrackRevisions on but caller wants
                # untracked replace_all — disable temporarily to prevent
                # infinite loop (tracked deletions stay visible to Find).
                doc.TrackRevisions = False

            try:
                count = 0
                rng = doc.Content.Duplicate
                rng.Find.ClearFormatting()

                while True:
                    found = rng.Find.Execute(
                        FindText=find_text,
                        MatchCase=match_case,
                        MatchWholeWord=match_whole_word if not use_wildcards else False,
                        MatchWildcards=use_wildcards,
                        Forward=True,
                        Wrap=0,  # wdFindStop
                    )
                    if not found:
                        break
                    # Convert Word special characters to actual characters for rng.Text assignment
                    # (rng.Text doesn't interpret ^p/^t/^m like Find.Execute Replace does)
                    processed = replace_text.replace("^p", "\r").replace("^t", "\t").replace("^m", "\x0c")
                    rng.Text = processed
                    count += 1
                    if not replace_all:
                        break
                    rng.Collapse(0)  # wdCollapseEnd — move past replacement
            finally:
                doc.TrackRevisions = prev_tracking
                if track_changes:
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


async def word_live_insert_paragraphs(
    filename: str = None,
    paragraphs: list = None,
    target_text: str = None,
    target_paragraph_index: int = None,
    position: str = "after",
    style: str = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Insert one or more paragraphs near a target paragraph in an open Word document.

    Targets by text match or paragraph index (0-based, matching word_live_get_text output).
    Inserts all paragraphs in a single undo record.

    Args:
        filename: Document name or path (None = active document).
        paragraphs: List of paragraph texts to insert. Each string becomes one Word paragraph.
        target_text: Text to search for (first matching paragraph). Mutually exclusive with target_paragraph_index.
        target_paragraph_index: 0-based paragraph index (as returned by word_live_get_text).
        position: 'before' or 'after' the target paragraph (default 'after').
        style: Style name for inserted paragraphs. None = "Normal" (avoids inheriting heading styles).
        track_changes: Track insertions as revisions.

    Returns:
        JSON with result info including count of paragraphs inserted.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not paragraphs or not isinstance(paragraphs, list):
        return json.dumps({"error": "paragraphs must be a non-empty list of strings"})

    if target_text is None and target_paragraph_index is None:
        return json.dumps({"error": "Provide either target_text or target_paragraph_index"})

    if target_text is not None and target_paragraph_index is not None:
        return json.dumps({"error": "Provide target_text or target_paragraph_index, not both"})

    if position not in ("before", "after"):
        return json.dumps({"error": f"position must be 'before' or 'after', got '{position}'"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Find the target paragraph
        total_paras = doc.Paragraphs.Count
        target_para = None

        if target_paragraph_index is not None:
            com_index = target_paragraph_index + 1  # 0-based API → 1-based COM
            if com_index < 1 or com_index > total_paras:
                return json.dumps({
                    "error": f"target_paragraph_index {target_paragraph_index} out of range "
                    f"(0-{total_paras - 1})"
                })
            target_para = doc.Paragraphs(com_index)
        else:
            for i in range(1, total_paras + 1):
                para = doc.Paragraphs(i)
                para_text = para.Range.Text.rstrip("\r\x07")
                if target_text in para_text:
                    target_para = para
                    break
            if target_para is None:
                return json.dumps({"error": f"No paragraph found containing '{target_text}'"})

        resolved_style = style if style else "Normal"

        with undo_record(app, "MCP: Insert Paragraphs"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                inserted = 0

                if position == "after":
                    rng = target_para.Range.Duplicate
                    rng.Collapse(0)  # wdCollapseEnd
                    for para_text in paragraphs:
                        rng.InsertParagraphAfter()
                        rng.Collapse(0)  # wdCollapseEnd
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        rng.Collapse(0)  # wdCollapseEnd
                        inserted += 1
                else:  # "before"
                    for para_text in reversed(paragraphs):
                        rng = target_para.Range.Duplicate
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertParagraphBefore()
                        rng.Collapse(1)  # wdCollapseStart
                        rng.InsertAfter(para_text)
                        try:
                            rng.Style = resolved_style
                        except Exception:
                            pass
                        inserted += 1
            finally:
                doc.TrackRevisions = prev_tracking
                if track_changes:
                    app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_inserted": inserted,
            "position": position,
            "style": resolved_style,
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
    style: str = "Table Grid",
    autofit: str = "window",
    track_changes: bool = False,
) -> str:
    """Add a table to an open Word document.

    Args:
        filename: Document name or path.
        rows: Number of rows.
        cols: Number of columns.
        position: "start", "end", or character offset.
        data: Optional 2D list of cell data.
        style: Table style name. Default "Table Grid" (bordered).
            Use None or "" for no style.
        autofit: "window" (fit page width, default), "content" (fit cell content),
            "fixed" (fixed widths), or None for legacy behavior (no autofit).
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
                # AutoFit behavior constants
                AUTOFIT_MAP = {
                    "window": (1, 2),   # wdWord9TableBehavior, wdAutoFitWindow
                    "content": (1, 1),  # wdWord9TableBehavior, wdAutoFitContent
                    "fixed": (0, 0),    # wdWord8TableBehavior, wdAutoFitFixed
                }

                if autofit and autofit.lower() in AUTOFIT_MAP:
                    default_behavior, autofit_behavior = AUTOFIT_MAP[autofit.lower()]
                    table = doc.Tables.Add(rng, rows, cols, default_behavior, autofit_behavior)
                else:
                    table = doc.Tables.Add(rng, rows, cols)

                # Apply table style
                if style:
                    try:
                        table.Style = doc.Styles(style)
                    except Exception:
                        pass  # Style not found; proceed without

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
                "style": style or None,
                "autofit": autofit or None,
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_format_table(
    filename: str = None,
    table_index: int = -1,
    border_style: str = None,
    cell_bold: list = None,
    cell_alignment: list = None,
    column_widths: list = None,
    table_alignment: str = None,
    cell_shading: list = None,
    autofit: str = None,
) -> str:
    """Format a table in an open Word document via COM.

    Supports border removal, cell formatting, column sizing, and table alignment.
    Use table_index=-1 for the last table, 1 for the first, etc.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index, or -1 for the last table.
        border_style: Border style for all edges: "none", "single", "double", "dotted",
            "dashed", "thick". "none" removes all borders.
        cell_bold: List of [row, col, bold] entries (1-indexed) to set bold on cell text.
            Example: [[1, 1, true], [1, 2, true]] bolds row 1 cells.
        cell_alignment: List of [row, col, alignment] entries. alignment: "left", "center",
            "right", "justify". Row 0 = all rows, Col 0 = all cols.
        column_widths: List of column widths in points (1-indexed order).
            Example: [200, 200] sets col 1 to 200pt, col 2 to 200pt.
        table_alignment: Table alignment on page: "left", "center", "right".
        cell_shading: List of [row, col, color_hex] entries. color_hex as "#RRGGBB".
            Row 0 = all rows. Example: [[1, 0, "#DDDDDD"]] shades entire row 1.
        autofit: "window" (fit to page width), "content" (fit to cell content),
            "fixed" (fixed column widths).

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if doc.Tables.Count == 0:
            return json.dumps({"error": "Document has no tables"})

        idx = table_index if table_index > 0 else doc.Tables.Count
        if idx < 1 or idx > doc.Tables.Count:
            return json.dumps({"error": f"Table index {table_index} out of range (1-{doc.Tables.Count})"})

        tbl = doc.Tables(idx)
        actions = []

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap (thick)
        }

        BORDER_IDS = [-1, -2, -3, -4, -5, -6, -7, -8]  # top, left, bottom, right, horiz, vert, etc.

        with undo_record(app, "MCP: Format Table"):
            # --- Borders ---
            if border_style is not None:
                style_val = BORDER_STYLES.get(border_style.lower())
                if style_val is None:
                    return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                for bid in BORDER_IDS:
                    try:
                        tbl.Borders(bid).LineStyle = style_val
                    except Exception:
                        pass
                actions.append(f"borders={border_style}")

            # --- Autofit ---
            if autofit is not None:
                AUTOFIT = {"window": 2, "content": 1, "fixed": 0}  # wdAutoFitWindow=2, wdAutoFitContent=1, wdAutoFitFixed=0
                af_val = AUTOFIT.get(autofit.lower())
                if af_val is not None:
                    tbl.AutoFitBehavior(af_val)
                    actions.append(f"autofit={autofit}")

            # --- Table alignment ---
            if table_alignment is not None:
                ALIGN = {"left": 0, "center": 1, "right": 2}
                al_val = ALIGN.get(table_alignment.lower())
                if al_val is not None:
                    tbl.Rows.Alignment = al_val
                    actions.append(f"table_alignment={table_alignment}")

            # --- Column widths ---
            if column_widths is not None:
                for ci, width in enumerate(column_widths):
                    if ci < tbl.Columns.Count:
                        tbl.Columns(ci + 1).Width = float(width)
                actions.append(f"column_widths={column_widths}")

            # --- Cell bold ---
            if cell_bold is not None:
                for entry in cell_bold:
                    r, c, bold_val = int(entry[0]), int(entry[1]), bool(entry[2])
                    if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                        tbl.Cell(r, c).Range.Font.Bold = bold_val
                actions.append(f"cell_bold={len(cell_bold)} cells")

            # --- Cell alignment ---
            if cell_alignment is not None:
                PARA_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
                for entry in cell_alignment:
                    r, c, align = int(entry[0]), int(entry[1]), str(entry[2]).lower()
                    al = PARA_ALIGN.get(align, 0)
                    if r == 0 and c == 0:
                        # All cells
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                tbl.Cell(ri, ci).Range.ParagraphFormat.Alignment = al
                    elif r == 0:
                        # Entire column
                        for ri in range(1, tbl.Rows.Count + 1):
                            tbl.Cell(ri, c).Range.ParagraphFormat.Alignment = al
                    elif c == 0:
                        # Entire row
                        for ci in range(1, tbl.Columns.Count + 1):
                            tbl.Cell(r, ci).Range.ParagraphFormat.Alignment = al
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            tbl.Cell(r, c).Range.ParagraphFormat.Alignment = al
                actions.append(f"cell_alignment={len(cell_alignment)} entries")

            # --- Cell shading ---
            if cell_shading is not None:
                for entry in cell_shading:
                    r, c, color_hex = int(entry[0]), int(entry[1]), str(entry[2])
                    # Convert #RRGGBB to Word BGR integer
                    color_hex = color_hex.lstrip("#")
                    rr, gg, bb = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
                    bgr = bb * 65536 + gg * 256 + rr

                    def shade_cell(row_i, col_i):
                        tbl.Cell(row_i, col_i).Shading.BackgroundPatternColor = bgr

                    if r == 0 and c == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            for ci in range(1, tbl.Columns.Count + 1):
                                shade_cell(ri, ci)
                    elif r == 0:
                        for ri in range(1, tbl.Rows.Count + 1):
                            shade_cell(ri, c)
                    elif c == 0:
                        for ci in range(1, tbl.Columns.Count + 1):
                            shade_cell(r, ci)
                    else:
                        if 1 <= r <= tbl.Rows.Count and 1 <= c <= tbl.Columns.Count:
                            shade_cell(r, c)
                actions.append(f"cell_shading={len(cell_shading)} entries")

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "table_index": idx,
                "rows": tbl.Rows.Count,
                "cols": tbl.Columns.Count,
                "actions": actions,
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
                # Delete any table objects within the range first
                # (rng.Delete only removes text, leaving ghost table structure)
                for i in range(doc.Tables.Count, 0, -1):
                    tbl = doc.Tables(i)
                    if tbl.Range.Start >= start and tbl.Range.End <= end:
                        tbl.Delete()
                # Delete remaining text in the range
                rng = doc.Range(start, min(end, doc.Content.End))
                if rng.Start < rng.End:
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


async def word_live_modify_table(
    filename: str = None,
    table_index: int = 1,
    operation: str = "get_info",
    row: int = None,
    col: int = None,
    text: str = None,
    before_row: int = None,
    before_col: int = None,
    header: str = None,
    cells: list = None,
    start_row: int = None,
    start_col: int = None,
    end_row: int = None,
    end_col: int = None,
    autofit_mode: str = "content",
    accept_revisions: bool = False,
    track_changes: bool = False,
) -> str:
    """[Windows only] Modify a table in an open Word document.

    Operations: get_info, set_cell, set_row, set_range, add_column, delete_column,
    add_row, delete_row, merge_cells, autofit, delete_table.
    All row/col indices are 1-based (Word COM standard).

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index (default 1).
        operation: One of: get_info, set_cell, set_row, set_range, add_column,
            delete_column, add_row, delete_row, merge_cells, autofit, delete_table.
        row: Row index for set_cell, set_row, delete_row.
        col: Column index for set_cell, delete_column.
        text: Text for set_cell.
        before_row: Insert row before this index (add_row). None = append at end.
        before_col: Insert column before this index (add_column). None = append at end.
        header: Header text for new column (add_column, placed in row 1).
        cells: List of cell values for set_row (1D) or set_range (2D). None values skip that cell.
            Also used for new row/column values (add_row, add_column).
        start_row: Start row for merge_cells or set_range (default 1).
        start_col: Start column for merge_cells or set_range (default 1).
        end_row: End row for merge_cells.
        end_col: End column for merge_cells.
        autofit_mode: 'content', 'window', or 'fixed' (autofit operation).
        accept_revisions: For set_cell/set_row/set_range — accept tracked changes before writing
            (prevents layered text from old revisions persisting underneath new content).
        track_changes: Track modifications as revisions.

    Returns:
        JSON with operation result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        from word_document_server.core import table_com

        app = get_word_app()
        doc = find_document(app, filename)

        if doc.Tables.Count == 0:
            return json.dumps({"error": "Document has no tables"})

        if table_index < 1 or table_index > doc.Tables.Count:
            return json.dumps({"error": f"table_index {table_index} out of range (1-{doc.Tables.Count})"})

        table = doc.Tables(table_index)
        op = operation.lower()

        # get_info is read-only — no undo record needed
        if op == "get_info":
            result = table_com.get_info(table)
            result["document"] = doc.Name
            result["table_index"] = table_index
            return json.dumps(result, ensure_ascii=False)

        # All other operations are destructive
        with undo_record(app, "MCP: Modify Table"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                if op == "set_cell":
                    if row is None or col is None or text is None:
                        return json.dumps({"error": "set_cell requires row, col, and text"})
                    result = table_com.set_cell(table, row, col, text, accept_revisions=accept_revisions)

                elif op == "set_row":
                    if row is None or not cells:
                        return json.dumps({"error": "set_row requires row and cells (list of values)"})
                    result = table_com.set_row(table, row, cells, accept_revisions=accept_revisions)

                elif op == "set_range":
                    if not cells:
                        return json.dumps({"error": "set_range requires cells (2D list of values)"})
                    result = table_com.set_range(
                        table, cells,
                        start_row=start_row or 1,
                        start_col=start_col or 1,
                        accept_revisions=accept_revisions,
                    )

                elif op == "add_column":
                    result = table_com.add_column(table, before_col, header, cells)

                elif op == "delete_column":
                    if col is None:
                        return json.dumps({"error": "delete_column requires col"})
                    result = table_com.delete_column(table, col)

                elif op == "add_row":
                    result = table_com.add_row(table, before_row, cells)

                elif op == "delete_row":
                    if row is None:
                        return json.dumps({"error": "delete_row requires row"})
                    result = table_com.delete_row(table, row)

                elif op == "merge_cells":
                    if not all(v is not None for v in [start_row, start_col, end_row, end_col]):
                        return json.dumps({"error": "merge_cells requires start_row, start_col, end_row, end_col"})
                    result = table_com.merge_cells(table, start_row, start_col, end_row, end_col)

                elif op == "autofit":
                    result = table_com.autofit(table, autofit_mode)

                elif op == "delete_table":
                    result = table_com.delete_table(table)

                else:
                    return json.dumps({
                        "error": f"Unknown operation '{op}'. Use: get_info, set_cell, set_row, set_range, "
                        "add_column, delete_column, add_row, delete_row, merge_cells, autofit, delete_table"
                    })
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        result["success"] = True
        result["document"] = doc.Name
        result["table_index"] = table_index
        result["operation"] = op
        result["tracked"] = track_changes
        return json.dumps(result, ensure_ascii=False)

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


async def word_live_save(
    filename: str = None,
    save_as: str = None,
) -> str:
    """Save an open Word document.

    Saves the document. Optionally saves to a new path with save_as.

    Args:
        filename: Document name or path (None = active document).
        save_as: Optional new file path to save as. If omitted, saves in place.

    Returns:
        JSON with save result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if save_as:
            save_path = os.path.abspath(save_as)
            # Determine format from extension
            ext = os.path.splitext(save_path)[1].lower()
            format_map = {
                ".docx": 16,  # wdFormatXMLDocument
                ".doc": 0,    # wdFormatDocument
                ".pdf": 17,   # wdFormatPDF
                ".rtf": 6,    # wdFormatRTF
                ".txt": 2,    # wdFormatText
            }
            file_format = format_map.get(ext, 16)
            doc.SaveAs2(save_path, FileFormat=file_format)
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "saved_as": save_path,
                "format": ext,
            }, ensure_ascii=False)
        else:
            doc.Save()
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "path": doc.FullName,
            }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_toggle_track_changes(
    filename: str = None,
    enable: bool = None,
) -> str:
    """Toggle or set track changes mode on an open Word document.

    If enable is omitted, toggles the current state.

    Args:
        filename: Document name or path (None = active document).
        enable: True to enable, False to disable, None to toggle.

    Returns:
        JSON with the new track changes state.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        previous = bool(doc.TrackRevisions)
        if enable is None:
            doc.TrackRevisions = not previous
        else:
            doc.TrackRevisions = enable

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "previous_state": previous,
            "track_changes": bool(doc.TrackRevisions),
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_image(
    filename: str = None,
    image_path: str = "",
    paragraph_index: int = None,
    position: str = "end",
    width_inches: float = None,
    height_inches: float = None,
    width_pt: float = None,
    height_pt: float = None,
    alignment: str = None,
    wrapping: str = None,
    border_style: str = None,
    border_width_pt: float = None,
    border_color: str = None,
    link_to_file: bool = False,
) -> str:
    """Insert an image into an open Word document.

    The image can be placed at a specific paragraph, at the start or end,
    or at a character offset position.

    Args:
        filename: Document name or path (None = active document).
        image_path: Full path to the image file (PNG, JPG, BMP, etc.).
        paragraph_index: 1-indexed paragraph to insert before (image goes before the paragraph).
        position: "start", "end", or character offset as string. Only used if paragraph_index is None.
        width_inches: Optional width in inches (aspect ratio maintained if only one dimension given).
        height_inches: Optional height in inches.
        width_pt: Optional width in points (1 inch = 72 pt). Overrides width_inches if both given.
        height_pt: Optional height in points. Overrides height_inches if both given.
        alignment: Paragraph alignment for the image: "left", "center", "right". Default: unchanged.
        wrapping: Text wrapping style: "inline" (default), "square", "tight", "behind",
            "infront", "topbottom". Non-inline converts to a floating Shape.
        border_style: Border style around the image: "single", "double", "dotted", "dashed",
            "thick", "none". Default: no border.
        border_width_pt: Border line width in points (e.g. 1.0, 2.0). Default: 1.0.
        border_color: Border color as "#RRGGBB" hex string. Default: black (#000000).
        link_to_file: If True, links to the file instead of embedding it.

    Returns:
        JSON with image insertion result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not image_path:
        return json.dumps({"error": "image_path is required"})

    abs_path = os.path.abspath(image_path)
    if not os.path.isfile(abs_path):
        return json.dumps({"error": f"Image file not found: {abs_path}"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine insertion range
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif position == "start":
            rng = doc.Range(0, 0)
        elif position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        # Resolve final size in points (pt params override inches params)
        final_w = None
        final_h = None
        if width_pt is not None:
            final_w = float(width_pt)
        elif width_inches is not None:
            final_w = float(width_inches) * 72.0
        if height_pt is not None:
            final_h = float(height_pt)
        elif height_inches is not None:
            final_h = float(height_inches) * 72.0

        # Wrapping style constants (wdWrapType)
        WRAP_STYLES = {
            "inline": None,       # keep as InlineShape
            "square": 0,          # wdWrapSquare
            "tight": 1,           # wdWrapTight
            "behind": 3,          # wdWrapBehind
            "infront": 4,         # wdWrapFront
            "topbottom": 2,       # wdWrapTopBottom
        }
        wrap_val = None
        if wrapping is not None:
            wrap_val = WRAP_STYLES.get(wrapping.lower())
            if wrapping.lower() != "inline" and wrap_val is None:
                return json.dumps({"error": f"Unknown wrapping: {wrapping}. Use: {list(WRAP_STYLES.keys())}"})

        # Border style constants
        BORDER_STYLES = {
            "none": 0,     # wdLineStyleNone
            "single": 1,   # wdLineStyleSingle
            "double": 7,   # wdLineStyleDouble
            "dotted": 3,   # wdLineStyleDot
            "dashed": 2,   # wdLineStyleDash
            "thick": 6,    # wdLineStyleThickThinSmallGap
        }

        # Alignment map
        ALIGN_MAP = {"left": 0, "center": 1, "right": 2}

        with undo_record(app, "MCP: Insert Image"):
            inline_shape = rng.InlineShapes.AddPicture(
                FileName=abs_path,
                LinkToFile=link_to_file,
                SaveWithDocument=not link_to_file,
            )

            # Resize if requested (preserves aspect ratio if only one dimension given)
            if final_w is not None and final_h is not None:
                inline_shape.Width = final_w
                inline_shape.Height = final_h
            elif final_w is not None:
                original_ratio = inline_shape.Height / inline_shape.Width
                inline_shape.Width = final_w
                inline_shape.Height = final_w * original_ratio
            elif final_h is not None:
                original_ratio = inline_shape.Width / inline_shape.Height
                inline_shape.Height = final_h
                inline_shape.Width = final_h * original_ratio

            result_width = inline_shape.Width
            result_height = inline_shape.Height
            result_wrapping = "inline"

            # Convert to floating Shape for non-inline wrapping
            if wrap_val is not None:
                float_shape = inline_shape.ConvertToShape()
                float_shape.WrapFormat.Type = wrap_val
                result_wrapping = wrapping.lower()
                result_width = float_shape.Width
                result_height = float_shape.Height

                # Apply border to floating shape
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    # Parse border color
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr  # Word BGR
                    line = float_shape.Line
                    if b_style == 0:  # none
                        line.Visible = False
                    else:
                        line.Visible = True
                        DASH_MAP = {"single": 1, "double": 1, "dotted": 3, "dashed": 4, "thick": 1}
                        line.DashStyle = DASH_MAP.get(border_style.lower(), 1)
                        line.Weight = b_width
                        line.ForeColor.RGB = b_color
                        if border_style.lower() == "double":
                            line.Style = 3  # msoLineThinThin

                # Apply alignment for floating shape using relative positioning
                if alignment is not None:
                    al = alignment.lower()
                    if al in ALIGN_MAP:
                        # Use margin-relative positioning
                        float_shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
                        float_shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph
                        page_setup = doc.PageSetup
                        text_width = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
                        if al == "left":
                            float_shape.Left = 0
                        elif al == "right":
                            float_shape.Left = max(0, text_width - float_shape.Width)
                        else:  # center
                            float_shape.Left = max(0, (text_width - float_shape.Width) / 2)
            else:
                # Inline shape: apply border via inline shape borders
                if border_style is not None:
                    b_style = BORDER_STYLES.get(border_style.lower())
                    if b_style is None:
                        return json.dumps({"error": f"Unknown border_style: {border_style}. Use: {list(BORDER_STYLES.keys())}"})
                    b_width = float(border_width_pt) if border_width_pt else 1.0
                    b_color = 0  # black
                    if border_color:
                        bc = border_color.lstrip("#")
                        rr, gg, bb = int(bc[0:2], 16), int(bc[2:4], 16), int(bc[4:6], 16)
                        b_color = bb * 65536 + gg * 256 + rr
                    # Apply to all 4 borders of inline shape
                    for bid in [-1, -2, -3, -4]:  # top, left, bottom, right
                        try:
                            border = inline_shape.Borders(bid)
                            border.LineStyle = b_style
                            if b_style != 0:
                                border.LineWidth = b_width
                                border.Color = b_color
                        except Exception:
                            pass

                # Apply alignment for inline shape (set paragraph alignment)
                if alignment is not None:
                    al = ALIGN_MAP.get(alignment.lower())
                    if al is not None:
                        inline_shape.Range.ParagraphFormat.Alignment = al

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image": os.path.basename(abs_path),
            "width_pt": result_width,
            "height_pt": result_height,
            "alignment": alignment or "unchanged",
            "wrapping": result_wrapping,
            "border": border_style or "none",
            "linked": link_to_file,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_cross_reference(
    filename: str = None,
    ref_type: str = "heading",
    ref_item: int = 1,
    ref_kind: str = "text",
    insert_position: str = "end",
    paragraph_index: int = None,
    insert_as_hyperlink: bool = True,
) -> str:
    """Insert a cross-reference to a heading, bookmark, figure, or table.

    Cross-references are live fields that update automatically (e.g., "see Section 2.1").

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type of item to reference: "heading", "bookmark", "figure",
                  "table", "equation", "footnote", "endnote".
        ref_item: 1-indexed item number within that reference type.
        ref_kind: What to display: "text" (full text), "number" (label+number),
                  "number_no_context" (just number), "page" (page number),
                  "above_below" ("above" or "below").
        insert_position: "start", "end", or character offset. Used if paragraph_index is None.
        paragraph_index: Insert at the start of this 1-indexed paragraph.
        insert_as_hyperlink: If True, the reference is a clickable hyperlink.

    Returns:
        JSON with cross-reference result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    # Map ref_type to Word constants (wdRefType)
    ref_type_map = {
        "heading": 1,        # wdRefTypeHeading
        "bookmark": 2,       # wdRefTypeBookmark
        "footnote": 3,       # wdRefTypeFootnote
        "endnote": 4,        # wdRefTypeEndnote
        "figure": 10,        # wdRefTypeFigure (SEQ Figure)
        "table": 11,         # wdRefTypeTable (SEQ Table)
        "equation": 12,      # wdRefTypeEquation
    }

    # Map ref_kind to Word constants (wdReferenceKind)
    ref_kind_map = {
        "text": 0,                 # wdContentText
        "number": 1,               # wdNumberFullContext
        "number_no_context": 2,    # wdNumberNoContext
        "number_relative": 3,      # wdNumberRelativeContext
        "page": 7,                 # wdPageNumber
        "above_below": 6,          # wdAboveBelow
    }

    ref_type_lower = ref_type.lower()
    if ref_type_lower not in ref_type_map:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(ref_type_map.keys())}"
        })

    ref_kind_lower = ref_kind.lower()
    if ref_kind_lower not in ref_kind_map:
        return json.dumps({
            "error": f"Invalid ref_kind '{ref_kind}'. Use: {', '.join(ref_kind_map.keys())}"
        })

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Move selection to insertion point
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(1)  # wdCollapseStart
        elif insert_position == "start":
            rng = doc.Range(0, 0)
        elif insert_position == "end":
            rng = doc.Range()
            rng.Collapse(0)  # wdCollapseEnd
        else:
            try:
                offset = int(insert_position)
                rng = doc.Range(offset, offset)
            except (ValueError, TypeError):
                rng = doc.Range()
                rng.Collapse(0)

        rng.Select()

        with undo_record(app, "MCP: Insert Cross Reference"):
            app.Selection.InsertCrossReference(
                ReferenceType=ref_type_map[ref_type_lower],
                ReferenceKind=ref_kind_map[ref_kind_lower],
                ReferenceItem=ref_item,
                InsertAsHyperlink=insert_as_hyperlink,
            )

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "ref_item": ref_item,
            "ref_kind": ref_kind,
            "as_hyperlink": insert_as_hyperlink,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_cross_reference_items(
    filename: str = None,
    ref_type: str = "heading",
) -> str:
    """List all available cross-reference targets of a given type.

    Use this to discover which headings, bookmarks, figures, etc. can be
    referenced, and their 1-based index for use with word_live_insert_cross_reference.

    Args:
        filename: Document name or path (None = active document).
        ref_type: Type to list: "heading", "bookmark", "figure", "table", "equation",
                  "footnote", "endnote".

    Returns:
        JSON with list of referenceable items and their indices.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    valid_types = {"heading", "bookmark", "footnote", "endnote", "figure", "table", "equation"}
    ref_type_lower = ref_type.lower()
    if ref_type_lower not in valid_types:
        return json.dumps({
            "error": f"Invalid ref_type '{ref_type}'. Use: {', '.join(sorted(valid_types))}"
        })

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        result = []

        if ref_type_lower == "heading":
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                style_name = p.Style.NameLocal
                if style_name.startswith("Heading"):
                    text = p.Range.Text.strip()
                    if text:
                        result.append({
                            "index": idx,
                            "text": text,
                            "style": style_name,
                            "paragraph": i,
                        })
                        idx += 1

        elif ref_type_lower == "bookmark":
            for i in range(1, doc.Bookmarks.Count + 1):
                bm = doc.Bookmarks(i)
                text = bm.Range.Text.strip()[:100] if bm.Range else ""
                result.append({
                    "index": i,
                    "name": bm.Name,
                    "text": text,
                })

        elif ref_type_lower == "footnote":
            for i in range(1, doc.Footnotes.Count + 1):
                fn = doc.Footnotes(i)
                text = fn.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower == "endnote":
            for i in range(1, doc.Endnotes.Count + 1):
                en = doc.Endnotes(i)
                text = en.Range.Text.strip()[:100]
                result.append({
                    "index": i,
                    "text": text,
                })

        elif ref_type_lower in ("figure", "table", "equation"):
            # Scan for captioned items (SEQ fields)
            seq_label = {"figure": "Figure", "table": "Table", "equation": "Equation"}[ref_type_lower]
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                text = p.Range.Text.strip()
                if text.startswith(seq_label):
                    result.append({
                        "index": idx,
                        "text": text[:100],
                        "paragraph": i,
                    })
                    idx += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "ref_type": ref_type,
            "items": result,
            "count": len(result),
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_insert_equation(
    filename: str = None,
    equation: str = "",
    paragraph_index: int = None,
    position: str = "end",
    display_mode: bool = False,
) -> str:
    """Insert a mathematical equation into a Word document using UnicodeMath syntax.

    LaTeX-like commands (e.g. \\int, \\sum, \\alpha) are automatically converted to
    Unicode math symbols before insertion, ensuring proper rendering.

    Args:
        filename: Document name (uses active document if None).
        equation: Equation text in UnicodeMath syntax. Examples:
            Simple: "x^2 + y^2 = z^2", "E = mc^2"
            Fractions: "(a+b)/(c+d)"
            Square root: "\\sqrt(x^2+y^2)"
            Greek letters: "\\alpha + \\beta = \\gamma"
            Integrals: "\\int_0^\\infty e^(-x^2) dx"
            Summation: "\\sum_(i=1)^n i^2"
            Matrix: "\\matrix(a&b@c&d)"
            Taylor series: "f(x) = \\sum_(n=0)^\\infty (f^((n))(a))/(n!) (x-a)^n"
        paragraph_index: Insert after this paragraph (1-based). None = use position.
        position: "start" or "end" of document. Ignored if paragraph_index given.
        display_mode: If True, equation is centered on its own line (display style).
            If False, equation is inline with surrounding text.

    Returns:
        JSON with success status and equation details.
    """
    # LaTeX-like command to Unicode math symbol mapping.
    # Word's COM OMaths.Add + BuildUp doesn't process autocorrect entries,
    # so we must pre-convert commands like \int, \sum to their Unicode equivalents.
    UNICODE_MATH = {
        # Greek lowercase
        r"\alpha": "\u03B1", r"\beta": "\u03B2", r"\gamma": "\u03B3",
        r"\delta": "\u03B4", r"\epsilon": "\u03B5", r"\varepsilon": "\u03B5",
        r"\zeta": "\u03B6", r"\eta": "\u03B7", r"\theta": "\u03B8",
        r"\vartheta": "\u03D1", r"\iota": "\u03B9", r"\kappa": "\u03BA",
        r"\lambda": "\u03BB", r"\mu": "\u03BC", r"\nu": "\u03BD",
        r"\xi": "\u03BE", r"\pi": "\u03C0", r"\rho": "\u03C1",
        r"\sigma": "\u03C3", r"\varsigma": "\u03C2", r"\tau": "\u03C4",
        r"\upsilon": "\u03C5", r"\phi": "\u03C6", r"\varphi": "\u03D5",
        r"\chi": "\u03C7", r"\psi": "\u03C8", r"\omega": "\u03C9",
        # Greek uppercase
        r"\Gamma": "\u0393", r"\Delta": "\u0394", r"\Theta": "\u0398",
        r"\Lambda": "\u039B", r"\Xi": "\u039E", r"\Pi": "\u03A0",
        r"\Sigma": "\u03A3", r"\Upsilon": "\u03A5", r"\Phi": "\u03A6",
        r"\Psi": "\u03A8", r"\Omega": "\u03A9",
        # Operators / big operators
        r"\int": "\u222B", r"\iint": "\u222C", r"\iiint": "\u222D",
        r"\oint": "\u222E", r"\sum": "\u2211", r"\prod": "\u220F",
        r"\coprod": "\u2210",
        # Roots and radicals
        r"\sqrt": "\u221A", r"\cbrt": "\u221B",
        # Calculus / analysis
        r"\partial": "\u2202", r"\nabla": "\u2207",
        r"\infty": "\u221E",
        # Logic / set theory
        r"\forall": "\u2200", r"\exists": "\u2203", r"\nexists": "\u2204",
        r"\in": "\u2208", r"\notin": "\u2209",
        r"\subset": "\u2282", r"\supset": "\u2283",
        r"\subseteq": "\u2286", r"\supseteq": "\u2287",
        r"\cup": "\u222A", r"\cap": "\u2229",
        r"\emptyset": "\u2205",
        r"\neg": "\u00AC", r"\land": "\u2227", r"\lor": "\u2228",
        # Arithmetic / relations
        r"\pm": "\u00B1", r"\mp": "\u2213",
        r"\times": "\u00D7", r"\div": "\u00F7", r"\cdot": "\u22C5",
        r"\leq": "\u2264", r"\geq": "\u2265", r"\neq": "\u2260",
        r"\approx": "\u2248", r"\equiv": "\u2261", r"\cong": "\u2245",
        r"\sim": "\u223C", r"\propto": "\u221D",
        r"\ll": "\u226A", r"\gg": "\u226B",
        # Arrows
        r"\rightarrow": "\u2192", r"\leftarrow": "\u2190",
        r"\leftrightarrow": "\u2194",
        r"\Rightarrow": "\u21D2", r"\Leftarrow": "\u21D0",
        r"\Leftrightarrow": "\u21D4",
        r"\uparrow": "\u2191", r"\downarrow": "\u2193",
        r"\mapsto": "\u21A6",
        # Dots
        r"\cdots": "\u22EF", r"\ldots": "\u2026", r"\vdots": "\u22EE",
        r"\ddots": "\u22F1",
        # Miscellaneous
        r"\angle": "\u2220", r"\degree": "\u00B0",
        r"\star": "\u22C6", r"\circ": "\u2218",
        r"\bullet": "\u2022", r"\diamond": "\u22C4",
        r"\triangle": "\u25B3",
        r"\hbar": "\u210F", r"\ell": "\u2113",
        r"\Re": "\u211C", r"\Im": "\u2124",
        r"\aleph": "\u2135",
        # Matrix (Word UnicodeMath uses ■ for matrix)
        r"\matrix": "\u25A0", r"\pmatrix": "\u25A0",
        # Function names (these stay as text but without backslash)
        r"\lim": "lim", r"\sin": "sin", r"\cos": "cos", r"\tan": "tan",
        r"\sec": "sec", r"\csc": "csc", r"\cot": "cot",
        r"\arcsin": "arcsin", r"\arccos": "arccos", r"\arctan": "arctan",
        r"\sinh": "sinh", r"\cosh": "cosh", r"\tanh": "tanh",
        r"\log": "log", r"\ln": "ln", r"\exp": "exp",
        r"\det": "det", r"\dim": "dim", r"\ker": "ker",
        r"\min": "min", r"\max": "max", r"\inf": "inf", r"\sup": "sup",
        r"\gcd": "gcd", r"\arg": "arg", r"\mod": "mod",
    }
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if not equation or not equation.strip():
            return json.dumps({"error": "equation text is required"})

        with undo_record(app, "MCP: Insert Equation"):
            # Determine insertion range
            if paragraph_index is not None:
                if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                    return json.dumps({
                        "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                    })
                rng = doc.Paragraphs(paragraph_index).Range
                rng.Collapse(0)  # After the paragraph
                rng.InsertParagraphAfter()
                rng.Collapse(0)
            elif position == "start":
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)  # Before first paragraph
                rng.InsertParagraphBefore()
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)
            else:  # "end"
                rng = doc.Content
                rng.Collapse(0)  # After last content
                rng.InsertParagraphAfter()
                rng.Collapse(0)

            # Convert LaTeX-like commands to Unicode math symbols.
            # Sort by length descending so longer matches take priority
            # (e.g. \iint before \int, \infty before \in).
            # Use negative lookahead (?![a-zA-Z]) to avoid partial matches.
            _commands = sorted(UNICODE_MATH.keys(), key=len, reverse=True)
            _pattern = '|'.join(re.escape(c) for c in _commands)
            _pattern = f'({_pattern})(?![a-zA-Z])'
            eq_text = re.sub(_pattern, lambda m: UNICODE_MATH[m.group(1)], equation)

            # Insert the converted equation text
            rng.Text = eq_text

            # Convert to OMath
            doc.OMaths.Add(rng)
            omath = doc.OMaths(doc.OMaths.Count)

            # Set display mode (centered on own line) vs inline
            if display_mode:
                omath.Type = 1  # wdOMathDisplay
            else:
                omath.Type = 0  # wdOMathInline

            # Build up the equation (render UnicodeMath to formatted equation)
            omath.BuildUp()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "equation": equation,
            "display_mode": display_mode,
            "omath_count": doc.OMaths.Count,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})
