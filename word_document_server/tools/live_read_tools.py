"""COM-based read, comment, and revision tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide read access and comment/revision management for locked files
that python-docx cannot open.
"""

import json
import sys

from word_document_server.defaults import DEFAULT_AUTHOR


async def word_live_get_text(filename: str = None) -> str:
    """Get all text from an open Word document, paragraph by paragraph.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with paragraphs list.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            text = doc.Paragraphs(i).Range.Text.rstrip("\r\x07")
            paragraphs.append({"index": i, "text": text})

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_paragraph_format(
    filename: str = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
) -> str:
    """[Windows only] Inspect paragraph formatting properties for diagnostics.

    Returns detailed formatting info for each paragraph in the range. Essential for
    debugging layout issues like unexpected page breaks (caused by keep_with_next chains),
    broken list formatting, wrong styles, or inconsistent fonts.

    Per-paragraph fields returned: index, text_preview (first 80 chars), char_start, char_end,
    style, font_name, font_size, bold, italic, alignment, space_before_pt, space_after_pt,
    line_spacing, line_spacing_rule, page_break_before, keep_with_next, keep_together.
    Also: list_type, list_level, list_string (if paragraph is in a list), highlight_color.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph (1-indexed, required).
        end_paragraph: Last paragraph (1-indexed, defaults to start_paragraph).

    Returns:
        JSON with formatting details per paragraph.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

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
                "error": f"Range {start_paragraph}-{end_paragraph} out of bounds (doc has {total_paras} paragraphs)"
            })

        ALIGN_NAMES = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
        SPACING_RULE_NAMES = {
            0: "single", 1: "1.5_lines", 2: "double",
            3: "at_least", 4: "exactly", 5: "multiple",
        }

        results = []
        for i in range(start_paragraph, end_paragraph + 1):
            para = doc.Paragraphs(i)
            rng = para.Range
            fmt = para.Format
            text = rng.Text.rstrip("\r\x07")
            preview = text[:80] + ("..." if len(text) > 80 else "")

            info = {
                "index": i,
                "text_preview": preview,
                "char_start": rng.Start,
                "char_end": rng.End,
                "style": str(rng.Style) if rng.Style else "",
                "font_name": str(rng.Font.Name) if rng.Font.Name else "",
                "font_size": rng.Font.Size if rng.Font.Size else None,
                "bold": bool(rng.Font.Bold) if rng.Font.Bold != 9999999 else "mixed",
                "italic": bool(rng.Font.Italic) if rng.Font.Italic != 9999999 else "mixed",
                "alignment": ALIGN_NAMES.get(fmt.Alignment, str(fmt.Alignment)),
                "space_before_pt": fmt.SpaceBefore,
                "space_after_pt": fmt.SpaceAfter,
                "line_spacing": fmt.LineSpacing,
                "line_spacing_rule": SPACING_RULE_NAMES.get(fmt.LineSpacingRule, str(fmt.LineSpacingRule)),
                "page_break_before": bool(fmt.PageBreakBefore),
                "keep_with_next": bool(fmt.KeepWithNext),
                "keep_together": bool(fmt.KeepTogether),
            }

            # List info
            try:
                lf = rng.ListFormat
                if lf.ListType > 0:
                    info["list_type"] = {1: "bullet", 2: "simple_number", 3: "upper_roman",
                                          4: "lower_roman", 5: "upper_letter", 6: "lower_letter"
                                          }.get(lf.ListType, f"type_{lf.ListType}")
                    info["list_level"] = lf.ListLevelNumber
                    info["list_string"] = lf.ListString
            except Exception:
                pass

            # Highlight
            try:
                info["highlight_color"] = rng.HighlightColorIndex
            except Exception:
                pass

            results.append(info)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs": results,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_info(filename: str = None) -> str:
    """Get document info from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with document metadata (pages, words, paragraphs, sections, etc.).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # wdStatistic constants
        WD_STAT_PAGES = 2
        WD_STAT_WORDS = 0
        WD_STAT_CHARACTERS = 3
        WD_STAT_LINES = 1

        info = {
            "name": doc.Name,
            "full_path": doc.FullName,
            "pages": doc.ComputeStatistics(WD_STAT_PAGES),
            "words": doc.ComputeStatistics(WD_STAT_WORDS),
            "characters": doc.ComputeStatistics(WD_STAT_CHARACTERS),
            "lines": doc.ComputeStatistics(WD_STAT_LINES),
            "paragraphs": doc.Paragraphs.Count,
            "sections": doc.Sections.Count,
            "tables": doc.Tables.Count,
            "comments": doc.Comments.Count,
            "track_revisions": doc.TrackRevisions,
            "saved": doc.Saved,
        }

        # Built-in properties (best effort)
        try:
            props = doc.BuiltInDocumentProperties
            info["author"] = str(props("Author").Value) if props("Author").Value else ""
            info["title"] = str(props("Title").Value) if props("Title").Value else ""
            info["subject"] = str(props("Subject").Value) if props("Subject").Value else ""
        except Exception:
            pass

        return json.dumps({"success": True, **info}, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_find_text(
    filename: str = None,
    search_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    max_results: int = 50,
) -> str:
    """Find text in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        search_text: Text to search for.
        match_case: Case-sensitive search.
        whole_word: Match whole words only.
        max_results: Maximum number of matches to return.

    Returns:
        JSON with list of matches (position, context).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not search_text:
        return json.dumps({"error": "search_text is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        matches = []
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()

        while len(matches) < max_results:
            found = rng.Find.Execute(
                FindText=search_text,
                MatchCase=match_case,
                MatchWholeWord=whole_word,
                Forward=True,
                Wrap=0,  # wdFindStop
            )
            if not found:
                break

            # Get surrounding context
            context_rng = rng.Duplicate
            context_start = max(0, rng.Start - 30)
            context_end = min(doc.Content.End, rng.End + 30)
            context_rng.SetRange(context_start, context_end)

            matches.append({
                "start": rng.Start,
                "end": rng.End,
                "text": rng.Text,
                "context": context_rng.Text,
            })

            # Move past current match
            rng.SetRange(rng.End, doc.Content.End)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "search_text": search_text,
            "match_count": len(matches),
            "matches": matches,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_get_comments(filename: str = None) -> str:
    """Get all comments from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of comments (author, date, text, scope).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        comments = []
        for i in range(1, doc.Comments.Count + 1):
            c = doc.Comments(i)
            scope_text = ""
            try:
                scope_text = c.Scope.Text[:100] if c.Scope and c.Scope.Text else ""
            except Exception:
                pass

            comments.append({
                "index": i,
                "author": str(c.Author) if c.Author else "",
                "date": str(c.Date) if c.Date else "",
                "text": str(c.Range.Text) if c.Range and c.Range.Text else "",
                "scope": scope_text,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_count": len(comments),
            "comments": comments,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_comment(
    filename: str = None,
    start: int = None,
    end: int = None,
    paragraph_index: int = None,
    text: str = "",
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Add a comment to an open Word document.

    Specify either start/end character positions or paragraph_index.
    If paragraph_index is given, the comment is attached to the entire paragraph.

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        paragraph_index: 1-indexed paragraph to attach comment to.
        text: Comment text.
        author: Comment author name.

    Returns:
        JSON with result info.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not text:
        return json.dumps({"error": "Comment text is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine the range to attach the comment to
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps({
                    "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                })
            rng = doc.Paragraphs(paragraph_index).Range
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
        else:
            return json.dumps({
                "error": "Provide either start/end positions or paragraph_index"
            })

        # Save and restore author
        prev_author = app.UserName
        app.UserName = author
        try:
            comment = doc.Comments.Add(rng, text)
        finally:
            app.UserName = prev_author

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "comment_index": comment.Index,
            "author": author,
            "text": text[:100],
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_revisions(filename: str = None) -> str:
    """List all tracked changes (revisions) in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of revisions (type, author, date, text).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # Revision type names
        REV_TYPES = {
            1: "insert",
            2: "delete",
            3: "property",
            4: "paragraph_number",
            5: "display_field",
            6: "reconcile",
            7: "conflict",
            8: "style",
            9: "replace",
            10: "section_property",
            11: "table_property",
            12: "cell_insert",
            13: "cell_delete",
            14: "cell_merge",
        }

        revisions = []
        for i in range(1, doc.Revisions.Count + 1):
            rev = doc.Revisions(i)
            rev_text = ""
            try:
                rev_text = rev.Range.Text[:200] if rev.Range and rev.Range.Text else ""
            except Exception:
                pass

            revisions.append({
                "index": i,
                "type": REV_TYPES.get(rev.Type, f"unknown({rev.Type})"),
                "type_id": rev.Type,
                "author": str(rev.Author) if rev.Author else "",
                "date": str(rev.Date) if rev.Date else "",
                "text": rev_text,
            })

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "revision_count": len(revisions),
            "revisions": revisions,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_accept_revisions(
    filename: str = None,
    author: str = None,
    revision_ids: list = None,
) -> str:
    """Accept tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only accept revisions by this author.
        revision_ids: List of 1-indexed revision IDs to accept. If None + no author, accept all.

    Returns:
        JSON with count of accepted revisions.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if revision_ids is not None:
            # Accept specific revisions (process in reverse to preserve indices)
            accepted = 0
            for rid in sorted(revision_ids, reverse=True):
                if 1 <= rid <= doc.Revisions.Count:
                    doc.Revisions(rid).Accept()
                    accepted += 1
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "accepted": accepted,
                "mode": "specific_ids",
            })

        if author:
            # Accept revisions by author (iterate in reverse)
            accepted = 0
            for i in range(doc.Revisions.Count, 0, -1):
                rev = doc.Revisions(i)
                if str(rev.Author) == author:
                    rev.Accept()
                    accepted += 1
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "accepted": accepted,
                "mode": f"by_author:{author}",
            })

        # Accept all
        total = doc.Revisions.Count
        doc.AcceptAllRevisions()
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "accepted": total,
            "mode": "all",
        })

    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_reject_revisions(
    filename: str = None,
    author: str = None,
    revision_ids: list = None,
) -> str:
    """Reject tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only reject revisions by this author.
        revision_ids: List of 1-indexed revision IDs to reject. If None + no author, reject all.

    Returns:
        JSON with count of rejected revisions.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        if revision_ids is not None:
            rejected = 0
            for rid in sorted(revision_ids, reverse=True):
                if 1 <= rid <= doc.Revisions.Count:
                    doc.Revisions(rid).Reject()
                    rejected += 1
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "rejected": rejected,
                "mode": "specific_ids",
            })

        if author:
            rejected = 0
            for i in range(doc.Revisions.Count, 0, -1):
                rev = doc.Revisions(i)
                if str(rev.Author) == author:
                    rev.Reject()
                    rejected += 1
            return json.dumps({
                "success": True,
                "document": doc.Name,
                "rejected": rejected,
                "mode": f"by_author:{author}",
            })

        total = doc.Revisions.Count
        doc.RejectAllRevisions()
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "rejected": total,
            "mode": "all",
        })

    except Exception as e:
        return json.dumps({"error": str(e)})
