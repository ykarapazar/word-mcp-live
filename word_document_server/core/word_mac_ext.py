"""Extended macOS automation helpers for word-mcp-live.

Implements the Word Live features previously missing on macOS:
- Table formatting (borders, cell bold, alignment, widths, shading, autofit)
- Multi-paragraph insertion
- Image insertion (inline & floating with sizing, alignment, borders, wrapping)
- Cross-references (list targets & insert cross-reference)
- Equations (LaTeX/UnicodeMath -> Word native OMML equation)
- Core document properties (built-in properties)
- Watermarks (diagonal WordArt / text watermark in section header)
- Threaded comment replies & resolve/unresolve (OOXML hybrid save-edit-reopen)
- MCP-level undo history tracking on macOS
"""

from __future__ import annotations

import datetime
import json
import os
import random
import re
import shutil
import subprocess
import tempfile
import time
from lxml import etree as ET
import zipfile
from typing import Any

from word_document_server.core.word_mac import (
    _escape_as,
    _escape_js,
    _run_applescript,
    _run_jxa,
)

# ---------------------------------------------------------------------------
# In-memory MCP Undo History Tracker for macOS
# ---------------------------------------------------------------------------

_MAC_UNDO_HISTORY: dict[str, list[dict[str, Any]]] = {}
_MAX_UNDO_ENTRIES = 100


def _norm_doc_key(doc_name: str | None) -> str:
    return (doc_name or "active").strip().lower()


def record_mac_undo(doc_name: str | None, action: str) -> None:
    """Record an MCP-initiated edit action for macOS undo history inspection."""
    key = _norm_doc_key(doc_name)
    stack = _MAC_UNDO_HISTORY.setdefault(key, [])
    stack.insert(
        0,
        {
            "action": action,
            "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
        },
    )
    if len(stack) > _MAX_UNDO_ENTRIES:
        del stack[_MAX_UNDO_ENTRIES:]


def pop_mac_undo(doc_name: str | None, times: int = 1) -> None:
    """Pop undone entries when mac_undo is invoked."""
    key = _norm_doc_key(doc_name)
    stack = _MAC_UNDO_HISTORY.get(key)
    if not stack:
        return
    del stack[: max(0, int(times))]


def mac_get_undo_history(filename: str | None = None, limit: int = 20) -> str:
    """Return MCP-tracked undo history for an open Word document on macOS."""
    doc_name = _resolve_open_doc_name(filename)
    key = _norm_doc_key(doc_name)
    stack = _MAC_UNDO_HISTORY.get(key, [])
    entries = [item["action"] for item in stack[: max(1, int(limit))]]
    detailed = stack[: max(1, int(limit))]
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "undo_entries": entries,
            "details": detailed,
            "count": len(entries),
            "note": (
                "Tracks MCP-initiated actions on macOS. Word for Mac does not "
                "expose manual UI undo stack entries via AppleScript/JXA."
            ),
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# Shared AppleScript / Color Helpers
# ---------------------------------------------------------------------------


def _as_doc_resolver(filename: str | None) -> str:
    """Return AppleScript snippet that sets `targetDoc` to the requested or active document."""
    if not filename:
        return """
        if (count of documents) is 0 then error "No documents are open in Word"
        set targetDoc to active document
        """
    esc = _escape_as(filename)
    base = _escape_as(os.path.basename(filename))
    return f"""
    if (count of documents) is 0 then error "No documents are open in Word"
    set targetDoc to missing value
    set docCount to count of documents
    repeat with di from 1 to docCount
        set curDoc to document di
        set dName to name of curDoc
        if dName is "{esc}" or dName is "{base}" or dName contains "{base}" then
            set targetDoc to curDoc
            exit repeat
        end if
    end repeat
    if targetDoc is missing value then error "Document '{esc}' not open in Word"
    """


def _resolve_open_doc_name(filename: str | None = None) -> str:
    """Resolve the active or named open document's name."""
    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        return name of targetDoc
    end tell
    """
    return _run_applescript(script).strip()


def _resolve_open_doc_path(filename: str | None = None) -> tuple[str, str]:
    """Return (doc_name, posix_full_path) for an open Word document on macOS."""
    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        set dName to name of targetDoc
        set dFull to ""
        try
            set dFull to full name of targetDoc
        end try
        set dPath to ""
        try
            set dPath to path of targetDoc
        end try
        return dName & "|||" & dFull & "|||" & dPath
    end tell
    """
    raw = _run_applescript(script).strip()
    parts = raw.split("|||")
    doc_name = parts[0] if len(parts) > 0 else ""
    full_name = parts[1] if len(parts) > 1 else ""
    dir_path = parts[2] if len(parts) > 2 else ""

    def _to_posix(p: str) -> str:
        if not p:
            return ""
        if p.startswith("/"):
            return p
        # Convert HFS path (Macintosh HD:Users:...) to POSIX path
        if ":" in p and not p.startswith("http"):
            conv = _run_applescript(f'return POSIX path of "{_escape_as(p)}"').strip()
            return conv
        return p

    posix_full = _to_posix(full_name)
    if posix_full and os.path.exists(posix_full):
        return doc_name, posix_full

    posix_dir = _to_posix(dir_path)
    if posix_dir and doc_name:
        candidate = os.path.join(posix_dir, doc_name)
        if os.path.exists(candidate):
            return doc_name, candidate

    return doc_name, posix_full


def _hex_to_as_rgb(hex_color: str) -> str:
    """Convert '#RRGGBB' or 'RRGGBB' to AppleScript 16-bit RGB tuple '{r, g, b}'."""
    c = hex_color.strip().lstrip("#")
    if len(c) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    r = int(c[0:2], 16) * 257
    g = int(c[2:4], 16) * 257
    b = int(c[4:6], 16) * 257
    return f"{{{r}, {g}, {b}}}"


# ---------------------------------------------------------------------------
# 1. Table Formatting (word_live_format_table)
# ---------------------------------------------------------------------------

_AS_BORDER_STYLES = {
    "none": "line style none",
    "single": "line style single",
    "double": "line style double",
    "dotted": "line style dot",
    "dashed": "line style dash small gap",
    "thick": "line style single",
}

_AS_ALIGN_MAP = {
    "left": "align paragraph left",
    "center": "align paragraph center",
    "right": "align paragraph right",
    "justify": "align paragraph justify",
}

_AS_TABLE_ALIGN_MAP = {
    "left": "align row left",
    "center": "align row center",
    "right": "align row right",
}

_AS_AUTOFIT_MAP = {
    "fixed": "auto fit fixed",
    "content": "auto fit content",
    "window": "auto fit window",
}


def mac_format_table(
    filename: str | None = None,
    table_index: int = -1,
    border_style: str | None = None,
    cell_bold: list | None = None,
    cell_alignment: list | None = None,
    column_widths: list | None = None,
    table_alignment: str | None = None,
    cell_shading: list | None = None,
    autofit: str | None = None,
) -> str:
    """Format a table in an open Word document on macOS via AppleScript."""
    if border_style is not None and border_style.lower() not in _AS_BORDER_STYLES:
        return json.dumps(
            {
                "error": f"Unknown border_style: {border_style}. Use: {list(_AS_BORDER_STYLES.keys())}"
            }
        )
    if table_alignment is not None and table_alignment.lower() not in _AS_TABLE_ALIGN_MAP:
        return json.dumps(
            {
                "error": f"Unknown table_alignment: {table_alignment}. Use: left, center, right"
            }
        )
    if autofit is not None and autofit.lower() not in _AS_AUTOFIT_MAP:
        return json.dumps(
            {"error": f"Unknown autofit: {autofit}. Use: window, content, fixed"}
        )

    ops: list[str] = []
    actions: list[str] = []

    # Borders
    if border_style is not None:
        bs_key = border_style.lower()
        as_style = _AS_BORDER_STYLES[bs_key]
        width_stmt = ""
        if bs_key == "thick":
            width_stmt = "set line width of b to line width225 point"
        ops.append(
            f"""
            set borderList to {{border top, border bottom, border left, border right, border horizontal, border vertical}}
            repeat with bt in borderList
                try
                    set b to get border tbl which border bt
                    set line style of b to {as_style}
                    {width_stmt}
                end try
            end repeat
            """
        )
        actions.append(f"border_style={bs_key}")

    # Autofit
    if autofit is not None:
        af_key = autofit.lower()
        ops.append(f"auto fit behavior tbl behavior {_AS_AUTOFIT_MAP[af_key]}")
        actions.append(f"autofit={af_key}")

    # Table alignment
    if table_alignment is not None:
        ta_key = table_alignment.lower()
        ops.append(
            f"set alignment of rows of tbl to {_AS_TABLE_ALIGN_MAP[ta_key]}"
        )
        actions.append(f"table_alignment={ta_key}")

    # Column widths (in points)
    if column_widths:
        for ci, width_pt in enumerate(column_widths, start=1):
            w = float(width_pt)
            ops.append(
                f"""
                if {ci} <= colCount then
                    set width of column {ci} of tbl to {w}
                end if
                """
            )
        actions.append(f"column_widths={len(column_widths)} cols")

    # Cell bold: [[row, col, bool], ...] (0 = all rows/cols)
    if cell_bold:
        for entry in cell_bold:
            r, c, is_bold = int(entry[0]), int(entry[1]), bool(entry[2])
            b_val = "true" if is_bold else "false"
            ops.append(
                _as_cell_loop(
                    r,
                    c,
                    f"set bold of font object of text object of theCell to {b_val}",
                )
            )
        actions.append(f"cell_bold={len(cell_bold)} entries")

    # Cell alignment: [[row, col, align], ...]
    if cell_alignment:
        for entry in cell_alignment:
            r, c, align_str = int(entry[0]), int(entry[1]), str(entry[2]).lower()
            if align_str not in _AS_ALIGN_MAP:
                return json.dumps(
                    {
                        "error": f"Unknown cell alignment: {align_str}. Use: left, center, right, justify"
                    }
                )
            ops.append(
                _as_cell_loop(
                    r,
                    c,
                    f"set alignment of paragraphs of text object of theCell to {_AS_ALIGN_MAP[align_str]}",
                )
            )
        actions.append(f"cell_alignment={len(cell_alignment)} entries")

    # Cell shading: [[row, col, "#RRGGBB"], ...]
    if cell_shading:
        for entry in cell_shading:
            r, c, color_hex = int(entry[0]), int(entry[1]), str(entry[2])
            rgb_tuple = _hex_to_as_rgb(color_hex)
            ops.append(
                _as_cell_loop(
                    r,
                    c,
                    f"set background pattern color of shading of theCell to {rgb_tuple}",
                )
            )
        actions.append(f"cell_shading={len(cell_shading)} entries")

    ops_str = "\n".join(ops)
    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        set tblCount to count of tables of targetDoc
        if tblCount is 0 then error "Document has no tables"
        set idx to {int(table_index)}
        if idx is -1 then set idx to tblCount
        if idx < 1 or idx > tblCount then error "Table index " & idx & " out of range (1-" & tblCount & ")"
        set tbl to table idx of targetDoc
        set rowCount to count of rows of tbl
        set colCount to count of columns of tbl
        {ops_str}
        return (name of targetDoc) & "|||" & idx & "|||" & rowCount & "|||" & colCount
    end tell
    """
    raw = _run_applescript(script, timeout=60).strip()
    parts = raw.split("|||")
    doc_name = parts[0]
    idx = int(parts[1])
    rows = int(parts[2])
    cols = int(parts[3])
    record_mac_undo(doc_name, f"Format table {idx} ({', '.join(actions) or 'no-op'})")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "table_index": idx,
            "rows": rows,
            "cols": cols,
            "actions": actions,
        },
        ensure_ascii=False,
    )


def _as_cell_loop(r: int, c: int, body_stmt: str) -> str:
    """Generate AppleScript loop over cell(s) handling 0 = all rows/cols."""
    if r == 0 and c == 0:
        return f"""
        repeat with ri from 1 to rowCount
            repeat with ci from 1 to colCount
                try
                    set theCell to get cell from table tbl row ri column ci
                    {body_stmt}
                end try
            end repeat
        end repeat
        """
    if r == 0:
        return f"""
        repeat with ri from 1 to rowCount
            try
                set theCell to get cell from table tbl row ri column {c}
                {body_stmt}
            end try
        end repeat
        """
    if c == 0:
        return f"""
        repeat with ci from 1 to colCount
            try
                set theCell to get cell from table tbl row {r} column ci
                {body_stmt}
            end try
        end repeat
        """
    return f"""
    if {r} >= 1 and {r} <= rowCount and {c} >= 1 and {c} <= colCount then
        try
            set theCell to get cell from table tbl row {r} column {c}
            {body_stmt}
        end try
    end if
    """


# ---------------------------------------------------------------------------
# 2. Multi-Paragraph Insert (word_live_insert_paragraphs)
# ---------------------------------------------------------------------------


def mac_insert_paragraphs(
    filename: str | None = None,
    paragraphs: list | None = None,
    target_text: str | None = None,
    target_paragraph_index: int | None = None,
    position: str = "after",
    style: str | None = None,
    track_changes: bool = False,
) -> str:
    """Insert one or more paragraphs before/after a target paragraph on macOS."""
    if not paragraphs or not isinstance(paragraphs, list):
        return json.dumps({"error": "paragraphs must be a non-empty list of strings"})
    if target_text is None and target_paragraph_index is None:
        return json.dumps(
            {"error": "Provide either target_text or target_paragraph_index"}
        )
    if target_text is not None and target_paragraph_index is not None:
        return json.dumps(
            {"error": "Provide target_text or target_paragraph_index, not both"}
        )
    if position not in ("before", "after"):
        return json.dumps(
            {"error": f"position must be 'before' or 'after', got '{position}'"}
        )

    applied_style = style or "Normal"
    paras_json = json.dumps([str(p) for p in paragraphs], ensure_ascii=False)
    target_text_js = (
        f'"{_escape_js(target_text)}"' if target_text is not None else "null"
    )
    target_idx_js = (
        str(int(target_paragraph_index))
        if target_paragraph_index is not None
        else "null"
    )
    style_js = f'"{_escape_js(applied_style)}"'
    tc_js = "true" if track_changes else "false"

    doc_finder = (
        f'var targetName = "{_escape_js(filename)}";'
        if filename
        else "var targetName = null;"
    )

    script = f"""
    var app = Application("Microsoft Word");
    var docs = app.documents();
    if (!docs || docs.length === 0) throw new Error("No documents are open in Word");
    {doc_finder}
    var d = null;
    if (!targetName) {{
        d = app.activeDocument;
    }} else {{
        var base = targetName.split("/").pop();
        for (var i = 0; i < docs.length; i++) {{
            var n = docs[i].name();
            if (n === targetName || n === base || n.indexOf(base) !== -1) {{
                d = docs[i];
                break;
            }}
        }}
        if (!d) throw new Error("Document '" + targetName + "' not open in Word");
    }}

    var paras = d.paragraphs();
    var pCount = paras.length;
    var targetIdx = {target_idx_js};
    var targetText = {target_text_js};

    if (targetIdx !== null) {{
        if (targetIdx < 0 || targetIdx >= pCount) {{
            throw new Error("target_paragraph_index " + targetIdx + " out of range (0-" + (pCount - 1) + ")");
        }}
    }} else {{
        targetIdx = -1;
        for (var i = 0; i < pCount; i++) {{
            var t = paras[i].textObject.content();
            if (t && t.indexOf(targetText) !== -1) {{
                targetIdx = i;
                break;
            }}
        }}
        if (targetIdx === -1) {{
            throw new Error("No paragraph found containing '" + targetText + "'");
        }}
    }}

    var prevTrack = d.trackRevisions();
    if ({tc_js}) d.trackRevisions = true;

    try {{
        var newParas = {paras_json};
        var pos = "{position}";
        var targetP = paras[targetIdx];
        var insertPos = (pos === "before")
            ? targetP.textObject.startOfContent()
            : targetP.textObject.endOfContent();

        var joined = newParas.join("\\r") + "\\r";
        var r = d.createRange({{start: insertPos, end: insertPos}});
        r.content = joined;

        // Apply style to the inserted paragraphs
        var updatedParas = d.paragraphs();
        var startApply = (pos === "before") ? targetIdx : targetIdx + 1;
        for (var k = 0; k < newParas.length; k++) {{
            var idx = startApply + k;
            if (idx >= 0 && idx < updatedParas.length) {{
                try {{
                    updatedParas[idx].textObject.style = {style_js};
                }} catch (e) {{}}
            }}
        }}
    }} finally {{
        if ({tc_js}) d.trackRevisions = prevTrack;
    }}

    JSON.stringify({{
        success: true,
        document: d.name(),
        paragraphs_inserted: {len(paragraphs)},
        target_paragraph_index: targetIdx,
        position: "{position}",
        style: {style_js}
    }});
    """
    out = _run_jxa(script, timeout=60).strip()
    data = json.loads(out)
    record_mac_undo(
        data.get("document"),
        f"Insert {len(paragraphs)} paragraph(s) {position} paragraph {data.get('target_paragraph_index')}",
    )
    return out


# ---------------------------------------------------------------------------
# 3. Image Insertion (word_live_insert_image)
# ---------------------------------------------------------------------------

_AS_WRAP_MAP = {
    "inline": None,
    "square": "wrap square",
    "tight": "wrap tight",
    "behind": "wrap behind",
    "infront": "wrap front",
    "topbottom": "wrap top bottom",
}


_OFFICE_GROUP_CONTAINER = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/User Content.localized"
)


def _stage_image_for_word_sandbox(abs_path: str) -> tuple[str, bool]:
    """Copy an image into Word's sandbox-accessible Group Container to avoid macOS Grant Access dialogs.

    Returns (path_for_word, should_cleanup).
    """
    if not os.path.isdir(_OFFICE_GROUP_CONTAINER):
        return abs_path, False
    staging_dir = os.path.join(_OFFICE_GROUP_CONTAINER, "MCPStaging")
    os.makedirs(staging_dir, exist_ok=True)
    ext = os.path.splitext(abs_path)[1] or ".png"
    fd, staged_path = tempfile.mkstemp(prefix="img_", suffix=ext, dir=staging_dir)
    os.close(fd)
    shutil.copy2(abs_path, staged_path)
    return staged_path, True


def mac_insert_image(
    filename: str | None = None,
    image_path: str = "",
    paragraph_index: int | None = None,
    position: str = "end",
    width_inches: float | None = None,
    height_inches: float | None = None,
    width_pt: float | None = None,
    height_pt: float | None = None,
    alignment: str | None = None,
    wrapping: str | None = None,
    border_style: str | None = None,
    border_width_pt: float | None = None,
    border_color: str | None = None,
    link_to_file: bool = False,
) -> str:
    """Insert an image into an open Word document on macOS via AppleScript."""
    if not image_path:
        return json.dumps({"error": "image_path is required"})
    abs_path = os.path.abspath(os.path.expanduser(image_path))
    if not os.path.exists(abs_path):
        return json.dumps({"error": f"Image file not found: {abs_path}"})

    if wrapping is not None and wrapping.lower() not in _AS_WRAP_MAP:
        return json.dumps(
            {
                "error": f"Unknown wrapping: {wrapping}. Use: {list(_AS_WRAP_MAP.keys())}"
            }
        )
    if border_style is not None and border_style.lower() not in _AS_BORDER_STYLES:
        return json.dumps(
            {
                "error": f"Unknown border_style: {border_style}. Use: {list(_AS_BORDER_STYLES.keys())}"
            }
        )
    if alignment is not None and alignment.lower() not in _AS_ALIGN_MAP:
        return json.dumps(
            {"error": f"Unknown alignment: {alignment}. Use: left, center, right, justify"}
        )

    final_w_pt = width_pt if width_pt is not None else (width_inches * 72.0 if width_inches is not None else None)
    final_h_pt = height_pt if height_pt is not None else (height_inches * 72.0 if height_inches is not None else None)

    # Target range selection in AppleScript
    if paragraph_index is not None:
        p_idx = int(paragraph_index)
        range_as = f"""
        set pCount to count of paragraphs of targetDoc
        if {p_idx} < 1 or {p_idx} > pCount then error "paragraph_index {p_idx} out of range (1-" & pCount & ")"
        set pRng to text object of paragraph {p_idx} of targetDoc
        set posStart to start of content of pRng
        set targetRng to create range targetDoc start posStart end posStart
        """
    elif str(position).lower() == "start":
        range_as = "set targetRng to create range targetDoc start 0 end 0"
    elif str(position).lower() == "end":
        range_as = """
        set endPos to (end of content of text object of targetDoc) - 1
        if endPos < 0 then set endPos to 0
        set targetRng to create range targetDoc start endPos end endPos
        """
    else:
        try:
            off = max(0, int(position))
        except (ValueError, TypeError):
            off = 0
        range_as = f"set targetRng to create range targetDoc start {off} end {off}"

    # Sizing statements (preserving aspect ratio when only one dimension is given)
    size_as = ""
    if final_w_pt is not None and final_h_pt is not None:
        size_as = f"""
        set width of pic to {float(final_w_pt)}
        set height of pic to {float(final_h_pt)}
        """
    elif final_w_pt is not None:
        size_as = f"""
        set origW to width of pic
        set origH to height of pic
        set width of pic to {float(final_w_pt)}
        if origW > 0 then set height of pic to (origH * ({float(final_w_pt)} / origW))
        """
    elif final_h_pt is not None:
        size_as = f"""
        set origW to width of pic
        set origH to height of pic
        set height of pic to {float(final_h_pt)}
        if origH > 0 then set width of pic to (origW * ({float(final_h_pt)} / origH))
        """

    # Alignment statement
    align_as = ""
    if alignment is not None:
        align_as = f"""
        try
            set alignment of paragraph 1 of text object of targetRng to {_AS_ALIGN_MAP[alignment.lower()]}
        end try
        """

    # Border statement (inline shape borders)
    border_as = ""
    if border_style is not None:
        as_bs = _AS_BORDER_STYLES[border_style.lower()]
        color_stmt = ""
        if border_color:
            rgb_t = _hex_to_as_rgb(border_color)
            color_stmt = f"set color of b to {rgb_t}"
        border_as = f"""
        repeat with bt in {{border top, border bottom, border left, border right}}
            try
                set b to get border pic which border bt
                set line style of b to {as_bs}
                {color_stmt}
            end try
        end repeat
        """

    # Wrapping (convert to floating shape if non-inline)
    wrap_as = ""
    if wrapping and wrapping.lower() != "inline":
        as_wrap = _AS_WRAP_MAP[wrapping.lower()]
        wrap_as = f"""
        try
            set fltShape to convert to shape pic
            set wrap type of wrap format of fltShape to {as_wrap}
        end try
        """

    link_val = "true" if link_to_file else "false"
    save_val = "false" if link_to_file else "true"

    word_img_path, cleanup_staged = (
        (abs_path, False) if link_to_file else _stage_image_for_word_sandbox(abs_path)
    )
    try:
        script = f"""
        tell application "Microsoft Word"
            {_as_doc_resolver(filename)}
            {range_as}
            make new inline picture at targetRng with properties {{file name: (POSIX file "{_escape_as(word_img_path)}" as text), link to file: {link_val}, save with document: {save_val}}}
            set pic to inline shape (count of inline shapes of targetDoc) of targetDoc
            {size_as}
            set finalW to width of pic
            set finalH to height of pic
            {border_as}
            {align_as}
            {wrap_as}
            return (name of targetDoc) & "|||" & finalW & "|||" & finalH
        end tell
        """
        raw = _run_applescript(script, timeout=60).strip()
    finally:
        if cleanup_staged and os.path.exists(word_img_path):
            try:
                os.unlink(word_img_path)
            except OSError:
                pass
    parts = raw.split("|||")
    doc_name = parts[0]
    out_w = round(float(parts[1].replace(",", ".")), 1) if len(parts) > 1 else final_w_pt
    out_h = round(float(parts[2].replace(",", ".")), 1) if len(parts) > 2 else final_h_pt

    record_mac_undo(doc_name, f"Insert image {os.path.basename(abs_path)}")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "image_path": abs_path,
            "width_pt": out_w,
            "height_pt": out_h,
            "wrapping": (wrapping or "inline").lower(),
            "alignment": alignment,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# 4. Cross-References (list & insert)
# ---------------------------------------------------------------------------

_AS_REF_TYPE_MAP = {
    "heading": "reference type heading",
    "bookmark": "reference type bookmark",
    "footnote": "reference type footnote",
    "endnote": "reference type endnote",
    "numbered_item": "reference type numbered item",
}

_AS_REF_KIND_MAP = {
    "text": "reference content text",
    "page_number": "reference page number",
    "number": "reference number",
    "number_no_context": "reference number no context",
    "number_full_context": "reference number full context",
    "above_below": "reference position",
    "entire_caption": "reference entire caption",
    "only_label_and_number": "reference only label and number",
    "only_caption_text": "reference only caption text",
}


def mac_list_cross_reference_items(
    filename: str | None = None,
    ref_type: str = "heading",
) -> str:
    """List available cross-reference targets in an open Word document on macOS."""
    rt_key = ref_type.lower().strip()
    if rt_key in _AS_REF_TYPE_MAP:
        rt_expr = _AS_REF_TYPE_MAP[rt_key]
    else:
        # Custom caption label (e.g., "Figure", "Table", "Gambar", "Tabel")
        rt_expr = f'"{_escape_as(ref_type)}"'

    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        set dName to name of targetDoc
        set itemsList to get cross reference items targetDoc reference type {rt_expr}
        set outText to dName
        if itemsList is not missing value then
            repeat with itm in itemsList
                set outText to outText & "|||" & (itm as text)
            end repeat
        end if
        return outText
    end tell
    """
    raw = _run_applescript(script, timeout=30).strip()
    parts = raw.split("|||")
    doc_name = parts[0]
    raw_items = parts[1:] if len(parts) > 1 else []
    items = [
        {"index": i, "text": txt.strip()}
        for i, txt in enumerate(raw_items, start=1)
        if txt.strip()
    ]
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "ref_type": ref_type,
            "items": items,
            "count": len(items),
        },
        ensure_ascii=False,
    )


def mac_insert_cross_reference(
    filename: str | None = None,
    ref_type: str = "heading",
    ref_item: str = "1",
    ref_kind: str = "text",
    insert_as_hyperlink: bool = True,
    paragraph_index: int | None = None,
    insert_position: str = "end",
) -> str:
    """Insert a cross-reference field into an open Word document on macOS."""
    rt_key = ref_type.lower().strip()
    rk_key = ref_kind.lower().strip()
    if rk_key not in _AS_REF_KIND_MAP:
        return json.dumps(
            {
                "error": f"Invalid ref_kind '{ref_kind}'. Use: {', '.join(_AS_REF_KIND_MAP.keys())}"
            }
        )
    rt_expr = (
        _AS_REF_TYPE_MAP[rt_key]
        if rt_key in _AS_REF_TYPE_MAP
        else f'"{_escape_as(ref_type)}"'
    )
    rk_expr = _AS_REF_KIND_MAP[rk_key]
    hyp_val = "true" if insert_as_hyperlink else "false"

    if paragraph_index is not None:
        p_idx = int(paragraph_index)
        range_as = f"""
        set pCount to count of paragraphs of targetDoc
        if {p_idx} < 1 or {p_idx} > pCount then error "paragraph_index {p_idx} out of range (1-" & pCount & ")"
        set pRng to text object of paragraph {p_idx} of targetDoc
        set posEnd to (end of content of pRng) - 1
        if posEnd < 0 then set posEnd to 0
        set targetRng to create range targetDoc start posEnd end posEnd
        """
    elif str(insert_position).lower() == "start":
        range_as = "set targetRng to create range targetDoc start 0 end 0"
    elif str(insert_position).lower() == "end":
        range_as = """
        set endPos to (end of content of text object of targetDoc) - 1
        if endPos < 0 then set endPos to 0
        set targetRng to create range targetDoc start endPos end endPos
        """
    else:
        try:
            off = max(0, int(insert_position))
        except (ValueError, TypeError):
            off = 0
        range_as = f"set targetRng to create range targetDoc start {off} end {off}"

    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        {range_as}
        insert cross reference at targetRng reference type {rt_expr} reference kind {rk_expr} reference item "{_escape_as(str(ref_item))}" insert as hyperlink {hyp_val}
        return name of targetDoc
    end tell
    """
    doc_name = _run_applescript(script, timeout=30).strip()
    record_mac_undo(doc_name, f"Insert cross-reference ({ref_type} #{ref_item})")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "ref_type": ref_type,
            "ref_item": ref_item,
            "ref_kind": ref_kind,
            "as_hyperlink": insert_as_hyperlink,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# 5. Equation Insertion (word_live_insert_equation)
# ---------------------------------------------------------------------------


def _latex_to_unicodemath(expr: str, symbol_map: dict[str, str]) -> str:
    """Convert common LaTeX math constructs to Word UnicodeMath."""
    s = expr.strip()
    # \frac{a}{b} -> (a)/(b)
    for _ in range(10):
        m = re.search(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", s)
        if not m:
            break
        s = s[: m.start()] + f"({m.group(1)})/({m.group(2)})" + s[m.end() :]
    # \sqrt[n]{x} -> √(n&x), \sqrt{x} -> √(x)
    s = re.sub(r"\\sqrt\[([^\]]+)\]\{([^{}]+)\}", "\u221A(\\1&\\2)", s)
    for _ in range(10):
        m = re.search(r"\\sqrt\{([^{}]+)\}", s)
        if not m:
            break
        s = s[: m.start()] + f"\u221A({m.group(1)})" + s[m.end() :]
    # Replace longest LaTeX commands first
    for cmd in sorted(symbol_map.keys(), key=len, reverse=True):
        s = s.replace(cmd, symbol_map[cmd])
    # Convert ^{...} and _{...} to ^(...) and _(...)
    s = re.sub(r"\^\{([^{}]+)\}", r"^(\1)", s)
    s = re.sub(r"_\{([^{}]+)\}", r"_(\1)", s)
    return s


def mac_insert_equation(
    filename: str | None = None,
    equation_text: str = "",
    paragraph_index: int | None = None,
    position: str = "end",
    display_mode: bool = True,
    symbol_map: dict[str, str] | None = None,
) -> str:
    """Insert a native OMML equation into an open Word document on macOS."""
    if not equation_text or not equation_text.strip():
        return json.dumps({"error": "equation_text is required"})

    converted = (
        _latex_to_unicodemath(equation_text, symbol_map)
        if symbol_map
        else equation_text.strip()
    )

    doc_finder = (
        f'var targetName = "{_escape_js(filename)}";'
        if filename
        else "var targetName = null;"
    )
    p_idx_js = str(int(paragraph_index)) if paragraph_index is not None else "null"
    pos_js = f'"{_escape_js(str(position))}"'
    eq_js = json.dumps(converted, ensure_ascii=False)
    disp_js = "true" if display_mode else "false"

    script = f"""
    var app = Application("Microsoft Word");
    var docs = app.documents();
    if (!docs || docs.length === 0) throw new Error("No documents are open in Word");
    {doc_finder}
    var d = null;
    if (!targetName) {{
        d = app.activeDocument;
    }} else {{
        var base = targetName.split("/").pop();
        for (var i = 0; i < docs.length; i++) {{
            var n = docs[i].name();
            if (n === targetName || n === base || n.indexOf(base) !== -1) {{
                d = docs[i];
                break;
            }}
        }}
        if (!d) throw new Error("Document '" + targetName + "' not open in Word");
    }}

    var pIdx = {p_idx_js};
    var pos = {pos_js};
    var startPos = 0;
    if (pIdx !== null) {{
        var paras = d.paragraphs();
        if (pIdx < 1 || pIdx > paras.length) throw new Error("paragraph_index " + pIdx + " out of range (1-" + paras.length + ")");
        startPos = Math.max(0, paras[pIdx - 1].textObject.endOfContent() - 1);
    }} else if (pos.toLowerCase() === "start") {{
        startPos = 0;
    }} else if (pos.toLowerCase() === "end") {{
        startPos = Math.max(0, d.textObject.endOfContent() - 1);
    }} else {{
        var parsed = parseInt(pos, 10);
        startPos = isNaN(parsed) ? 0 : Math.max(0, parsed);
    }}

    var eqStr = {eq_js};
    var isDisp = {disp_js};
    var r = d.createRange({{start: startPos, end: startPos}});
    r.content = isDisp ? ("\\r" + eqStr) : eqStr;
    var eqStart = isDisp ? (startPos + 1) : startPos;
    var eqRng = d.createRange({{start: eqStart, end: eqStart + eqStr.length}});
    var eqObj = app.createNewEquation({{fromRange: eqRng, inDocument: d}});
    try {{
        eqObj.buildUp();
    }} catch (e) {{}}
    d.name();
    """
    doc_name = _run_jxa(script, timeout=30).strip()
    record_mac_undo(doc_name, f"Insert equation: {equation_text[:40]}")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "equation": equation_text,
            "unicode_math": converted,
            "display_mode": display_mode,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# 6. Core Document Properties (word_live_set_core_properties)
# ---------------------------------------------------------------------------

_MAC_CORE_PROP_MAP = {
    "title": "Title",
    "subject": "Subject",
    "author": "Author",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "manager": "Manager",
    "company": "Company",
    "last_author": "Last author",
}


def mac_set_core_properties(
    filename: str | None = None,
    title: str | None = None,
    subject: str | None = None,
    author: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
    category: str | None = None,
    manager: str | None = None,
    company: str | None = None,
    last_author: str | None = None,
) -> str:
    """Set built-in document properties on an open Word document on macOS."""
    inputs = {
        "title": title,
        "subject": subject,
        "author": author,
        "keywords": keywords,
        "comments": comments,
        "category": category,
        "manager": manager,
        "company": company,
        "last_author": last_author,
    }
    to_set = {k: v for k, v in inputs.items() if v is not None}
    if not to_set:
        doc_name = _resolve_open_doc_name(filename)
        return json.dumps(
            {"ok": True, "document": doc_name, "changed": {}, "errors": None},
            ensure_ascii=False,
        )

    stmts: list[str] = []
    for key, value in to_set.items():
        prop_name = _MAC_CORE_PROP_MAP[key]
        esc_prop = _escape_as(prop_name)
        esc_val = _escape_as(str(value))
        stmts.append(
            f"""
            try
                set oldVal to ""
                try
                    set oldVal to (value of document property "{esc_prop}" of targetDoc) as text
                end try
                set value of document property "{esc_prop}" of targetDoc to "{esc_val}"
                set outLines to outLines & ("OK|||" & "{key}" & "|||" & oldVal & "|||" & "{esc_val}" & linefeed)
            on error errMsg
                set outLines to outLines & ("ERR|||" & "{key}" & "|||" & errMsg & linefeed)
            end try
            """
        )

    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        set outLines to "DOC|||" & (name of targetDoc) & linefeed
        {"".join(stmts)}
        return outLines
    end tell
    """
    raw = _run_applescript(script, timeout=30)
    doc_name = filename or "active"
    changes: dict[str, dict[str, str | None]] = {}
    errors: dict[str, str] = {}
    for line in raw.splitlines():
        parts = line.split("|||")
        if not parts:
            continue
        tag = parts[0]
        if tag == "DOC" and len(parts) >= 2:
            doc_name = parts[1]
        elif tag == "OK" and len(parts) >= 4:
            changes[parts[1]] = {"old": parts[2] or None, "new": parts[3]}
        elif tag == "ERR" and len(parts) >= 3:
            errors[parts[1]] = parts[2]

    if changes:
        record_mac_undo(doc_name, f"Set core properties ({', '.join(changes.keys())})")

    return json.dumps(
        {
            "ok": len(errors) == 0,
            "document": doc_name,
            "changed": changes,
            "errors": errors or None,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# 7. Watermark (word_live_add_watermark)
# ---------------------------------------------------------------------------


def mac_add_watermark(
    filename: str | None = None,
    text: str = "TASLAK",
    font_size: int = 72,
    font_color: str = "C0C0C0",
    rotation: int = -45,
    section_index: int = 1,
) -> str:
    """Add a diagonal text watermark to a section header on macOS via AppleScript."""
    rgb_tuple = _hex_to_as_rgb(font_color)
    esc_text = _escape_as(text)
    rot_val = float(rotation)
    fs_val = float(font_size)
    sec_idx = int(section_index)

    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        set secCount to count of sections of targetDoc
        if {sec_idx} < 1 or {sec_idx} > secCount then error "Section {sec_idx} out of range (1-" & secCount & ")"
        set hdr to get header (section {sec_idx} of targetDoc) index header footer primary

        set wmShape to missing value
        try
            set wmShape to make new word art at hdr with properties {{word art text: "{esc_text}", font name: "Calibri", font size: {fs_val}, bold: false, italic: false, preset word art effect: wordart format1, left position: 100, top: 300}}
        on error
            set wmShape to make new text box at hdr with properties {{left position: 100, top: 300, width: 420, height: 140}}
            set content of text object of text frame of wmShape to "{esc_text}"
            set font size of font object of text object of text frame of wmShape to {fs_val}
            set name of font object of text object of text frame of wmShape to "Calibri"
            set color of font object of text object of text frame of wmShape to {rgb_tuple}
            set visible of fill format of wmShape to false
        end try

        try
            set fore color of fill format of wmShape to {rgb_tuple}
        end try
        try
            set transparency of fill format of wmShape to 0.5
        end try
        try
            set visible of line format of wmShape to false
        end try
        try
            set rotation of wmShape to {rot_val}
        end try
        try
            set relative horizontal position of wmShape to relative horizontal position margin
            set relative vertical position of wmShape to relative vertical position margin
            set left position of wmShape to -999995
            set top of wmShape to -999995
        end try
        try
            set wrap type of wrap format of wmShape to wrap behind
            set allow overlap of wrap format of wmShape to true
        end try

        return name of targetDoc
    end tell
    """
    doc_name = _run_applescript(script, timeout=45).strip()
    record_mac_undo(doc_name, f"Add watermark '{text}' (section {sec_idx})")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "text": text,
            "font_size": font_size,
            "color": font_color,
            "rotation": rotation,
            "section": section_index,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# 8. Threaded Comment Replies & Resolve/Unresolve (OOXML Hybrid)
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_COMMENTS_EX_REL_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
_COMMENTS_EX_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)


def _gen_para_id() -> str:
    """Generate a valid 8-hex-digit Word paraId (< 0x80000000)."""
    return f"{random.randint(0x10000000, 0x7FFFFFFF):08X}"


def _save_and_close_for_ooxml(filename: str | None) -> tuple[str, str]:
    """Save the open document in Word, close it cleanly, and return (doc_name, posix_path)."""
    doc_name, posix_path = _resolve_open_doc_path(filename)
    if not posix_path or not os.path.exists(posix_path):
        raise RuntimeError(
            f"Document '{doc_name}' is not saved locally on disk. Save it as a local .docx file first."
        )
    if not posix_path.lower().endswith(".docx"):
        raise RuntimeError(
            f"OOXML comment operations require a .docx file (got '{os.path.basename(posix_path)}')."
        )
    script = f"""
    tell application "Microsoft Word"
        {_as_doc_resolver(filename)}
        save targetDoc
        close targetDoc saving no
    end tell
    """
    _run_applescript(script, timeout=30)
    return doc_name, posix_path


def _reopen_in_word(posix_path: str) -> None:
    """Reopen the modified .docx file in Microsoft Word and wait until it appears open."""
    base = _escape_as(os.path.basename(posix_path))
    _run_applescript(
        f"""
        tell application "Microsoft Word"
            open POSIX file "{_escape_as(posix_path)}"
        end tell
        """,
        timeout=30,
    )
    for _ in range(20):
        check = _run_applescript(
            f"""
            tell application "Microsoft Word"
                repeat with di from 1 to (count of documents)
                    if (name of document di) is "{base}" then return "OPEN"
                end repeat
                return "WAIT"
            end tell
            """,
            timeout=10,
        ).strip()
        if check == "OPEN":
            return
        time.sleep(0.25)


def _rewrite_docx_zip(posix_path: str, updates: dict[str, bytes]) -> None:
    """Safely rewrite entries inside a .docx zip archive in-place.

    Writes and verifies a temporary zip first, then overwrites the target file
    in-place (preserving its inode) so macOS App Sandbox security-scoped
    bookmarks in Microsoft Word remain valid without re-prompting Grant Access.
    """
    dir_name = os.path.dirname(posix_path)
    fd, tmp_path = tempfile.mkstemp(prefix=".word_mcp_", suffix=".docx", dir=dir_name)
    os.close(fd)
    try:
        with zipfile.ZipFile(posix_path, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            written = set()
            for item in zin.infolist():
                if item.filename in updates:
                    zout.writestr(item, updates[item.filename])
                    written.add(item.filename)
                else:
                    zout.writestr(item, zin.read(item.filename))
            for fname, payload in updates.items():
                if fname not in written:
                    zout.writestr(fname, payload)
        # Verify archive integrity before touching the original file
        with zipfile.ZipFile(tmp_path, "r") as zcheck:
            bad = zcheck.testzip()
            if bad is not None:
                raise RuntimeError(f"Corrupt zip entry after OOXML edit: {bad}")
        # Copy bytes in-place to preserve inode for macOS App Sandbox bookmark
        with open(tmp_path, "rb") as src, open(posix_path, "wb") as dst:
            shutil.copyfileobj(src, dst)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _ensure_comments_extended_wiring(
    zf: zipfile.ZipFile, updates: dict[str, bytes]
) -> None:
    """Ensure [Content_Types].xml and word/_rels/document.xml.rels register commentsExtended.xml.

    Uses targeted insertion before closing tags so original default XML namespaces
    are preserved byte-for-byte for Microsoft Word compatibility.
    """
    # 1. [Content_Types].xml
    ct_text = zf.read("[Content_Types].xml").decode("utf-8")
    if "/word/commentsExtended.xml" not in ct_text:
        override_tag = (
            f'<Override PartName="/word/commentsExtended.xml" '
            f'ContentType="{_COMMENTS_EX_CT}"/>'
        )
        ct_text = ct_text.replace("</Types>", f"{override_tag}</Types>", 1)
        updates["[Content_Types].xml"] = ct_text.encode("utf-8")

    # 2. word/_rels/document.xml.rels
    rels_path = "word/_rels/document.xml.rels"
    rels_text = zf.read(rels_path).decode("utf-8")
    if _COMMENTS_EX_REL_TYPE not in rels_text:
        existing_nums = [
            int(m) for m in re.findall(r'Id="rId(\d+)"', rels_text)
        ]
        next_num = (max(existing_nums) + 1) if existing_nums else 1
        rel_tag = (
            f'<Relationship Id="rId{next_num}" '
            f'Type="{_COMMENTS_EX_REL_TYPE}" '
            f'Target="commentsExtended.xml"/>'
        )
        rels_text = rels_text.replace(
            "</Relationships>", f"{rel_tag}</Relationships>", 1
        )
        updates[rels_path] = rels_text.encode("utf-8")


def _load_or_init_comments_ex(zf: zipfile.ZipFile) -> ET._Element:
    if "word/commentsExtended.xml" in zf.namelist():
        return ET.fromstring(zf.read("word/commentsExtended.xml"))
    return ET.Element(f"{{{_W15_NS}}}commentsEx", nsmap={"w15": _W15_NS})


def _get_or_assign_comment_para_id(comment_el: ET.Element) -> str:
    """Return the last paragraph's w14:paraId inside a <w:comment>, creating one if needed."""
    paras = comment_el.findall(f"{{{_W_NS}}}p")
    if not paras:
        p = ET.SubElement(comment_el, f"{{{_W_NS}}}p")
        paras = [p]
    target_p = paras[-1]
    pid = target_p.get(f"{{{_W14_NS}}}paraId")
    if not pid:
        pid = _gen_para_id()
        target_p.set(f"{{{_W14_NS}}}paraId", pid)
        if not target_p.get(f"{{{_W14_NS}}}textId"):
            target_p.set(f"{{{_W14_NS}}}textId", "77777777")
    return pid


def mac_reply_to_comment(
    filename: str | None = None,
    comment_index: int = 1,
    text: str = "",
    author: str | None = None,
) -> str:
    """Add a threaded reply to an existing comment on macOS via OOXML hybrid edit."""
    if not comment_index or int(comment_index) < 1:
        return json.dumps({"error": "comment_index must be >= 1"})
    if not text or not text.strip():
        return json.dumps({"error": "Reply text is required"})

    reply_author = author or os.environ.get("MCP_AUTHOR") or "User"
    reply_initials = os.environ.get("MCP_AUTHOR_INITIALS") or "".join(
        w[0].upper() for w in reply_author.split()[:2] if w
    )

    doc_name, posix_path = _save_and_close_for_ooxml(filename)
    try:
        updates: dict[str, bytes] = {}
        with zipfile.ZipFile(posix_path, "r") as zf:
            if "word/comments.xml" not in zf.namelist():
                raise RuntimeError("Document has no comments")
            c_root = ET.fromstring(zf.read("word/comments.xml"))
            comments = c_root.findall(f"{{{_W_NS}}}comment")
            idx = int(comment_index)
            if idx < 1 or idx > len(comments):
                raise RuntimeError(
                    f"comment_index {idx} out of range (1-{len(comments)})"
                )
            parent_el = comments[idx - 1]
            parent_para_id = _get_or_assign_comment_para_id(parent_el)

            # Determine next comment id
            max_id = max(
                (int(c.get(f"{{{_W_NS}}}id", "0")) for c in comments), default=0
            )
            new_cid = str(max_id + 1)
            new_para_id = _gen_para_id()
            now_iso = datetime.datetime.now(datetime.timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%SZ"
            )

            reply_el = ET.SubElement(c_root, f"{{{_W_NS}}}comment")
            reply_el.set(f"{{{_W_NS}}}id", new_cid)
            reply_el.set(f"{{{_W_NS}}}author", reply_author)
            reply_el.set(f"{{{_W_NS}}}date", now_iso)
            reply_el.set(f"{{{_W_NS}}}initials", reply_initials)
            p_el = ET.SubElement(reply_el, f"{{{_W_NS}}}p")
            p_el.set(f"{{{_W14_NS}}}paraId", new_para_id)
            p_el.set(f"{{{_W14_NS}}}textId", "77777777")
            r_el = ET.SubElement(p_el, f"{{{_W_NS}}}r")
            t_el = ET.SubElement(r_el, f"{{{_W_NS}}}t")
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_el.text = text

            # Update commentsExtended.xml
            ex_root = _load_or_init_comments_ex(zf)
            has_parent_ex = any(
                el.get(f"{{{_W15_NS}}}paraId") == parent_para_id
                for el in ex_root.findall(f"{{{_W15_NS}}}commentEx")
            )
            if not has_parent_ex:
                pex = ET.SubElement(ex_root, f"{{{_W15_NS}}}commentEx")
                pex.set(f"{{{_W15_NS}}}paraId", parent_para_id)
                pex.set(f"{{{_W15_NS}}}done", "0")

            rex = ET.SubElement(ex_root, f"{{{_W15_NS}}}commentEx")
            rex.set(f"{{{_W15_NS}}}paraId", new_para_id)
            rex.set(f"{{{_W15_NS}}}paraIdParent", parent_para_id)
            rex.set(f"{{{_W15_NS}}}done", "0")

            updates["word/comments.xml"] = ET.tostring(
                c_root, encoding="utf-8", xml_declaration=True
            )
            updates["word/commentsExtended.xml"] = ET.tostring(
                ex_root, encoding="utf-8", xml_declaration=True
            )
            _ensure_comments_extended_wiring(zf, updates)

        _rewrite_docx_zip(posix_path, updates)
    finally:
        _reopen_in_word(posix_path)

    record_mac_undo(doc_name, f"Reply to comment #{comment_index} (OOXML)")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "parent_comment_index": int(comment_index),
            "reply_comment_id": new_cid,
            "reply_text": text,
            "author": reply_author,
            "method": "ooxml_hybrid",
        },
        ensure_ascii=False,
    )


def mac_resolve_comment(
    filename: str | None = None,
    comment_index: int = 1,
    resolve: bool = True,
) -> str:
    """Resolve or unresolve a comment on macOS via OOXML commentsExtended.xml."""
    if not comment_index or int(comment_index) < 1:
        return json.dumps({"error": "comment_index must be >= 1"})

    doc_name, posix_path = _save_and_close_for_ooxml(filename)
    try:
        updates: dict[str, bytes] = {}
        with zipfile.ZipFile(posix_path, "r") as zf:
            if "word/comments.xml" not in zf.namelist():
                raise RuntimeError("Document has no comments")
            c_root = ET.fromstring(zf.read("word/comments.xml"))
            comments = c_root.findall(f"{{{_W_NS}}}comment")
            idx = int(comment_index)
            if idx < 1 or idx > len(comments):
                raise RuntimeError(
                    f"comment_index {idx} out of range (1-{len(comments)})"
                )
            target_el = comments[idx - 1]
            target_para_id = _get_or_assign_comment_para_id(target_el)

            ex_root = _load_or_init_comments_ex(zf)
            done_str = "1" if resolve else "0"
            matched = None
            for el in ex_root.findall(f"{{{_W15_NS}}}commentEx"):
                if el.get(f"{{{_W15_NS}}}paraId") == target_para_id:
                    matched = el
                    break
            if matched is None:
                matched = ET.SubElement(ex_root, f"{{{_W15_NS}}}commentEx")
                matched.set(f"{{{_W15_NS}}}paraId", target_para_id)
            matched.set(f"{{{_W15_NS}}}done", done_str)

            updates["word/comments.xml"] = ET.tostring(
                c_root, encoding="utf-8", xml_declaration=True
            )
            updates["word/commentsExtended.xml"] = ET.tostring(
                ex_root, encoding="utf-8", xml_declaration=True
            )
            _ensure_comments_extended_wiring(zf, updates)

        _rewrite_docx_zip(posix_path, updates)
    finally:
        _reopen_in_word(posix_path)

    action_word = "Resolve" if resolve else "Unresolve"
    record_mac_undo(doc_name, f"{action_word} comment #{comment_index} (OOXML)")
    return json.dumps(
        {
            "success": True,
            "document": doc_name,
            "comment_index": int(comment_index),
            "resolved": bool(resolve),
            "method": "ooxml_hybrid",
        },
        ensure_ascii=False,
    )
