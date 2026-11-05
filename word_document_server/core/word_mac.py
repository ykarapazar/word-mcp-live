"""JXA (JavaScript for Automation) bridge for Microsoft Word on macOS.

Provides functions to interact with Word for Mac via osascript.
Each function is self-contained: builds a JXA script, executes it,
and returns JSON-compatible results matching the Windows COM tool outputs.

Only works on macOS with Microsoft Word installed.
"""

import json
import os
import re
import subprocess
import sys
import unicodedata
from contextlib import contextmanager


_TIMEOUT_MSG = (
    "Microsoft Word did not respond within {timeout} s. Word may be showing a dialog "
    "(close it and retry) or still be busy with a previous request, which it keeps "
    "processing after a timeout."
)


def _run_jxa(script: str, timeout: int = 30) -> str:
    """Execute a JXA script via osascript and return stdout.

    Args:
        script: JavaScript for Automation code string.
        timeout: Max seconds to wait.

    Returns:
        stdout as string (typically JSON from JSON.stringify).

    Raises:
        RuntimeError on timeout or execution error.
    """
    try:
        result = subprocess.run(
            ["/usr/bin/osascript", "-l", "JavaScript"],
            input=script,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(_TIMEOUT_MSG.format(timeout=timeout)) from None
    if result.returncode != 0:
        stderr = result.stderr.strip()
        # Provide cleaner error messages
        if "is not running" in stderr or "Connection is invalid" in stderr:
            raise RuntimeError("Microsoft Word is not running. Please open Word first.")
        raise RuntimeError(f"JXA error: {stderr}")
    return result.stdout.strip()


def _run_applescript(script: str, timeout: int = 30) -> str:
    """Execute an AppleScript and return stdout.

    Some Word for Mac features (e.g., make new Word comment) only work
    via AppleScript, not JXA. This is the fallback for those cases.
    """
    try:
        result = subprocess.run(
            ["/usr/bin/osascript"],
            input=script,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(_TIMEOUT_MSG.format(timeout=timeout)) from None
    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise RuntimeError(f"AppleScript error: {stderr}")
    return result.stdout.strip()


def _escape_as(s: str) -> str:
    """Escape a Python string for safe embedding in AppleScript."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _color_to_word_int(hex_color: str) -> int:
    """Convert #RRGGBB hex to Word BGR integer."""
    c = hex_color.lstrip("#")
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return r + (g * 256) + (b * 65536)


def _color_to_mac_rgb(hex_color: str) -> str:
    """Convert #RRGGBB hex to Mac Word 16-bit RGB list string '[R, G, B]'.
    Mac Word AppleScript/JXA requires RGB as 3-element list of 16-bit ints (0-65535).
    Single-integer color assignment is silently ignored on Mac."""
    c = hex_color.lstrip("#")
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return f"[{r * 257}, {g * 257}, {b * 257}]"


def _escape_js(s: str) -> str:
    """Escape a Python string for safe embedding in JavaScript."""
    return (
        s.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", "\\n")
        .replace("\r", "\\r")
        .replace("\t", "\\t")
    )


def _normalize_breaks(s: str) -> str:
    """Turn literal backslash sequences (\\r\\n, \\r, \\n — as they arrive via
    double-escaped JSON) into real carriage returns, Word's paragraph mark.
    Real CR/LF/CRLF are left alone: Word for Mac already treats each of them
    as a paragraph break when written via content."""
    return s.replace("\\r\\n", "\r").replace("\\r", "\r").replace("\\n", "\r")


def _escape_js_text(s: str) -> str:
    """Escape document text for a JS string literal, normalising paragraph
    breaks first (see _normalize_breaks). Use for any text written into the
    document body (insert_text, table cells) so "\\n" never lands literally."""
    return _escape_js(_normalize_breaks(s))


def _doc_finder_js(filename: str = None) -> str:
    """Return JXA code snippet that sets `d` to the target document.

    If filename is None/empty, uses documents[0] (front document).
    Otherwise, matches by basename (NFC-normalized, case-insensitive).
    """
    if not filename:
        return """
var d = app.documents[0];
if (!d) throw new Error("No documents are open in Word");
"""
    basename = unicodedata.normalize("NFC", os.path.basename(filename)).lower()
    return f"""
var d = null;
var target = "{_escape_js(basename)}";
for (var _i = 0; _i < app.documents.length; _i++) {{
    var _name = app.documents[_i].name().normalize("NFC").toLowerCase();
    if (_name === target) {{ d = app.documents[_i]; break; }}
}}
if (!d) {{
    var _open = [];
    for (var _j = 0; _j < app.documents.length; _j++) _open.push(app.documents[_j].name());
    throw new Error("Document '" + target + "' not open. Open: " + _open.join(", "));
}}
// Word for Mac returns ranges bound to documents.byName(), so two open
// documents with the same name cannot be told apart reliably.
var _same = app.documents.name().filter(function (n) {{ return n.normalize("NFC").toLowerCase() === target; }}).length;
if (_same > 1) throw new Error("Two or more open documents are named '" + d.name() + "' (different folders). Word for Mac cannot address them separately; close all but one.");
"""


# app.selection belongs to the active window, not to `d`. Selection-based
# code activates `d` first (same pattern as modify_table add_row) and
# restores the previously active document afterwards.
_ACTIVATE_TARGET_JS = """
var _prevName = app.activeDocument.name();
app.activateObject(d);
if (app.activeDocument.name() !== d.name())
    throw new Error("Could not activate '" + d.name() + "'");
"""

_RESTORE_ACTIVE_JS = """
if (_prevName !== d.name()) {
    for (var _di = 0; _di < app.documents.length; _di++) {
        if (app.documents[_di].name() === _prevName) { try { app.activateObject(app.documents[_di]); } catch (e) {} break; }
    }
}
"""


# ── Core functions (matching word_com.py interface) ──────────────────────


def get_word_app():
    """Verify Word for Mac is running and has documents.

    Returns a lightweight sentinel dict (not a COM object).
    On Mac, each JXA call is a separate osascript process,
    so we don't hold persistent references.
    """
    if sys.platform != "darwin":
        raise RuntimeError("word_mac is only available on macOS")

    result = _run_jxa("""
var app = Application("Microsoft Word");
JSON.stringify({
    version: app.version(),
    docCount: app.documents.length
});
""")
    info = json.loads(result)
    if info["docCount"] == 0:
        raise RuntimeError("No documents are open in Word for Mac")
    return {"platform": "darwin", "version": info["version"]}


def find_document(app_ref, filename: str = None):
    """Find an open document by name. Returns a dict with doc info.

    Args:
        app_ref: Sentinel from get_word_app() (unused, kept for interface parity).
        filename: Document basename or full path. None = front document.

    Returns:
        Dict with name, path keys.
    """
    finder = _doc_finder_js(filename)
    result = _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
JSON.stringify({{name: d.name(), path: d.posixFullName()}});
""")
    return json.loads(result)


@contextmanager
def undo_record(app_ref, name: str):
    """No-op context manager on macOS.

    Word for Mac's UndoRecord is not accessible via AppleScript/JXA.
    Each operation becomes a separate undo entry.
    """
    yield


# ── List / Info ──────────────────────────────────────────────────────────


def mac_list_open() -> str:
    """List all documents currently open in Word for Mac."""
    return _run_jxa("""
var app = Application("Microsoft Word");
var docs = [];
for (var i = 0; i < app.documents.length; i++) {
    var d = app.documents[i];
    docs.push({name: d.name(), path: d.posixFullName()});
}
JSON.stringify({documents: docs, count: docs.length});
""")


def mac_get_info(filename: str = None) -> str:
    """Get document metadata."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var pages = app.getRangeInformation(d.textObject, {{informationType: "number of pages in document"}});
JSON.stringify({{
    name: d.name(),
    path: d.posixFullName(),
    saved: d.saved(),
    track_revisions: d.trackRevisions(),
    pages: parseInt(pages),
    version: app.version()
}});
""")


def mac_save_as_pdf(filename: str = None, output_path: str = None) -> str:
    """Save document as PDF without closing it."""
    finder = _doc_finder_js(filename)
    if not output_path:
        return json.dumps({"error": "output_path required"})
    escaped_path = _escape_js(output_path)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.save();
app.saveAs(d, {{fileName: "{escaped_path}", fileFormat: "format PDF"}});
JSON.stringify({{converted: true, path: "{escaped_path}"}});
""")


def mac_save(filename: str = None, save_as: str = None) -> str:
    """Save the document."""
    finder = _doc_finder_js(filename)
    if save_as:
        escaped_path = _escape_js(save_as)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
app.saveAs(d, {{fileName: "{escaped_path}"}});
JSON.stringify({{saved: true, path: "{escaped_path}"}});
""")
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.save();
JSON.stringify({{saved: true, name: d.name()}});
""")


def mac_undo(filename: str = None, times: int = 1) -> str:
    """Undo N times."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var results = [];
for (var i = 0; i < {times}; i++) {{
    results.push(d.undo());
}}
JSON.stringify({{undone: {times}, results: results}});
""")


# ── Read ─────────────────────────────────────────────────────────────────


def mac_get_text(filename: str = None, timeout: int = 120) -> str:
    """Get all paragraph text from document.

    Reads the entire document text as a single string (one AppleEvent call)
    then splits by \\r in Python to recover paragraphs. This is ~10-100x
    faster than iterating paragraphs[i].textObject.content() per paragraph
    when Word/OneDrive is slow.
    """
    finder = _doc_finder_js(filename)
    bulk_result = _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
JSON.stringify({{text: d.textObject.content()}});
""", timeout=timeout)
    bulk_text = json.loads(bulk_result).get("text", "")
    # Word paragraphs are separated by \r in d.textObject.content()
    parts = bulk_text.split("\r")
    # Word's paragraphs collection includes a trailing empty paragraph; drop the
    # last split entry if it's empty (would create a phantom paragraph at the end).
    if parts and parts[-1] == "":
        parts = parts[:-1]
    paras = [{"index": i, "text": text + "\r"} for i, text in enumerate(parts)]
    return json.dumps({"paragraphs": paras, "count": len(paras)}, ensure_ascii=False)


def mac_get_page_text(filename: str = None, page: int = 1, end_page: int = None) -> str:
    """Get text from a specific page range.

    Uses binary search with createRange + getRangeInformation to find page
    boundaries efficiently (O(log N) IPC calls per boundary). Returns one
    entry per page with text content and char offsets.
    """
    finder = _doc_finder_js(filename)
    ep = end_page or page
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var totalEnd = d.textObject.endOfContent();
var totalPages = parseInt(app.getRangeInformation(d.textObject, {{informationType: "number of pages in document"}}));
if ({page} > totalPages) throw new Error("Page {page} exceeds document (" + totalPages + " pages)");
var ep = Math.min({ep}, totalPages);

// Page starts by hopping page to page ("go to next" page, one Apple Event
// each). A binary search on "active end page number" needed ~20 calls per
// boundary, each forcing layout, and timed out on long documents.
// Word binds the range "go to next" returns to the ACTIVE document, so d is
// activated while hopping (and the previous document restored after), and
// every hop starts from a fresh range on d.
var pageStarts = [0, 0];
function findPageStart(pg) {{
    if (pg <= 1) return 0;
    if (pg > totalPages) return totalEnd;
    while (pageStarts.length <= pg) {{
        var prev = pageStarts[pageStarts.length - 1];
        var r = d.createRange({{start: prev, end: prev}});
        pageStarts.push(app.goToNext(r, {{what: "goto a page item"}}).startOfContent());
    }}
    return pageStarts[pg];
}}

var pages = [];
{_ACTIVATE_TARGET_JS}
try {{
    for (var pg = {page}; pg <= ep; pg++) {{
        var pStart = findPageStart(pg);
        var pEnd = (pg >= totalPages) ? totalEnd : findPageStart(pg + 1);
        var pageRange = app.createRange(d, {{start: pStart, end: pEnd}});
        pages.push({{
            page: pg,
            text: pageRange.content(),
            char_start: pStart,
            char_end: pEnd
        }});
    }}
}} finally {{
{_RESTORE_ACTIVE_JS}
}}
JSON.stringify({{pages: pages, count: pages.length, page: {page}, end_page: ep, total_pages: totalPages}});
""", timeout=60)


def mac_find_text(
    filename: str = None,
    search_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    context_chars: int = 50,
    max_results: int = 20,
) -> str:
    """Find text in document using selection-based search."""
    finder = _doc_finder_js(filename)
    escaped_search = _escape_js(search_text)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{_ACTIVATE_TARGET_JS}
var results = [];
try {{
    var sel = app.selection;
    sel.selectionStart = 0;
    sel.selectionEnd = 0;
    var f = sel.findObject;
    f.clearFormatting();
    f.replacement.clearFormatting();
    f.replacement.content = "";
    // Find settings persist across executeFind calls; set them once
    f.content = "{escaped_search}";
    f.forward = true;
    f.wrap = "find stop";
    f.matchCase = {"true" if match_case else "false"};
    f.matchWholeWord = {"true" if whole_word else "false"};
    f.matchWildcards = {"true" if use_wildcards else "false"};
    var docEnd = d.textObject.endOfContent();
    for (var i = 0; i < {max_results}; i++) {{
        if (!f.executeFind()) break;
        var s = sel.selectionStart();
        var e = sel.selectionEnd();
        var ctxStart = Math.max(0, s - {context_chars});
        var ctxEnd = Math.min(docEnd, e + {context_chars});
        var ctxRange = d.createRange({{start: ctxStart, end: ctxEnd}});
        results.push({{
            text: sel.content(),
            start: s,
            end: e,
            context: ctxRange.content()
        }});
        // Move past this match
        sel.selectionStart = e;
    }}
}} finally {{
{_RESTORE_ACTIVE_JS}
}}
JSON.stringify({{matches: results, count: results.length, searchText: "{escaped_search}"}});
""")


def mac_get_paragraph_format(
    filename: str = None,
    start_paragraph: int = 0,
    end_paragraph: int = None,
    include_runs: bool = False,
) -> str:
    """Get formatting details for paragraph range."""
    finder = _doc_finder_js(filename)
    ep = f"{end_paragraph}" if end_paragraph is not None else "null"
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
// d.paragraphs() resolves every paragraph (tens of seconds on long documents);
// address only the requested ones by index instead.
var paraCount = d.paragraphs.length;
var startP = {start_paragraph};
var endP = {ep} !== null ? {ep} : startP;
endP = Math.min(endP, paraCount - 1);
var results = [];
for (var i = startP; i <= endP; i++) {{
    var p = d.paragraphs[i];
    var pf = p.paragraphFormat;
    var fo = p.textObject.fontObject.properties();  // one Apple Event instead of four
    var info = {{
        index: i,
        text: p.textObject.content(),
        style: null,
        alignment: pf.alignment(),
        spaceBefore: pf.spaceBefore(),
        spaceAfter: pf.spaceAfter(),
        lineSpacing: pf.lineSpacing(),
        keepWithNext: pf.keepWithNext(),
        keepTogether: pf.keepTogether(),
        pageBreakBefore: pf.pageBreakBefore(),
        fontName: fo.name,
        fontSize: fo.fontSize,
        bold: fo.bold,
        italic: fo.italic
    }};
    try {{ info.style = p.style(); }} catch(e) {{}}
    results.push(info);
}}
JSON.stringify({{paragraphs: results}});
""")


def mac_diagnose_layout(filename: str = None) -> str:
    """Diagnose layout issues (keep_with_next chains, etc.)."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
// d.paragraphs() resolves every paragraph (tens of seconds on long documents);
// address them by index and read text only for paragraphs that are reported.
var paraCount = d.paragraphs.length;
var issues = [];
var kwnChain = [];
for (var i = 0; i < paraCount; i++) {{
    var p = d.paragraphs[i];
    var pf = p.paragraphFormat;
    var kwn = pf.keepWithNext();
    var pbb = pf.pageBreakBefore();
    var text = null;
    if (kwn) {{
        kwnChain.push(i);
    }} else if (kwnChain.length > 2) {{
        text = p.textObject.content().substring(0, 60);
        issues.push({{type: "keep_with_next_chain", paragraphs: kwnChain.slice(), length: kwnChain.length, firstText: text}});
        kwnChain = [];
    }} else {{
        kwnChain = [];
    }}
    if (pbb) {{
        if (text === null) text = p.textObject.content().substring(0, 60);
        issues.push({{type: "page_break_before", paragraph: i, text: text}});
    }}
}}
if (kwnChain.length > 2) {{
    issues.push({{type: "keep_with_next_chain", paragraphs: kwnChain, length: kwnChain.length}});
}}
JSON.stringify({{issues: issues, totalParagraphs: paraCount}});
""")


# ── Edit ─────────────────────────────────────────────────────────────────


def mac_insert_text(
    filename: str = None,
    text: str = "",
    position: str = "end",
    bookmark: str = None,
    track_changes: bool = False,
) -> str:
    """Insert text into document."""
    finder = _doc_finder_js(filename)
    escaped = _escape_js_text(text)

    bookmark_js = ""
    if bookmark:
        bookmark_js = f"""
    var bm = d.bookmarks["{_escape_js(bookmark)}"];
    if (!bm) throw new Error("Bookmark '{_escape_js(bookmark)}' not found");
    var r = bm.bookmarkRange;
    r.startOfContent = r.endOfContent();
    r.content = text;
"""
    elif position == "start":
        bookmark_js = """
    var r = d.createRange({start: 0, end: 0});
    r.content = text;
"""
    elif position == "end":
        bookmark_js = """
    var endPos = d.textObject.endOfContent() - 1;
    var r = d.createRange({start: endPos, end: endPos});
    r.content = text;
"""
    elif position == "cursor":
        # The cursor of `d`, not of whichever document is in front
        bookmark_js = _ACTIVATE_TARGET_JS + """
    try {
        var sel = app.selection;
        sel.content = text;
    } finally {
""" + _RESTORE_ACTIVE_JS + """
    }
"""
    else:
        # Numeric position — coerce in Python so nothing unescaped reaches JS
        try:
            pos = int(str(position).strip())
        except ValueError:
            return json.dumps({"error": f"Invalid position: {position!r} (use 'start', 'end', 'cursor' or a character offset)"})
        bookmark_js = f"""
    var pos = {pos};
    var r = d.createRange({{start: pos, end: pos}});
    r.content = text;
"""

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var text = "{escaped}";
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    {bookmark_js}
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{inserted: true, length: text.length}});
""")


def mac_delete_text(
    filename: str = None,
    start: int = 0,
    end: int = 0,
    track_changes: bool = False,
) -> str:
    """Delete text range."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    var r = d.createRange({{start: {start}, end: {end}}});
    var deleted = r.content();
    r.content = "";
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{deleted: true, text: deleted, start: {start}, end: {end}}});
""")


def mac_replace_text(
    filename: str = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    match_whole_word: bool = False,
    use_wildcards: bool = False,
    replace_all: bool = True,
    track_changes: bool = False,
) -> str:
    """Find and replace text."""
    finder = _doc_finder_js(filename)
    escaped_find = _escape_js(find_text)
    escaped_replace = _escape_js(replace_text)
    replace_mode = "replace all" if replace_all else "replace one"
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{_ACTIVATE_TARGET_JS}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    var sel = app.selection;
    sel.selectionStart = 0;
    sel.selectionEnd = 0;
    var f = sel.findObject;
    f.clearFormatting();
    f.replacement.clearFormatting();
    f.content = "{escaped_find}";
    f.replacement.content = "{escaped_replace}";
    f.forward = true;
    f.wrap = "find continue";
    f.matchCase = {"true" if match_case else "false"};
    f.matchWholeWord = {"true" if match_whole_word else "false"};
    f.matchWildcards = {"true" if use_wildcards else "false"};
    var result = f.executeFind({{replace: "{replace_mode}"}});
}} finally {{
    d.trackRevisions = prevTracking;
{_RESTORE_ACTIVE_JS}
}}
JSON.stringify({{replaced: result, find: "{escaped_find}", replaceWith: "{escaped_replace}"}});
""")


def mac_format_text(
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
    highlight_color: str = None,
    style_name: str = None,
    paragraph_alignment: str = None,
    page_break_before: bool = None,
    preserve_direct_formatting: bool = False,
    track_changes: bool = False,
) -> str:
    """Format text range or paragraph range."""
    finder = _doc_finder_js(filename)

    # Build range selection JS
    if start is not None and end is not None:
        range_js = f"var r = d.createRange({{start: {start}, end: {end}}});"
    elif start_paragraph is not None:
        ep = end_paragraph if end_paragraph is not None else start_paragraph
        range_js = f"""
var startP = d.paragraphs[{start_paragraph}].textObject.startOfContent();
var endP = d.paragraphs[{ep}].textObject.endOfContent();
var r = d.createRange({{start: startP, end: endP}});
"""
    else:
        return json.dumps({"error": "Must provide start/end or start_paragraph"})

    # Build formatting JS
    fmt_lines = []
    if bold is not None:
        fmt_lines.append(f"r.bold = {'true' if bold else 'false'};")
    if italic is not None:
        fmt_lines.append(f"r.italic = {'true' if italic else 'false'};")
    if underline is not None:
        fmt_lines.append(f"r.underline = {'true' if underline else 'false'};")
    if strikethrough is not None:
        fmt_lines.append(f"r.fontObject.strikeThrough = {'true' if strikethrough else 'false'};")
    if font_name:
        fmt_lines.append(f'r.fontObject.name = "{_escape_js(font_name)}";')
    if font_size is not None:
        fmt_lines.append(f"r.fontObject.fontSize = {font_size};")
    if font_color:
        fmt_lines.append(f'r.fontObject.color = {_color_to_mac_rgb(font_color)};')
    if highlight_color is not None:
        # Mac WdColorIndex enum strings; invalid values silently CLEAR the highlight, so validate here
        _hl_names = ["no highlight", "black", "blue", "turquoise", "bright green", "pink", "red",
                     "yellow", "white", "dark blue", "teal", "green", "violet", "dark red",
                     "dark yellow", "gray50", "gray25"]
        _hl = str(highlight_color).strip().lower()
        if _hl.isdigit() and int(_hl) < len(_hl_names):
            _hl = _hl_names[int(_hl)]
        elif _hl in ("none", "no", "auto"):
            _hl = "no highlight"
        if _hl not in _hl_names:
            return json.dumps({"error": f"Invalid highlight_color: {highlight_color!r} (use 0-16)"})
        fmt_lines.append(f'r.highlightColorIndex = "{_hl}";')
    if style_name:
        fmt_lines.append(f'r.style = d.wordStyles["{_escape_js(style_name.lower())}"];')
    if paragraph_alignment:
        _align_map = {"left": "align paragraph left", "center": "align paragraph center",
                      "right": "align paragraph right", "justify": "align paragraph justify",
                      "0": "align paragraph left", "1": "align paragraph center",
                      "2": "align paragraph right", "3": "align paragraph justify"}
        _align_val = _align_map.get(paragraph_alignment.lower().strip(), paragraph_alignment)
        fmt_lines.append(f'r.paragraphFormat.alignment = "{_escape_js(_align_val)}";')
    if page_break_before is not None:
        fmt_lines.append(f"r.paragraphFormat.pageBreakBefore = {'true' if page_break_before else 'false'};")
    fmt_js = "\n    ".join(fmt_lines)

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    {range_js}
    {fmt_js}
    // Prevent bold bleed: unbold paragraph marks within the range
    {"" if bold is not True else """
    var rStart = r.startOfContent();
    var rEnd = r.endOfContent();
    var txt = r.content();
    for (var pi = 0; pi < txt.length; pi++) {
        if (txt[pi] === '\\r') {
            var pm = d.createRange({start: rStart + pi, end: rStart + pi + 1});
            pm.bold = false;
        }
    }
    """}
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{formatted: true}});
""")


def mac_toggle_track_changes(filename: str = None, enable: bool = True) -> str:
    """Toggle track changes on/off."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.trackRevisions = {"true" if enable else "false"};
JSON.stringify({{trackRevisions: d.trackRevisions()}});
""")


# ── Comments ─────────────────────────────────────────────────────────────


def mac_get_comments(filename: str = None) -> str:
    """Get all comments from document."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var comments = d.wordComments() || [];
var result = [];
for (var i = 0; i < comments.length; i++) {{
    var c = comments[i];
    result.push({{
        index: i,
        author: c.author(),
        text: c.commentText.content(),
        scope: c.scope.content(),
        date: c.dateValue().toString()
    }});
}}
JSON.stringify({{comments: result, count: result.length}});
""")


def mac_add_comment(
    filename: str = None,
    start: int = None,
    end: int = None,
    paragraph_index: int = None,
    text: str = "",
    author: str = None,
) -> str:
    """Add a comment to a text range."""
    finder = _doc_finder_js(filename)
    escaped_text = _escape_js(text)

    if start is not None and end is not None:
        range_js = f"var r = d.createRange({{start: {start}, end: {end}}});"
    elif paragraph_index is not None:
        range_js = f"""
var p = d.paragraphs[{paragraph_index}];
var r = p.textObject;
"""
    else:
        return json.dumps({"error": "Must provide start/end or paragraph_index"})

    # Resolve the target document by name (same matching as every other tool)
    doc_name = json.loads(_run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
JSON.stringify(d.name());
"""))

    # JXA's make() doesn't work for Word comments — use AppleScript
    escaped_as_text = _escape_as(text)
    if start is not None and end is not None:
        range_as = f"set r to create range d start {start} end {end}"
    elif paragraph_index is not None:
        range_as = f"set r to text object of paragraph {paragraph_index + 1} of d"
    else:
        return json.dumps({"error": "Must provide start/end or paragraph_index"})

    # Word for Mac ignores "user name" for new comments (it uses the signed-in
    # identity), so set the author on the created comment instead.
    author_as = f'set author of c to "{_escape_as(author)}"' if author else ""

    result = _run_applescript(f'''
tell application "Microsoft Word"
    set d to document "{_escape_as(doc_name)}"
    {range_as}
    set c to make new Word comment at d with properties {{comment text:"{escaped_as_text}", scope:r}}
    {author_as}
    return count of Word comments of d
end tell
''')
    return json.dumps({"added": True, "commentCount": int(result)})


def mac_delete_comment(filename: str = None, comment_index: int = 0) -> str:
    """Delete a comment by index."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
if ({comment_index} >= d.wordComments.length) throw new Error("Comment index out of range");
app.delete(d.wordComments[{comment_index}]);
JSON.stringify({{deleted: true, remaining: d.wordComments.length}});
""")


# ── Revisions ────────────────────────────────────────────────────────────


def mac_list_revisions(filename: str = None) -> str:
    """List all tracked changes."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var revs = d.revisions() || [];
var result = [];
for (var i = 0; i < Math.min(revs.length, 200); i++) {{
    result.push({{
        index: i,
        author: revs[i].author(),
        type: revs[i].revisionType(),
        date: revs[i].dateValue().toString()
    }});
}}
JSON.stringify({{revisions: result, count: revs.length}});
""")


def mac_accept_revisions(filename: str = None, author: str = None, revision_ids: list = None) -> str:
    """Accept tracked changes."""
    finder = _doc_finder_js(filename)
    if revision_ids:
        ids_js = json.dumps(revision_ids)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var ids = {ids_js};
var accepted = 0;
// Accept in reverse order to preserve indices
for (var i = ids.length - 1; i >= 0; i--) {{
    app.accept(d.revisions[ids[i]]);
    accepted++;
}}
JSON.stringify({{accepted: accepted}});
""")
    elif author:
        escaped_author = _escape_js(author)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var revs = d.revisions();
var accepted = 0;
for (var i = revs.length - 1; i >= 0; i--) {{
    if (revs[i].author() === "{escaped_author}") {{
        app.accept(revs[i]);
        accepted++;
    }}
}}
JSON.stringify({{accepted: accepted}});
""")
    else:
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
app.acceptAllRevisions(d);
JSON.stringify({{accepted: "all"}});
""")


def mac_reject_revisions(filename: str = None, author: str = None, revision_ids: list = None) -> str:
    """Reject tracked changes."""
    finder = _doc_finder_js(filename)
    if revision_ids:
        ids_js = json.dumps(revision_ids)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var ids = {ids_js};
var rejected = 0;
for (var i = ids.length - 1; i >= 0; i--) {{
    app.reject(d.revisions[ids[i]]);
    rejected++;
}}
JSON.stringify({{rejected: rejected}});
""")
    elif author:
        escaped_author = _escape_js(author)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var revs = d.revisions();
var rejected = 0;
for (var i = revs.length - 1; i >= 0; i--) {{
    if (revs[i].author() === "{escaped_author}") {{
        app.reject(revs[i]);
        rejected++;
    }}
}}
JSON.stringify({{rejected: rejected}});
""")
    else:
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
app.rejectAllRevisions(d);
JSON.stringify({{rejected: "all"}});
""")


# ── Layout ───────────────────────────────────────────────────────────────


def mac_set_page_layout(
    filename: str = None,
    section_index: int = 0,
    orientation: str = None,
    page_width: float = None,
    page_height: float = None,
    top_margin: float = None,
    bottom_margin: float = None,
    left_margin: float = None,
    right_margin: float = None,
) -> str:
    """Set page layout for a section."""
    finder = _doc_finder_js(filename)
    props = []
    if orientation:
        props.append(f'ps.orientation = "orient {orientation}";')
    if page_width is not None:
        props.append(f"ps.pageWidth = {page_width};")
    if page_height is not None:
        props.append(f"ps.pageHeight = {page_height};")
    if top_margin is not None:
        props.append(f"ps.topMargin = {top_margin};")
    if bottom_margin is not None:
        props.append(f"ps.bottomMargin = {bottom_margin};")
    if left_margin is not None:
        props.append(f"ps.leftMargin = {left_margin};")
    if right_margin is not None:
        props.append(f"ps.rightMargin = {right_margin};")
    props_js = "\n    ".join(props)

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var secIdx = Math.max(0, {section_index} - 1);  // Tool sends 1-based, JXA is 0-based
var ps = d.sections[secIdx].pageSetup;
{props_js}
JSON.stringify({{
    orientation: ps.orientation(),
    topMargin: ps.topMargin(),
    bottomMargin: ps.bottomMargin(),
    leftMargin: ps.leftMargin(),
    rightMargin: ps.rightMargin()
}});
""")


def mac_add_header_footer(
    filename: str = None,
    section_index: int = 0,
    header_text: str = None,
    footer_text: str = None,
    alignment: str = None,
) -> str:
    """Add header and/or footer text."""
    finder = _doc_finder_js(filename)
    header_js = ""
    footer_js = ""
    if header_text is not None:
        header_js = f"""
    var h = app.getHeader(s, {{index: "header footer primary"}});
    h.textObject.content = "{_escape_js(header_text)}";
"""
    if footer_text is not None:
        footer_js = f"""
    var f = app.getFooter(s, {{index: "header footer primary"}});
    f.textObject.content = "{_escape_js(footer_text)}";
"""
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var s = d.sections[Math.max(0, {section_index} - 1)];  // 1-based → 0-based
{header_js}
{footer_js}
JSON.stringify({{added: true}});
""")


def mac_add_page_numbers(
    filename: str = None,
    section_index: int = 1,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    suffix: str = "",
    include_total: bool = False,
) -> str:
    """Add PAGE (and optionally NUMPAGES) fields to the primary header/footer.

    Mac gotcha: `create new field` REPLACES the text range it is given, so the
    header/footer text is written first with private-use placeholder
    characters (U+E000 / U+E001), then each placeholder's one-character range
    is turned into a field — rightmost first so earlier indices stay valid.
    """
    finder = _doc_finder_js(filename)
    getter = "getHeader" if (position or "").lower() == "header" else "getFooter"
    align_map = {"left": "align paragraph left", "center": "align paragraph center", "right": "align paragraph right"}
    align_val = align_map.get((alignment or "").lower(), "align paragraph center")
    prefix = prefix or ""
    suffix = suffix or ""
    mid = " / " if include_total else ""
    ph_page = "\\uE000"
    ph_total = "\\uE001" if include_total else ""
    text_js = '"' + _escape_js(prefix) + ph_page + _escape_js(mid) + ph_total + _escape_js(suffix) + '"'
    idx_page = len(prefix)
    idx_total = len(prefix) + 1 + len(mid)
    total_js = ""
    if include_total:
        total_js = f'app.createNewField({{textRange: hf.textObject.characters[{idx_total}], fieldType: "field num pages"}});'

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var secIdx = Math.max(0, {section_index} - 1);  // 1-based → 0-based
if (secIdx >= d.sections.length) throw new Error("Section {section_index} out of range (1-" + d.sections.length + ")");
var hf = app.{getter}(d.sections[secIdx], {{index: "header footer primary"}});
hf.textObject.content = {text_js};
{total_js}
app.createNewField({{textRange: hf.textObject.characters[{idx_page}], fieldType: "field page"}});
hf.textObject.paragraphFormat.alignment = "{align_val}";
JSON.stringify({{success: true, document: d.name(), section: secIdx + 1, position: {json.dumps(position)},
    alignment: {json.dumps(alignment)}, include_total: {"true" if include_total else "false"},
    fields: hf.textObject.fields.length, text: hf.textObject.content().replace(/\\r$/, "")}});
""")


def mac_add_section_break(filename: str = None, break_type: str = "section break next page") -> str:
    """Insert a section break."""
    finder = _doc_finder_js(filename)
    escaped_type = _escape_js(break_type)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var lastPara = d.paragraphs[d.paragraphs.length - 1];
app.insertBreak(lastPara.textObject, {{breakType: "{escaped_type}"}});
JSON.stringify({{sections: d.sections.length}});
""")


def mac_set_paragraph_spacing(
    filename: str = None,
    paragraph_index: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    space_before: float = None,
    space_after: float = None,
    line_spacing: float = None,
    keep_with_next: bool = None,
    keep_together: bool = None,
    alignment: str = None,
) -> str:
    """Set paragraph spacing and properties."""
    finder = _doc_finder_js(filename)
    start_p = paragraph_index if paragraph_index is not None else (start_paragraph or 0)
    end_p = end_paragraph if end_paragraph is not None else start_p

    props = []
    if space_before is not None:
        props.append(f"pf.spaceBefore = {space_before};")
    if space_after is not None:
        props.append(f"pf.spaceAfter = {space_after};")
    if line_spacing is not None:
        props.append(f"pf.lineSpacing = {line_spacing};")
    if keep_with_next is not None:
        props.append(f"pf.keepWithNext = {'true' if keep_with_next else 'false'};")
    if keep_together is not None:
        props.append(f"pf.keepTogether = {'true' if keep_together else 'false'};")
    if alignment:
        props.append(f'pf.alignment = "{_escape_js(alignment)}";')
    props_js = "\n        ".join(props)

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
for (var i = {start_p}; i <= Math.min({end_p}, d.paragraphs.length - 1); i++) {{
    var pf = d.paragraphs[i].paragraphFormat;
    {props_js}
}}
JSON.stringify({{updated: true, from: {start_p}, to: {end_p}}});
""")


def mac_add_bookmark(filename: str = None, paragraph_index: int = 0, bookmark_name: str = "Bookmark") -> str:
    """Create a named bookmark."""
    finder = _doc_finder_js(filename)
    escaped_name = _escape_js(bookmark_name)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var pIdx = Math.max(0, {paragraph_index} - 1);  // 1-based → 0-based
var r = d.paragraphs[pIdx].textObject;
app.make({{new: "bookmark", at: d, withProperties: {{name: "{escaped_name}", bookmarkRange: r}}}});
JSON.stringify({{added: true, name: "{escaped_name}"}});
""")


# ── Tables ───────────────────────────────────────────────────────────────


def mac_add_table(
    filename: str = None,
    rows: int = 3,
    cols: int = 3,
    position: str = "end",
    data: list = None,
    style: str = "Table Grid",
    autofit: str = "window",
    track_changes: bool = False,
) -> str:
    """Add a table to the document.

    Mac gotchas: the range handed to make() must come from d.createRange()
    (app.createRange(d, ...) raises -1708) and be passed as textObject in
    withProperties (otherwise Word ignores the row/column counts). The object
    make() returns is not usable with getCellFromTable, so the new table is
    re-fetched from d.tables by position before styling/filling it.
    """
    finder = _doc_finder_js(filename)
    if position == "start":
        pos_js = "var pos = 0;"
    elif position == "end":
        pos_js = "var pos = d.textObject.endOfContent() - 1;"
    else:
        try:
            pos_js = f"var pos = {int(position)};"
        except (TypeError, ValueError):
            return json.dumps({"error": f"Invalid position: {position}"})

    autofit_map = {"window": "auto fit window", "content": "auto fit content", "fixed": "auto fit fixed"}
    autofit_js = ""
    if autofit and autofit.lower() in autofit_map:
        autofit_js = f'app.autoFitBehavior(t, {{behavior: "{autofit_map[autofit.lower()]}"}});'
    # Mac: assigning the style NAME is silently ignored; assign the wordStyles object (readback is always null — verify via XML)
    style_js = f'try {{ t.style = d.wordStyles["{_escape_js(style)}"]; }} catch (e) {{}}' if style else ""

    data_js = ""
    if data:
        norm = [[_normalize_breaks(v) if isinstance(v, str) else v for v in r] for r in data]
        data_js = f"""
    var data = {json.dumps(norm)};
    for (var ri = 0; ri < Math.min(data.length, {rows}); ri++) {{
        for (var ci = 0; ci < Math.min(data[ri].length, {cols}); ci++) {{
            app.getCellFromTable(t, {{row: ri + 1, column: ci + 1}}).textObject.content = String(data[ri][ci]);
        }}
    }}
"""

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{pos_js}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
var t = null;
try {{
    var rng = d.createRange({{start: pos, end: pos}});
    app.make({{new: "table", at: d, withProperties: {{
        textObject: rng,
        numberOfRows: {rows},
        numberOfColumns: {cols}
    }}}});
    for (var ti = 0; ti < d.tables.length; ti++) {{  // re-fetch: first table at/after pos
        if (d.tables[ti].textObject.startOfContent() >= pos) {{ t = d.tables[ti]; break; }}
    }}
    if (!t) t = d.tables[d.tables.length - 1];
    {style_js}
    {autofit_js}
    {data_js}
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{success: true, added: true, document: d.name(), rows: t.rows.length, cols: t.columns.length,
    tables: d.tables.length, position: {json.dumps(str(position))},
    style: {json.dumps(style or None)}, autofit: {json.dumps(autofit or None)}, tracked: {"true" if track_changes else "false"}}});
""")


def mac_modify_table(
    filename: str = None,
    table_index: int = 0,
    operation: str = "get_info",
    row: int = None,
    col: int = None,
    text: str = None,
    before_row: int = None,
    cells: list = None,
    track_changes: bool = False,
) -> str:
    """Modify table structure or content (get_info, set_cell, set_row, add_row, delete_row).

    Mac gotchas: getCellFromTable does NOT bounds-check (row 10 of a 9-row
    table silently returns some other cell), so every operation validates
    indices first. Word for Mac has no make-new-row verb: add_row positions
    the app-level selection (selectionStart/End) inside the anchor row and
    calls `insert rows` on it, which only works when the target document is
    the active one (activateObject; the previous document is re-activated
    afterwards).
    """
    finder = _doc_finder_js(filename)
    op = (operation or "").lower()
    track_js = "true" if track_changes else "false"
    table_js = f"""
var tCount = d.tables.length;
if (tCount === 0) throw new Error("Document has no tables");
var tIdx = Math.max(0, {table_index} - 1);  // 1-based → 0-based
if (tIdx >= tCount) throw new Error("table_index {table_index} out of range (1-" + tCount + ")");
var t = d.tables[tIdx];
var nRows = t.rows.length;
var nCols = t.columns.length;
"""

    if op == "get_info":
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{table_js}
var cells = [];
for (var r = 1; r <= nRows; r++) {{
    for (var c = 1; c <= nCols; c++) {{
        var cell = app.getCellFromTable(t, {{row: r, column: c}});
        cells.push({{row: r, col: c, text: cell.textObject.content().replace(/[\\r\\x07]/g, "")}});
    }}
}}
JSON.stringify({{rows: nRows, cols: nCols, cells: cells, document: d.name(), table_index: tIdx + 1}});
""")

    elif op == "set_cell":
        if row is None or col is None or text is None:
            return json.dumps({"error": "set_cell requires row, col, and text"})
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{table_js}
if ({row} < 1 || {row} > nRows) throw new Error("Row {row} out of range (1-" + nRows + ")");
if ({col} < 1 || {col} > nCols) throw new Error("Col {col} out of range (1-" + nCols + ")");
var prevTracking = d.trackRevisions();
if ({track_js}) d.trackRevisions = true;
try {{
    app.getCellFromTable(t, {{row: {row}, column: {col}}}).textObject.content = "{_escape_js_text(text)}";
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{success: true, operation: "set_cell", row: {row}, col: {col},
    document: d.name(), table_index: tIdx + 1, tracked: {track_js}}});
""")

    elif op == "set_row":
        if row is None or not cells:
            return json.dumps({"error": "set_row requires row and cells (list of values)"})
        vals = [_normalize_breaks(v) if isinstance(v, str) else v for v in cells]
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{table_js}
if ({row} < 1 || {row} > nRows) throw new Error("Row {row} out of range (1-" + nRows + ")");
var vals = {json.dumps(vals)};
var updated = [];
var prevTracking = d.trackRevisions();
if ({track_js}) d.trackRevisions = true;
try {{
    for (var i = 0; i < vals.length && i < nCols; i++) {{
        if (vals[i] === null) continue;
        app.getCellFromTable(t, {{row: {row}, column: i + 1}}).textObject.content = String(vals[i]);
        updated.push(i + 1);
    }}
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{success: true, operation: "set_row", row: {row}, cells_updated: updated, total_cols: nCols,
    document: d.name(), table_index: tIdx + 1, tracked: {track_js}}});
""")

    elif op in ("add_row", "insert_row"):
        if op == "insert_row" and before_row is None:
            before_row = row
        vals = [_normalize_breaks(v) if isinstance(v, str) else v for v in (cells or [])]
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{table_js}
var beforeRow = {json.dumps(before_row)};
if (beforeRow !== null && (beforeRow < 1 || beforeRow > nRows + 1))
    throw new Error("before_row " + beforeRow + " out of range (1-" + (nRows + 1) + ")");
var append = (beforeRow === null || beforeRow > nRows);
var anchorRow = append ? nRows : beforeRow;
var newRow = append ? nRows + 1 : beforeRow;
var anchor = app.getCellFromTable(t, {{row: anchorRow, column: 1}}).textObject.startOfContent();
var prevName = app.activeDocument.name();
var prevTracking = d.trackRevisions();
if ({track_js}) d.trackRevisions = true;
try {{
    app.activateObject(d);
    if (app.activeDocument.name() !== d.name())
        throw new Error("Could not activate '" + d.name() + "' for row insertion");
    var sel = app.selection;
    sel.selectionStart = anchor;
    sel.selectionEnd = anchor;
    if (sel.selectionStart() !== anchor || sel.document.name() !== d.name())
        throw new Error("Selection could not be positioned in the target table");
    sel.insertRows({{position: append ? "below" : "above", numberOfRows: 1}});
    if (t.rows.length !== nRows + 1)
        throw new Error("Row insertion did not take effect (rows=" + t.rows.length + ")");
    var vals = {json.dumps(vals)};
    for (var i = 0; i < vals.length && i < nCols; i++) {{
        if (vals[i] === null) continue;
        app.getCellFromTable(t, {{row: newRow, column: i + 1}}).textObject.content = String(vals[i]);
    }}
}} finally {{
    d.trackRevisions = prevTracking;
    if (prevName !== d.name()) {{
        for (var di = 0; di < app.documents.length; di++) {{
            if (app.documents[di].name() === prevName) {{ try {{ app.activateObject(app.documents[di]); }} catch (e) {{}} break; }}
        }}
    }}
}}
JSON.stringify({{success: true, operation: "add_row", new_row: newRow, total_rows: t.rows.length,
    document: d.name(), table_index: tIdx + 1, tracked: {track_js}}});
""")

    elif op == "delete_row":
        if row is None:
            return json.dumps({"error": "delete_row requires row"})
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{table_js}
if ({row} < 1 || {row} > nRows) throw new Error("Row {row} out of range (1-" + nRows + ")");
var prevTracking = d.trackRevisions();
if ({track_js}) d.trackRevisions = true;
try {{
    app.delete(t.rows[{row} - 1]);
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{success: true, operation: "delete_row", deleted_row: {row}, remaining_rows: t.rows.length,
    document: d.name(), table_index: tIdx + 1, tracked: {track_js}}});
""")

    return json.dumps({"error": f"Unknown operation '{op}'. On macOS use: get_info, set_cell, set_row, add_row, delete_row"})


# ── Screen Capture ───────────────────────────────────────────────────────


def mac_screen_capture(filename: str = None, output_path: str = "/tmp/word_capture.png") -> str:
    """Capture the Word window on macOS."""
    # Activate Word
    _run_jxa("""
var app = Application("Microsoft Word");
app.activate();
""")

    import time
    time.sleep(0.5)

    # Use screencapture with the frontmost window
    result = subprocess.run(
        ["screencapture", "-x", "-o", "-w", output_path],
        capture_output=True,
        text=True,
        timeout=10,
    )
    if result.returncode != 0:
        # Fallback to full screen capture
        subprocess.run(
            ["screencapture", "-x", output_path],
            capture_output=True,
            text=True,
            timeout=10,
        )

    if os.path.exists(output_path):
        size = os.path.getsize(output_path)
        return json.dumps({"captured": True, "path": output_path, "size": size})
    return json.dumps({"error": "Screen capture failed"})


def mac_apply_list(
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
    font_color: str = None,
    outline_numbered: bool = True,
) -> str:
    """Apply list formatting to paragraphs on Mac using AppleScript.

    outline_numbered: when False, the list template is created without the
    outline-numbered flag. Use this for non-cascading per-section lists like
    (a)/(b)/(c) under MADDE 11/13/17/18 — outline-numbered templates would
    displace the document's single active outline list (Heading 1/MADDE).
    """
    _outline_js = "true" if outline_numbered else "false"
    if start_paragraph is None:
        return json.dumps({"error": "start_paragraph required"})
    ep = end_paragraph or start_paragraph

    finder = _doc_finder_js(filename)

    if remove:
        # Bulk removal: build ONE range spanning all target paragraphs, call
        # removeNumbers once. Avoids per-paragraph AppleEvent overhead.
        body_result = _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var paraCount = d.paragraphs.length;
var startIdx = {start_paragraph - 1};
var endIdx = Math.min({ep}, paraCount);
if (startIdx < 0) startIdx = 0;
var bodyCount = 0;
if (endIdx > startIdx) {{
    var startPara = d.paragraphs[startIdx];
    var endPara = d.paragraphs[endIdx - 1];
    var r = d.createRange({{
        start: startPara.textObject.startOfContent(),
        end: endPara.textObject.endOfContent()
    }});
    r.listFormat.removeNumbers();
    bodyCount = endIdx - startIdx;
}}
JSON.stringify({{removed: true, count: bodyCount}});
""", timeout=120)
        body_data = json.loads(body_result)

        # Clear list formatting from header/footer stories via VBA (do Visual Basic).
        # JXA's sec.headers[...] / sec.footers[...] accessors don't reliably return
        # HeaderFooter objects, but VBA's Sections collection works cleanly.
        # The VBA writes its iteration count into the clipboard so AppleScript
        # can read it back (Variables(...) round-trip requires quoting that
        # breaks AppleScript's `do Visual Basic` string literal).
        try:
            # VBA source (write a flag file so AppleScript doesn't have to
            # round-trip the count through a string-quoted Variables() call).
            vba_lines = [
                "Dim sec As Section",
                "Dim hf As HeaderFooter",
                "Dim n As Long: n = 0",
                "For Each sec In ActiveDocument.Sections",
                "    For Each hf In sec.Footers",
                "        hf.Range.ListFormat.RemoveNumbers",
                "        n = n + 1",
                "    Next",
                "    For Each hf In sec.Headers",
                "        hf.Range.ListFormat.RemoveNumbers",
                "        n = n + 1",
                "    Next",
                "Next",
                "Dim fileNum As Integer: fileNum = FreeFile",
                'Open "/tmp/_word_hf_cleared.txt" For Output As #fileNum',
                "Print #fileNum, CStr(n)",
                "Close #fileNum",
            ]
            # Build a single-line VBA program (do Visual Basic accepts colon
            # separators between statements). Then escape every literal " as ""
            # for AppleScript's nested string literal.
            vba_program = ": ".join(vba_lines)
            as_escaped = vba_program.replace('"', '""')
            applescript = (
                'tell application "Microsoft Word"\n'
                f'    do Visual Basic "{as_escaped}"\n'
                'end tell\n'
            )
            _run_applescript(applescript, timeout=30)
            try:
                with open("/tmp/_word_hf_cleared.txt") as f:
                    hf_cleared = int(f.read().strip())
                os.remove("/tmp/_word_hf_cleared.txt")
            except Exception:
                hf_cleared = -1
        except Exception as e:
            hf_cleared = 0
            body_data["headerFooterError"] = str(e)[:200]

        body_data["headerFooterCleared"] = hf_cleared
        return json.dumps(body_data)

    if list_type == "multilevel":
        nf = dict(number_format or {"1": "%1.", "2": "%1.%2."})
        ns = dict(number_style or {})
        sa = dict(start_at or {})
        lm = level_map or {}
        if "3" not in nf:
            nf["3"] = "(%3)"
        if "3" not in ns:
            ns["3"] = "lowercase_letter"
        style_map = {
            "arabic": "list number style arabic",
            "lowercase_letter": "list number style lowercase letter",
            "uppercase_letter": "list number style uppercase letter",
            "lowercase_roman": "list number style lowercase roman",
            "uppercase_roman": "list number style upper case roman",
        }
        levels_js = ""
        for lvl_str, fmt in nf.items():
            lvl = int(lvl_str)
            ns_val = style_map.get(str(ns.get(str(lvl), ns.get(lvl, "arabic"))), "list number style arabic")
            sa_val = sa.get(str(lvl), sa.get(lvl, 1))
            indent = 28 * (lvl - 1)
            text_indent = indent + 28 if lvl > 1 else 0
            # restartAfter: Level N restarts counter when level (N-1) increments.
            # Level 1 never restarts. Level 2 restarts after Level 1. Level 3 restarts after Level 2.
            restart_after = lvl - 1 if lvl > 1 else 0
            levels_js += f"""
    var lv{lvl} = lt.listLevels[{lvl - 1}];
    lv{lvl}.numberFormat = "{_escape_js(fmt)}";
    lv{lvl}.numberStyle = "{ns_val}";
    lv{lvl}.startAt = {sa_val};
    lv{lvl}.numberPosition = {indent};
    lv{lvl}.textPosition = {text_indent};
    try {{ lv{lvl}.restartAfter = {restart_after}; }} catch(e) {{}}
"""
        default_lvl = level + 1 if level > 0 else 1

        # Detect level_map format:
        # 1) {heading_text: level_number} — text-to-level mapping (keys non-numeric)
        # 2) {para_index: level_number} — numeric indices (keys numeric, values numeric)
        # 3) {index: heading_text_string} — legacy heading texts (keys numeric, values non-numeric)
        heading_level_map = None
        heading_texts = None
        lm_js = "{}"
        if lm:
            has_text_keys = any(not str(k).strip().isdigit() for k in lm.keys())
            if has_text_keys:
                heading_level_map = {}
                for k, v in lm.items():
                    try:
                        heading_level_map[str(k)] = int(v)
                    except (ValueError, TypeError):
                        heading_level_map[str(k)] = 1
            else:
                try:
                    lm_js = json.dumps({str(k): int(v) for k, v in lm.items()})
                except (ValueError, TypeError):
                    heading_texts = [str(v) for v in lm.values()]

        if heading_level_map:
            # Hybrid approach: do text matching in Python (avoids JXA AppleEvent
            # timeout on long iterations), then send numeric indices to a fast JXA.
            import re as _re
            paras_data = json.loads(mac_get_text(filename=filename)).get("paragraphs", [])

            _quote_chars = [chr(0x2018), chr(0x2019), chr(0x201c), chr(0x201d), chr(0x22)]
            def _norm_text(s: str) -> str:
                s = s.upper()
                for c in _quote_chars:
                    s = s.replace(c, "'")
                return _re.sub(r"\s+", " ", s).strip()

            num_prefix_re = _re.compile(r"^\s*\d{1,2}\.\d{1,2}[\.:]?[\s\t]*")
            letter_a_re = _re.compile(r"^\([a-zğüşıöç]\)[\s\t]")
            letter_b_re = _re.compile(r"^[a-z]\)[\s\t]")
            roman_chars = set("ivxlcdm")

            # Build normalized heading map (text → level)
            norm_map = {_norm_text(k): v for k, v in heading_level_map.items()}

            # Match paragraphs (text → level)
            indices_levels = []  # list of (para_idx, level)
            content_changes = {}  # {para_idx: new_content_without_prefix}
            first_h1_idx = -1
            for pdata in paras_data:
                idx = pdata.get("index", -1)
                raw = (pdata.get("text") or "").replace("\r", "").replace("\n", "")
                if not raw:
                    continue
                p_text = _norm_text(raw)
                p_text_no_num = _norm_text(num_prefix_re.sub("", raw))
                level = 0
                matched_key = None

                if p_text in norm_map:
                    level = norm_map[p_text]; matched_key = p_text
                elif p_text_no_num != p_text and p_text_no_num in norm_map:
                    level = norm_map[p_text_no_num]; matched_key = p_text_no_num
                else:
                    for key in list(norm_map.keys()):
                        if p_text.startswith(key) or (p_text_no_num != p_text and p_text_no_num.startswith(key)):
                            level = norm_map[key]; matched_key = key; break

                # Auto-detect Level 3 lettered items (only after first L1 heading)
                if level == 0 and first_h1_idx >= 0 and len(raw) > 10:
                    if letter_a_re.match(raw):
                        level = 3
                    elif letter_b_re.match(raw):
                        ch = raw[0]
                        if ch not in roman_chars:
                            level = 3

                if level > 0:
                    if matched_key:
                        del norm_map[matched_key]
                    if level >= 2 and num_prefix_re.match(raw):
                        content_changes[idx] = num_prefix_re.sub("", raw)
                    if level == 3:
                        stripped3 = _re.sub(r"^\(?[a-zğüşıöç]\)[\s\t]*", "", content_changes.get(idx, raw))
                        if stripped3 != content_changes.get(idx, raw):
                            content_changes[idx] = stripped3
                    indices_levels.append((idx, level))
                    if level == 1 and first_h1_idx < 0:
                        first_h1_idx = idx

            indices_levels_json = json.dumps(indices_levels)
            content_changes_json = json.dumps(
                {str(k): v for k, v in content_changes.items()},
                ensure_ascii=False,
            )
            first_h1_js = first_h1_idx if first_h1_idx >= 0 else 0
            color_block = ""
            if font_color:
                color_block = f"""
var fc = {_color_to_mac_rgb(font_color)};
for (var ci = 0; ci < indices.length; ci++) {{
    try {{ paras[indices[ci][0]].textObject.fontObject.color = fc; }} catch(e) {{}}
}}
"""
            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var lt = app.make({{new: "listTemplate", at: d, withProperties: {{outlineNumbered: {_outline_js}}}}});
{levels_js}
var indices = {indices_levels_json};
var contentChanges = {content_changes_json};
var paras = d.paragraphs();
var counts = [0, 0, 0, 0];

for (var ii = 0; ii < indices.length; ii++) {{
    var idx = indices[ii][0];
    var lvl = indices[ii][1];
    if (idx < 0 || idx >= paras.length) continue;
    var p = paras[idx];
    var newContent = contentChanges[String(idx)];
    if (newContent !== undefined) {{
        p.textObject.content = newContent + "\\r";
    }}
    var r = d.createRange({{start: p.textObject.startOfContent(), end: p.textObject.endOfContent()}});
    // applyListFormatTemplate: only the very first application starts a new list;
    // subsequent applications continue the existing list so Level 1 counters
    // increment across separately-applied paragraphs.
    r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: (ii > 0)}});
    // Only set explicit listLevelNumber for L2/L3. Setting it to 1 on L1
    // paragraphs resets the L1 counter (causes "MADDE 1" twice instead of
    // "MADDE 1" + "MADDE 2"). Word_com.py uses the same pattern.
    if (lvl > 1) r.listFormat.listLevelNumber = lvl;
    counts[lvl]++;
}}
{color_block}
JSON.stringify({{applied: true, type: "multilevel", h1: counts[1], h2: counts[2], h3: counts[3], total: indices.length}});
""", timeout=600)

        elif heading_texts:
            texts_json = json.dumps(heading_texts)
            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var lt = app.make({{new: "listTemplate", at: d, withProperties: {{outlineNumbered: {_outline_js}}}}});
{levels_js}
var headingTexts = {texts_json};
var headingSet = {{}};
var norm = function(s) {{ return s.toUpperCase().replace(/[\\u2018\\u2019\\u201C\\u201D\\u0027]/g, "\\u0027"); }};
for (var hi = 0; hi < headingTexts.length; hi++) headingSet[norm(headingTexts[hi])] = true;
var paras = d.paragraphs();
var h1Applied = 0;
var h2Applied = 0;
var subArticleRe = /^\\d{{1,2}}\\.\\d{{1,2}}[\\.\\s]/;
var firstH1 = -1;
for (var i = 0; i < paras.length; i++) {{
    var raw = paras[i].textObject.content().replace(/[\\r\\n]/g, "");
    var pText = norm(raw);
    var matched = headingSet[pText];
    if (!matched) {{ for (var key in headingSet) {{ if (pText.indexOf(key) === 0) {{ matched = true; pText = key; break; }} }} }}
    if (matched) {{
        var r = d.createRange({{start: paras[i].textObject.startOfContent(), end: paras[i].textObject.endOfContent()}});
        r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: (h1Applied > 0)}});
        h1Applied++;
        if (firstH1 < 0) firstH1 = i;
        delete headingSet[pText];
    }}
}}
for (var i = firstH1; i < paras.length; i++) {{
    var raw = paras[i].textObject.content();
    var clean = raw.replace(/[\\r\\n]/g, "");
    if (subArticleRe.test(clean)) {{
        var r = d.createRange({{start: paras[i].textObject.startOfContent(), end: paras[i].textObject.endOfContent()}});
        r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: true}});
        r.listFormat.listLevelNumber = 2;
        var stripped = clean.replace(/^\\d{{1,2}}\\.\\d{{1,2}}[\\.:]?[\\s\\t]*/, "");
        if (stripped !== clean) {{ paras[i].textObject.content = stripped + "\\r"; }}
        h2Applied++;
    }}
}}
{"" if not font_color else f"""
var fc = {_color_to_mac_rgb(font_color)};
for (var i = firstH1; i < paras.length; i++) {{
    try {{ paras[i].textObject.fontObject.color = fc; }} catch(e) {{}}
}}
"""}
JSON.stringify({{applied: true, type: "multilevel", h1: h1Applied, h2: h2Applied}});
""", timeout=180)
        else:
            para_indices_js = json.dumps([i - 1 for i in range(start_paragraph, ep + 1)])
            if lm:
                para_indices_js = json.dumps(sorted([int(k) - 1 for k in lm.keys()]))

            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var lt = app.make({{new: "listTemplate", at: d, withProperties: {{outlineNumbered: {_outline_js}}}}});
{levels_js}
var paraIndices = {para_indices_js};
var applied = 0;
for (var pi = 0; pi < paraIndices.length; pi++) {{
    var idx = paraIndices[pi];
    if (idx >= 0 && idx < d.paragraphs.length) {{
        var p = d.paragraphs[idx];
        var r = d.createRange({{start: p.textObject.startOfContent(), end: p.textObject.endOfContent()}});
        r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: (pi > 0)}});
        applied++;
    }}
}}
JSON.stringify({{applied: true, type: "multilevel", count: applied}});
""", timeout=120)

    else:
        gallery_idx = 0 if list_type == "bullet" else 1
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var gallery = app.listGalleries[{gallery_idx}];
var lt = gallery.listTemplates[0];
for (var i = {start_paragraph - 1}; i < {ep}; i++) {{
    var p = d.paragraphs[i];
    var r = d.createRange({{start: p.textObject.startOfContent(), end: p.textObject.endOfContent()}});
    var shouldContinue = (i > {start_paragraph - 1}) || {"true" if continue_previous else "false"};
    r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: shouldContinue}});
}}
JSON.stringify({{applied: true, type: "{list_type}", count: {ep - start_paragraph + 1}}});
""")


def mac_setup_heading_numbering(
    filename: str = None,
    h1_paragraphs: list = None,
    h2_paragraphs: list = None,
    heading_map: dict = None,
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
    """Set up auto-numbered headings with multilevel list on Mac.

    heading_map: optional dict {heading_text: level (1 or 2)} — text-based
    matching that's robust to index shifts. When provided, takes precedence
    over h1_paragraphs/h2_paragraphs. Uses the same matching as mac_apply_list:
    case-insensitive, apostrophe-normalized, with startsWith fallback.
    """
    import re as _re

    if heading_map:
        # Resolve heading_map -> h1_paragraphs / h2_paragraphs by text matching.
        _quote_chars = [chr(0x2018), chr(0x2019), chr(0x201c), chr(0x201d), chr(0x22)]

        def _norm(s: str) -> str:
            s = s.upper()
            for c in _quote_chars:
                s = s.replace(c, "'")
            return _re.sub(r"\s+", " ", s).strip()

        norm_map = {_norm(k): int(v) for k, v in heading_map.items()}
        paras_data = json.loads(mac_get_text(filename=filename)).get("paragraphs", [])
        h1_paragraphs = []
        h2_paragraphs = []
        num_prefix_re = _re.compile(r"^\s*\d{1,2}\.\d{1,2}[\.:]?[\s\t]*")
        for pdata in paras_data:
            raw = (pdata.get("text") or "").replace("\r", "").replace("\n", "")
            idx_1based = pdata.get("index", -1) + 1  # the function expects 1-indexed
            if not raw:
                continue
            p_text = _norm(raw)
            p_text_no_num = _norm(num_prefix_re.sub("", raw))
            level = 0
            matched_key = None
            if p_text in norm_map:
                level = norm_map[p_text]
                matched_key = p_text
            elif p_text_no_num != p_text and p_text_no_num in norm_map:
                level = norm_map[p_text_no_num]
                matched_key = p_text_no_num
            else:
                for key in list(norm_map.keys()):
                    if p_text.startswith(key) or (
                        p_text_no_num != p_text and p_text_no_num.startswith(key)
                    ):
                        level = norm_map[key]
                        matched_key = key
                        break
            if level > 0 and matched_key:
                del norm_map[matched_key]
                if level == 1:
                    h1_paragraphs.append(idx_1based)
                elif level == 2:
                    h2_paragraphs.append(idx_1based)

    if not h1_paragraphs and not h2_paragraphs:
        return json.dumps({"error": "Provide h1_paragraphs and/or h2_paragraphs (or heading_map)"})

    finder = _doc_finder_js(filename)
    h1_fmt = _escape_js(h1_number_format or "%1.")
    h2_fmt = _escape_js(h2_number_format or "%1.%2")

    align_map = {"left": "align paragraph left", "center": "align paragraph center",
                 "right": "align paragraph right", "justify": "align paragraph justify"}
    align_val = f'"{align_map.get(alignment.lower())}"' if alignment else "null"

    color_js = "null"
    if font_color:
        color_js = _color_to_mac_rgb(font_color)

    h1_indices_js = json.dumps([i - 1 for i in (h1_paragraphs or [])])
    h2_indices_js = json.dumps([i - 1 for i in (h2_paragraphs or [])])

    result_raw = _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var paras = d.paragraphs();

// --- Customize heading styles ---
var fontName = {json.dumps(font_name)};
var h1Size = {json.dumps(h1_size)};
var h2Size = {json.dumps(h2_size)};
var boldVal = {json.dumps(bold)};
var alignVal = {align_val};
var colorVal = {color_js};
var h1SpBefore = {json.dumps(h1_space_before)};
var h1SpAfter = {json.dumps(h1_space_after)};
var h2SpBefore = {json.dumps(h2_space_before)};
var h2SpAfter = {json.dumps(h2_space_after)};
var lineSpacing = {json.dumps(line_spacing)};

var ws = d.wordStyles;
var h1Style = ws["heading 1"];
var h2Style = ws["heading 2"];
var sizes = [h1Size, h2Size];
var spBefores = [h1SpBefore, h2SpBefore];
var spAfters = [h1SpAfter, h2SpAfter];
var styleObjs = [h1Style, h2Style];

for (var si = 0; si < 2; si++) {{
    var s = styleObjs[si];
    if (fontName !== null) s.font.name = fontName;
    if (sizes[si] !== null) s.font.size = sizes[si];
    if (boldVal !== null) {{ s.font.bold = boldVal; s.font.italic = false; }}
    if (colorVal !== null) s.font.color = colorVal;
    if (alignVal !== null) s.paragraphFormat.alignment = alignVal;
    if (spBefores[si] !== null) s.paragraphFormat.spaceBefore = spBefores[si];
    if (spAfters[si] !== null) s.paragraphFormat.spaceAfter = spAfters[si];
    if (lineSpacing !== null) {{
        s.paragraphFormat.lineSpacingRule = "line spacing multiple";
        s.paragraphFormat.lineSpacing = lineSpacing;
    }}
    s.paragraphFormat.keepWithNext = (si === 0);
    s.paragraphFormat.keepTogether = false;
}}

// --- Create multilevel list template ---
var lt = app.make({{new: "listTemplate", at: d, withProperties: {{outlineNumbered: true}}}});

var lv1 = lt.listLevels[0];
lv1.numberFormat = "{h1_fmt}";
lv1.numberStyle = "{
    'list number style upper case roman' if any(k in (h1_number_format or '').upper() for k in ['BÖLÜM', 'BOLUM', 'ROMAN'])
    else 'list number style arabic'
}";
lv1.startAt = 1;
lv1.numberPosition = 0;
lv1.textPosition = {0 if len(h1_number_format or "") > 5 else 28};
// Link Level 1 to Heading 1 so paragraphs styled Heading 1 auto-number.
// Don't link to Normal — that pollutes every Normal-styled paragraph.
try {{ lv1.linkedStyle = "heading 1"; }} catch(e1) {{}}

var lv2 = lt.listLevels[1];
lv2.numberFormat = "{h2_fmt}";
lv2.numberStyle = "list number style arabic";
lv2.startAt = 1;
lv2.numberPosition = 0;
lv2.textPosition = {0 if len(h2_number_format or "") > 5 else 28};
try {{ lv2.linkedStyle = "heading 2"; }} catch(e2) {{}}
try {{ lv2.restartAfter = 1; }} catch(e3) {{}}

// --- Apply styles AND list template to each paragraph ---
// linkedStyle assignment is unreliable on Mac JXA; explicitly apply the
// list template to each heading paragraph so they get the configured
// "MADDE %1 –" / "%1.%2." format. continuePreviousList:true on every
// paragraph after the first chains them into one outline.
var h1Indices = {h1_indices_js};
var h2Indices = {h2_indices_js};
var h1Applied = 0;
var h2Applied = 0;
var firstHeading = null;

for (var i = 0; i < h1Indices.length; i++) {{
    var idx = h1Indices[i];
    if (idx >= 0 && idx < paras.length) {{
        paras[idx].textObject.style = h1Style;
        var p = paras[idx];
        var r = d.createRange({{start: p.textObject.startOfContent(), end: p.textObject.endOfContent()}});
        r.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: (firstHeading !== null)}});
        if (firstHeading === null) firstHeading = idx;
        h1Applied++;
    }}
}}
for (var i = 0; i < h2Indices.length; i++) {{
    var idx = h2Indices[i];
    if (idx >= 0 && idx < paras.length) {{
        paras[idx].textObject.style = h2Style;
        var p2 = paras[idx];
        var r2 = d.createRange({{start: p2.textObject.startOfContent(), end: p2.textObject.endOfContent()}});
        r2.listFormat.applyListFormatTemplate({{listTemplate: lt, continuePreviousList: true}});
        r2.listFormat.listLevelNumber = 2;
        h2Applied++;
    }}
}}

JSON.stringify({{h1_applied: h1Applied, h2_applied: h2Applied}});
""", timeout=600)

    result = json.loads(result_raw)

    if strip_manual_numbers:
        stripped = 0
        all_indices = [(i - 1) for i in (h1_paragraphs or [])] + [(i - 1) for i in (h2_paragraphs or [])]
        full_text = json.loads(mac_get_text(filename))
        paragraphs = full_text.get("paragraphs", [])

        for idx in all_indices:
            if idx < 0 or idx >= len(paragraphs):
                continue
            text = paragraphs[idx].get("text", "")
            cleaned = re.sub(r'^(MADDE\s+\d+\s*[–\-:]\s*|\d+(\.\d+)*\.?\s+)', '', text)
            if cleaned != text:
                mac_replace_text(
                    filename=filename,
                    find_text=text[:80].rstrip(),
                    replace_text=cleaned[:80].rstrip(),
                    match_case=True,
                    replace_all=False,
                )
                stripped += 1

        result["stripped"] = stripped

    return json.dumps(result)
