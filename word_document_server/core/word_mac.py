"""JXA (JavaScript for Automation) bridge for Microsoft Word on macOS.

Provides functions to interact with Word for Mac via osascript.
Each function is self-contained: builds a JXA script, executes it,
and returns JSON-compatible results matching the Windows COM tool outputs.

Only works on macOS with Microsoft Word installed.
"""

import json
import os
import subprocess
import sys
import unicodedata
from contextlib import contextmanager


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
    result = subprocess.run(
        ["/usr/bin/osascript", "-l", "JavaScript"],
        input=script,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
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
    result = subprocess.run(
        ["/usr/bin/osascript"],
        input=script,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
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

function findPageStart(pg) {{
    if (pg <= 1) return 0;
    if (pg > totalPages) return totalEnd;
    var lo = 0, hi = totalEnd;
    while (lo < hi) {{
        var mid = Math.floor((lo + hi) / 2);
        var r = app.createRange(d, {{start: mid, end: Math.min(mid + 1, totalEnd)}});
        var pn = parseInt(app.getRangeInformation(r, {{informationType: "active end page number"}}));
        if (pn < pg) lo = mid + 1;
        else hi = mid;
    }}
    return lo;
}}

var pages = [];
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
var sel = app.selection;
sel.selectionStart = 0;
sel.selectionEnd = 0;
var f = sel.findObject;
f.clearFormatting();
f.replacement.clearFormatting();
f.replacement.content = "";
var results = [];
for (var i = 0; i < {max_results}; i++) {{
    f.content = "{escaped_search}";
    f.forward = true;
    f.wrap = "find stop";
    f.matchCase = {"true" if match_case else "false"};
    f.matchWholeWord = {"true" if whole_word else "false"};
    f.matchWildcards = {"true" if use_wildcards else "false"};
    if (!f.executeFind()) break;
    var s = sel.selectionStart();
    var e = sel.selectionEnd();
    var ctxStart = Math.max(0, s - {context_chars});
    var ctxEnd = Math.min(d.textObject.endOfContent(), e + {context_chars});
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
var paras = d.paragraphs();
var startP = {start_paragraph};
var endP = {ep} !== null ? {ep} : startP;
endP = Math.min(endP, paras.length - 1);
var results = [];
for (var i = startP; i <= endP; i++) {{
    var p = paras[i];
    var pf = p.paragraphFormat;
    var fo = p.textObject.fontObject;
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
        fontName: fo.name(),
        fontSize: fo.fontSize(),
        bold: fo.bold(),
        italic: fo.italic()
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
var paras = d.paragraphs();
var issues = [];
var kwnChain = [];
for (var i = 0; i < paras.length; i++) {{
    var pf = paras[i].paragraphFormat;
    var kwn = pf.keepWithNext();
    var kt = pf.keepTogether();
    var pbb = pf.pageBreakBefore();
    var text = paras[i].textObject.content().substring(0, 60);
    if (kwn) {{
        kwnChain.push(i);
    }} else if (kwnChain.length > 2) {{
        issues.push({{type: "keep_with_next_chain", paragraphs: kwnChain.slice(), length: kwnChain.length, firstText: text}});
        kwnChain = [];
    }} else {{
        kwnChain = [];
    }}
    if (pbb) issues.push({{type: "page_break_before", paragraph: i, text: text}});
}}
if (kwnChain.length > 2) {{
    issues.push({{type: "keep_with_next_chain", paragraphs: kwnChain, length: kwnChain.length}});
}}
JSON.stringify({{issues: issues, totalParagraphs: paras.length}});
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
    escaped = _escape_js(text)
    # Handle literal \r\n → actual newlines for Word
    escaped = escaped.replace("\\\\r\\\\n", "\\r").replace("\\\\r", "\\r").replace("\\\\n", "\\r")

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
        bookmark_js = """
    var sel = app.selection;
    sel.content = text;
"""
    else:
        # Numeric position
        bookmark_js = f"""
    var pos = parseInt("{position}");
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
    if highlight_color:
        fmt_lines.append(f'r.highlightColorIndex = "{_escape_js(highlight_color)}";')
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

    bold_bleed_js = ""
    if bold is True:
        bold_bleed_js = (
            "var rStart = r.startOfContent();\n"
            "    var rEnd = r.endOfContent();\n"
            "    var txt = r.content();\n"
            "    for (var pi = 0; pi < txt.length; pi++) {\n"
            "        if (txt[pi] === '\\r') {\n"
            "            var pm = d.createRange({start: rStart + pi, end: rStart + pi + 1});\n"
            "            pm.bold = false;\n"
            "        }\n"
            "    }"
        )

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    {range_js}
    {fmt_js}
    // Prevent bold bleed: unbold paragraph marks within the range
    {bold_bleed_js}
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

    # JXA's make() doesn't work for Word comments — use AppleScript
    escaped_as_text = _escape_as(text)
    if start is not None and end is not None:
        range_as = f"set r to create range active document start {start} end {end}"
    elif paragraph_index is not None:
        range_as = f"set r to text object of paragraph {paragraph_index + 1} of active document"
    else:
        return json.dumps({"error": "Must provide start/end or paragraph_index"})

    result = _run_applescript(f'''
tell application "Microsoft Word"
    {range_as}
    make new Word comment at active document with properties {{comment text:"{escaped_as_text}", scope:r}}
    return count of Word comments of active document
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
    track_changes: bool = False,
) -> str:
    """Add a table to the document."""
    finder = _doc_finder_js(filename)
    pos_js = "var pos = d.textObject.endOfContent() - 1;" if position == "end" else f"var pos = {position};"
    if position == "start":
        pos_js = "var pos = 0;"

    data_js = ""
    if data:
        data_json = json.dumps(data)
        data_js = f"""
    var data = {data_json};
    for (var ri = 0; ri < Math.min(data.length, {rows}); ri++) {{
        for (var ci = 0; ci < Math.min(data[ri].length, {cols}); ci++) {{
            var tRef = d.tables[d.tables.length - 1];
        var cell = app.getCellFromTable(tRef, {{row: ri + 1, column: ci + 1}});
            cell.textObject.content = String(data[ri][ci]);
        }}
    }}
"""

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{pos_js}
var rng = app.createRange(d, {{start: pos, end: pos}});
var t = app.make({{new: "table", at: d, withProperties: {{
    textObject: rng,
    numberOfRows: {rows},
    numberOfColumns: {cols}
}}}});
{data_js}
JSON.stringify({{added: true, rows: {rows}, cols: {cols}, tables: d.tables.length}});
""")


def mac_modify_table(
    filename: str = None,
    table_index: int = 0,
    operation: str = "get_info",
    row: int = None,
    col: int = None,
    text: str = None,
    track_changes: bool = False,
) -> str:
    """Modify table structure or content."""
    finder = _doc_finder_js(filename)

    if operation == "get_info":
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var t = d.tables[Math.max(0, {table_index} - 1)];  // 1-based → 0-based
var rows = t.rows.length;
var cols = t.columns.length;
var cells = [];
for (var r = 1; r <= rows; r++) {{
    for (var c = 1; c <= cols; c++) {{
        var cell = app.getCellFromTable(t, {{row: r, column: c}});
        cells.push({{row: r, col: c, text: cell.textObject.content().replace(/[\\r\\x07]/g, "")}});
    }}
}}
JSON.stringify({{rows: rows, cols: cols, cells: cells}});
""")

    elif operation == "set_cell":
        escaped_text = _escape_js(text or "")
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    var t = d.tables[Math.max(0, {table_index} - 1)];  // 1-based → 0-based
    var cell = app.getCellFromTable(t, {{row: {row}, column: {col}}});
    cell.textObject.content = "{escaped_text}";
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{set: true, row: {row}, col: {col}}});
""")

    elif operation == "insert_row":
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var t = d.tables[Math.max(0, {table_index} - 1)];  // 1-based → 0-based
var targetRow = {row or "t.rows.length"};
var cell = app.getCellFromTable(t, {{row: targetRow, column: 1}});
app.select(cell.textObject);
app.insertRows(app.selection, {{numberOfRows: 1}});
JSON.stringify({{inserted: true, rows: t.rows.length}});
""")

    elif operation == "delete_row":
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var t = d.tables[Math.max(0, {table_index} - 1)];  // 1-based → 0-based
var targetRow = {row or "t.rows.length"};
app.delete(t.rows[targetRow - 1]);
JSON.stringify({{deleted: true, rows: t.rows.length}});
""")

    return json.dumps({"error": f"Unknown operation: {operation}"})


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
            h1_color_block = ""
            if font_color:
                h1_color_block = (
                    "var fc = " + _color_to_mac_rgb(font_color) + ";\n"
                    "for (var i = firstH1; i < paras.length; i++) {\n"
                    "    try { paras[i].textObject.fontObject.color = fc; } catch(e) {}\n"
                    "}"
                )
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
{h1_color_block}
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
