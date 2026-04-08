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


def mac_get_text(filename: str = None) -> str:
    """Get all paragraph text from document."""
    finder = _doc_finder_js(filename)
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var paras = d.paragraphs();
var result = [];
for (var i = 0; i < paras.length; i++) {{
    var text = paras[i].textObject.content();
    result.push({{index: i, text: text}});
}}
JSON.stringify({{paragraphs: result, count: result.length}});
""")


def mac_get_page_text(filename: str = None, page: int = 1, end_page: int = None) -> str:
    """Get text from a specific page range."""
    finder = _doc_finder_js(filename)
    ep = end_page or page
    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var totalPages = parseInt(app.getRangeInformation(d.textObject, {{informationType: "number of pages in document"}}));
if ({page} > totalPages) throw new Error("Page {page} exceeds document (" + totalPages + " pages)");
var results = [];
for (var p = {page}; p <= Math.min({ep}, totalPages); p++) {{
    // Navigate to page
    var r = app.goTo(d.content, {{what: "go to page", which: "go to absolute", count: p}});
    var startPos = r.startOfContent();
    var endPos;
    if (p < totalPages) {{
        var r2 = app.goTo(d.content, {{what: "go to page", which: "go to absolute", count: p + 1}});
        endPos = r2.startOfContent();
    }} else {{
        endPos = d.textObject.endOfContent();
    }}
    var pageRange = d.createRange({{start: startPos, end: endPos}});
    results.push({{page: p, text: pageRange.content()}});
}}
JSON.stringify({{pages: results, totalPages: totalPages}});
""")


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
sel.homeKey({{unit: "a story", extend: "move"}});
var results = [];
for (var i = 0; i < {max_results}; i++) {{
    var f = sel.findObject;
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
    sel.homeKey({{unit: "a story", extend: "move"}});
    var f = sel.findObject;
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
        fmt_lines.append(f'r.fontObject.color = "{_escape_js(font_color)}";')
    if highlight_color:
        fmt_lines.append(f'r.highlightColorIndex = "{_escape_js(highlight_color)}";')
    if style_name:
        fmt_lines.append(f'r.style = "{_escape_js(style_name)}";')
    if paragraph_alignment:
        fmt_lines.append(f'r.paragraphFormat.alignment = "{_escape_js(paragraph_alignment)}";')
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
var ps = d.sections[{section_index}].pageSetup;
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
var s = d.sections[{section_index}];
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
var r = d.paragraphs[{paragraph_index}].textObject;
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
    # Build data population JS
    data_js = ""
    if data:
        data_json = json.dumps(data)
        data_js = f"""
    var data = {data_json};
    for (var r = 0; r < Math.min(data.length, {rows}); r++) {{
        for (var c = 0; c < Math.min(data[r].length, {cols}); c++) {{
            var cell = app.getCellFromTable(tbl, {{row: r + 1, column: c + 1}});
            cell.textObject.content = String(data[r][c]);
        }}
    }}
"""
    pos_js = "var pos = d.textObject.endOfContent() - 1;" if position == "end" else f"var pos = {position};"
    if position == "start":
        pos_js = "var pos = 0;"

    return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var prevTracking = d.trackRevisions();
if ({"true" if track_changes else "false"}) d.trackRevisions = true;
try {{
    {pos_js}
    var r = d.createRange({{start: pos, end: pos}});
    // Insert tab-delimited text and convert to table
    var rows = [];
    for (var i = 0; i < {rows}; i++) {{
        var cells = [];
        for (var j = 0; j < {cols}; j++) cells.push(" ");
        rows.push(cells.join("\\t"));
    }}
    r.content = rows.join("\\r");
    var newRange = d.createRange({{start: pos, end: pos + rows.join("\\r").length}});
    var tbl = app.convertToTable(newRange, {{numberOfColumns: {cols}, separator: "separate by tabs"}});
    {data_js}
}} finally {{
    d.trackRevisions = prevTracking;
}}
JSON.stringify({{added: true, tables: d.tables.length}});
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
var t = d.tables[{table_index}];
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
    var t = d.tables[{table_index}];
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
var t = d.tables[{table_index}];
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
var t = d.tables[{table_index}];
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
