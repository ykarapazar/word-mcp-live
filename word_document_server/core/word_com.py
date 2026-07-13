"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
import re
import sys
import unicodedata
from contextlib import contextmanager


def get_word_app():
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object that has open documents.
    When multiple Word instances are running, iterates through all
    Running Object Table (ROT) entries to find one with documents.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Word.Application")
        if app.Documents.Count > 0:
            return app
        # GetActiveObject found an empty instance — scan ROT for others
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        # No instance has documents; return the empty one (caller may open a file)
        return app
    except Exception:
        # GetActiveObject failed entirely — try ROT scan
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        raise RuntimeError(
            "Microsoft Word is not running. Please open Word first."
        )


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        # Pass 1: look for a Word.Application ROT entry with documents
        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                # Direct Application entry
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                # Remember file monikers for pass 2
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                # Also collect file monikers we couldn't QI yet
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        # Pass 2: try file monikers → Document → Application
        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_document_for_write(app, filename: str = None):
    """Resolve a document for a MUTATING operation. ``filename`` is REQUIRED.

    Read tools may default to the active document — a wrong guess merely returns
    the wrong text, and the caller sees it. A *write* that guesses wrong silently
    edits a file the user never named, and Word's ActiveDocument follows whatever
    window they last clicked, across virtual desktops. So mutating tools must say
    which document they mean.

    Set ``WORD_MCP_ALLOW_ACTIVE_WRITES=1`` to restore the old implicit-active
    behaviour (not recommended; provided as an escape hatch).

    Raises:
        ValueError: If *filename* is omitted and the escape hatch is not set.
    """
    if not filename:
        if os.environ.get("WORD_MCP_ALLOW_ACTIVE_WRITES") == "1":
            return find_document(app, None)
        try:
            active = app.ActiveDocument.FullName
        except Exception:
            active = "<unknown>"
        raise ValueError(
            "filename is required for operations that modify a document. "
            "Word's active document follows whichever window you last clicked, "
            "so an omitted filename can silently edit the wrong file. "
            f"The currently active document is '{active}' — pass its full path "
            "explicitly if that is the one you meant. "
            "(Use word_live_list_open to see all open documents.)"
        )
    return find_document(app, filename)


def validate_live_filename(filename: str) -> str:
    """Reject filenames that the live (COM) tools must never accept.

    The live tools act only on documents already open in Word, so a caller has
    no legitimate reason to pass a traversal sequence, a UNC path, or a device
    path. ``find_document`` enforces the real invariant (the resolved document
    must be in ``app.Documents``), which is what actually contains the blast
    radius; this is a cheap early reject for obviously hostile input.

    Args:
        filename: Caller-supplied document name or path (may be None).

    Returns:
        The filename unchanged, if acceptable.

    Raises:
        ValueError: If the filename contains a traversal or device/UNC prefix.
    """
    if not filename:
        return filename

    if "\x00" in filename:
        raise ValueError("filename contains a null byte")

    # Reject UNC (\\server\share) and Win32 device paths (\\?\, \\.\).
    if filename.startswith("\\\\") or filename.startswith("//"):
        raise ValueError(
            f"Refusing UNC/device path: '{filename}'. "
            "Live tools operate only on documents already open in Word."
        )

    # Reject traversal segments. Normpath collapses them, so compare before/after:
    # any path whose normalized form still escapes upward is rejected outright.
    parts = re.split(r"[\\/]+", filename)
    if any(part == ".." for part in parts):
        raise ValueError(
            f"Refusing path traversal sequence: '{filename}'. "
            "Pass either a bare document name or a full absolute path."
        )

    return filename


def _norm(value: str) -> str:
    """Normalize a path or name for case-insensitive, Unicode-safe comparison."""
    return unicodedata.normalize('NFC', value).lower()


def _norm_path(value: str) -> str:
    """Normalize a full path: resolve separators, case, and Unicode form.

    ``os.path.normpath`` collapses ``..`` segments and unifies separators, so
    ``C:\\a\\..\\b\\x.docx`` and ``C:/b/x.docx`` compare equal.
    """
    return _norm(os.path.normpath(value))


def find_document(app, filename: str = None):
    """Find an open document by full path or basename.

    Resolution is deliberately two-pass and fails loudly on ambiguity:

    1. **Full-path pass.** If *filename* is an absolute path, match it against
       every open document's ``FullName``. An exact path is unambiguous, so it
       always wins.
    2. **Basename pass.** Only if the full-path pass found nothing. If two or
       more open documents share the basename, raise rather than guess.

    The old implementation checked basename *before* full path inside a single
    loop, so with ``C:\\A\\Lettre.docx`` and ``C:\\B\\Lettre.docx`` both open,
    asking for ``C:\\B\\Lettre.docx`` returned ``C:\\A\\Lettre.docx`` — a silent
    write to the wrong file. Never reintroduce a single-pass loop here.

    Args:
        app: Word.Application COM object.
        filename: Document basename or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If no documents are open, the document is not open, or the
                    basename is ambiguous across two or more open documents.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    validate_live_filename(filename)

    # Snapshot the open set once: app.Documents(i) is a live COM collection and
    # the user may open/close documents while we iterate.
    open_docs = []
    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        open_docs.append((doc, doc.Name, doc.FullName))

    # Pass 1 — full path wins outright when the caller gave us one.
    if os.path.isabs(filename):
        target_fullpath = _norm_path(filename)
        for doc, _name, full_name in open_docs:
            if _norm_path(full_name) == target_fullpath:
                return doc

    # Pass 2 — basename, but only when it identifies exactly one document.
    target_basename = _norm(os.path.basename(filename))
    matches = [
        (doc, full_name)
        for doc, name, full_name in open_docs
        if _norm(name) == target_basename
    ]

    if len(matches) == 1:
        return matches[0][0]

    if len(matches) > 1:
        # Join explicitly: a list repr double-escapes Windows backslashes,
        # turning C:\Clients into C:\\Clients in the message the user reads.
        candidates = "; ".join(full_name for _doc, full_name in matches)
        raise ValueError(
            f"Ambiguous document name '{os.path.basename(filename)}': "
            f"{len(matches)} open documents share this name. "
            f"Pass the full path to disambiguate. Candidates: {candidates}"
        )

    open_list = "; ".join(full_name for _d, _n, full_name in open_docs)
    raise ValueError(
        f"Document '{filename}' is not open in Word. Open documents: {open_list}"
    )


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    rec = None
    try:
        rec = app.UndoRecord
        # Clean up stale undo record from a previous crash/interrupted session
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None  # Word 2007 or earlier — proceed without
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
