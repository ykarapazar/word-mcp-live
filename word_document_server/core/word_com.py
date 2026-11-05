"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
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


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.
                  A basename shared by two open documents is refused;
                  a full path picks the exact one.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    target_basename = unicodedata.normalize('NFC', os.path.basename(filename)).lower()
    target_fullpath = (
        unicodedata.normalize('NFC', os.path.normpath(filename)).lower()
        if os.path.isabs(filename) else None
    )

    docs = [app.Documents(i) for i in range(1, app.Documents.Count + 1)]
    if target_fullpath:
        for doc in docs:
            if unicodedata.normalize('NFC', os.path.normpath(doc.FullName)).lower() == target_fullpath:
                return doc

    same_name = [doc for doc in docs
                 if unicodedata.normalize('NFC', doc.Name).lower() == target_basename]
    if len(same_name) == 1:
        return same_name[0]
    if len(same_name) > 1:
        raise ValueError(
            f"Two or more open documents are named '{same_name[0].Name}': "
            f"{[d.FullName for d in same_name]}. "
            "Pass the full path as filename to choose one."
        )

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise ValueError(
        f"Document '{filename}' is not open in Word. "
        f"Open documents: {open_docs}"
    )


@contextmanager
def undo_record(app, name: str, doc=None):
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
    # UndoRecord groups edits only on the ActiveDocument's undo stack; when the
    # target doc is not active, each COM edit lands as its own undo entry.
    prev = None
    if doc is not None:
        try:
            active = app.ActiveDocument
            if active.FullName != doc.FullName:
                prev = active
                doc.Activate()
        except Exception:
            prev = None
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
        if prev is not None:
            try:
                prev.Activate()
            except Exception:
                pass


def set_author(app, name: str, initials: str = None):
    """Make Word attribute new revisions and comments to *name*.

    Word ignores Application.UserName when the user is signed in to Office,
    unless Options.UseLocalUserInfo ("Always use these values regardless of
    sign in to Office") is on. Turn it on for the duration of the edit.
    Returns the previous state for restore_author().
    """
    state = {"name": app.UserName, "initials": None, "local": None}
    try:
        state["local"] = app.Options.UseLocalUserInfo
        if not state["local"]:
            app.Options.UseLocalUserInfo = True
    except Exception:
        state["local"] = None  # Word 2013 or earlier: no such option
    app.UserName = name
    if initials:
        state["initials"] = app.UserInitials
        app.UserInitials = initials
    return state


def restore_author(app, state) -> None:
    """Undo set_author()."""
    app.UserName = state["name"]
    if state["initials"] is not None:
        app.UserInitials = state["initials"]
    if state["local"] is False:
        try:
            app.Options.UseLocalUserInfo = False
        except Exception:
            pass
