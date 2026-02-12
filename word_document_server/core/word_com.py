"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
import sys
from contextlib import contextmanager


def get_word_app():
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    import win32com.client

    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        raise RuntimeError(
            "Microsoft Word is not running. Please open Word first."
        )


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    target_basename = os.path.basename(filename).lower()
    target_fullpath = (
        os.path.normpath(filename).lower() if os.path.isabs(filename) else None
    )

    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if doc.Name.lower() == target_basename:
            return doc
        if target_fullpath and os.path.normpath(doc.FullName).lower() == target_fullpath:
            return doc

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise ValueError(
        f"Document '{filename}' is not open in Word. "
        f"Open documents: {open_docs}"
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
