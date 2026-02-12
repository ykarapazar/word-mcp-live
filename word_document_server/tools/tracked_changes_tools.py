"""
Tracked changes tools for Word Document Server.

These tools provide MCP interfaces for creating, listing, accepting,
and rejecting tracked changes in Word documents.
"""

import json
import os
from typing import Optional

from word_document_server.defaults import DEFAULT_AUTHOR

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.core.tracked_changes import (
    track_replace_in_doc,
    track_insert_in_doc,
    track_delete_in_doc,
    list_tracked_changes_in_doc,
    accept_tracked_changes_in_doc,
    reject_tracked_changes_in_doc,
)


async def track_replace(
    filename: str,
    old_text: str,
    new_text: str,
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Replace text content with tracked changes (marks old as deleted, new as inserted).
    This changes TEXT CONTENT only — it does not change formatting (font, highlight, style).
    To change formatting, use word_live_format_text instead.

    Args:
        filename: Path to Word document.
        old_text: Text to find and mark as deleted.
        new_text: Replacement text to insert (must differ from old_text).
        author: Author name for the tracked change.

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not old_text:
        return json.dumps({"success": False, "error": "old_text cannot be empty"})

    try:
        result = track_replace_in_doc(filename, old_text, new_text, author)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to track replace: {str(e)}"})


async def track_insert(
    filename: str,
    after_text: str,
    insert_text: str,
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Insert text content after a specific string, marked as a tracked insertion.
    This changes TEXT CONTENT only — it does not change formatting.

    Args:
        filename: Path to Word document.
        after_text: Text to search for; new text is inserted right after this.
        insert_text: Text to insert.
        author: Author name for the tracked change.

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not after_text:
        return json.dumps({"success": False, "error": "after_text cannot be empty"})
    if not insert_text:
        return json.dumps({"success": False, "error": "insert_text cannot be empty"})

    try:
        result = track_insert_in_doc(filename, after_text, insert_text, author)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to track insert: {str(e)}"})


async def track_delete(
    filename: str,
    text: str,
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Mark text content as deleted (tracked deletion).
    This changes TEXT CONTENT only — it does not change formatting.

    Args:
        filename: Path to Word document.
        text: Text to mark as deleted.
        author: Author name for the tracked change.

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not text:
        return json.dumps({"success": False, "error": "text cannot be empty"})

    try:
        result = track_delete_in_doc(filename, text, author)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to track delete: {str(e)}"})


async def list_tracked_changes(filename: str) -> str:
    """List all tracked changes in a Word document.

    Args:
        filename: Path to Word document

    Returns:
        JSON string with insertions, deletions, and counts
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    try:
        result = list_tracked_changes_in_doc(filename)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to list tracked changes: {str(e)}"})


async def accept_tracked_changes(
    filename: str,
    author: Optional[str] = None,
    change_ids: Optional[list[int]] = None,
) -> str:
    """Accept tracked changes (apply insertions, remove deletions).

    Args:
        filename: Path to Word document
        author: If specified, only accept changes by this author
        change_ids: If specified, only accept changes with these IDs

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    try:
        result = accept_tracked_changes_in_doc(filename, author, change_ids)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to accept tracked changes: {str(e)}"})


async def reject_tracked_changes(
    filename: str,
    author: Optional[str] = None,
    change_ids: Optional[list[int]] = None,
) -> str:
    """Reject tracked changes (remove insertions, restore deletions).

    Args:
        filename: Path to Word document
        author: If specified, only reject changes by this author
        change_ids: If specified, only reject changes with these IDs

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    try:
        result = reject_tracked_changes_in_doc(filename, author, change_ids)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to reject tracked changes: {str(e)}"})
