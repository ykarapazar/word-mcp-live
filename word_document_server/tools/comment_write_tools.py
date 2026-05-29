"""
Comment writing tools for Word Document Server.

These tools provide MCP interfaces for adding comments to Word documents.
"""

import json
import os

from word_document_server.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock
from word_document_server.core.comment_writer import add_comment_to_doc, add_comment_to_doc_by_paragraph_index


async def add_comment(
    filename: str,
    target_text: str,
    comment_text: str,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> str:
    """Add a comment to a Word document anchored to specific text.

    Args:
        filename: Path to Word document
        target_text: Text in the document to attach the comment to
        comment_text: The comment content
        author: Comment author name
        initials: Author initials

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not target_text:
        return json.dumps({"success": False, "error": "target_text cannot be empty"})
    if not comment_text:
        return json.dumps({"success": False, "error": "comment_text cannot be empty"})

    try:
        async with get_file_lock(filename):
            result = add_comment_to_doc(filename, target_text, comment_text, author, initials)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add comment: {str(e)}"})


async def add_comment_by_paragraph_index(
    filename: str,
    paragraph_index: int,
    comment_text: str,
    target_text: str = None,
    target_start: int = None,
    target_end: int = None,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> str:
    """Add a comment scoped to a specific document.xml paragraph index.

    paragraph_index follows raw XML order (includes table-cell paragraphs).
    If target_text is given, matching happens only inside that paragraph.
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if paragraph_index is None or paragraph_index < 0:
        return json.dumps({"success": False, "error": "paragraph_index must be a non-negative integer"})
    if not comment_text:
        return json.dumps({"success": False, "error": "comment_text cannot be empty"})

    try:
        async with get_file_lock(filename):
            result = add_comment_to_doc_by_paragraph_index(
                filename,
                paragraph_index,
                comment_text,
                author,
                initials,
                target_text=target_text,
                target_start=target_start,
                target_end=target_end,
            )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add comment by paragraph_index: {str(e)}"})
