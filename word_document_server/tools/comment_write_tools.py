"""
Comment writing tools for Word Document Server.

These tools provide MCP interfaces for adding comments to Word documents.
"""

import json
import os

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.core.comment_writer import add_comment_to_doc


async def add_comment(
    filename: str,
    target_text: str,
    comment_text: str,
    author: str = "Av. Yüce Karapazar",
    initials: str = "AYK",
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
        result = add_comment_to_doc(filename, target_text, comment_text, author, initials)
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add comment: {str(e)}"})
