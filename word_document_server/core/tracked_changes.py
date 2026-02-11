"""
Core tracked changes functionality for Word documents.

Implements insertion, deletion, and replacement of text as tracked changes
using lxml to manipulate the underlying OOXML structure. python-docx doesn't
support tracked changes natively, so we work directly with the XML.

Reference: Anthropic docx skill SKILL.md tracked changes patterns.
"""

import copy
import json
import random
import re
import zipfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Optional

from lxml import etree

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}

# Qualified name helpers
W = lambda tag: f"{{{WORD_NS}}}{tag}"


def _generate_id(root: etree._Element) -> int:
    """Generate a unique w:id by finding max existing id + 1."""
    max_id = 0
    for elem in root.iter():
        val = elem.get(W("id"))
        if val is not None:
            try:
                max_id = max(max_id, int(val))
            except ValueError:
                pass
    return max_id + 1


def _now_iso() -> str:
    """Return current UTC timestamp in OOXML format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _get_run_text(run: etree._Element) -> str:
    """Extract all text from a run's <w:t> elements."""
    parts = []
    for t in run.findall(W("t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _get_run_rpr(run: etree._Element) -> Optional[etree._Element]:
    """Get a deep copy of a run's <w:rPr> (formatting), or None."""
    rpr = run.find(W("rPr"))
    if rpr is not None:
        return copy.deepcopy(rpr)
    return None


def _make_run(text: str, rpr: Optional[etree._Element] = None, is_del: bool = False) -> etree._Element:
    """Create a <w:r> element with text and optional formatting.

    Args:
        text: The text content
        rpr: Optional formatting to copy
        is_del: If True, use <w:delText> instead of <w:t>
    """
    r = etree.SubElement(etree.Element("dummy"), W("r"))
    r.getparent().remove(r)  # Detach from dummy

    if rpr is not None:
        r.append(copy.deepcopy(rpr))

    tag = W("delText") if is_del else W("t")
    t = etree.SubElement(r, tag)
    t.text = text
    # Preserve whitespace
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return r


def _load_document_xml(filepath: str) -> tuple[etree._Element, bytes]:
    """Load and parse document.xml from a .docx file.

    Returns:
        Tuple of (parsed XML root, original zip bytes)
    """
    filepath = Path(filepath)
    zip_bytes = filepath.read_bytes()

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        doc_xml = zf.read("word/document.xml")

    root = etree.fromstring(doc_xml)
    return root, zip_bytes


def _save_document_xml(filepath: str, root: etree._Element, original_zip_bytes: bytes) -> None:
    """Save modified document.xml back into the .docx zip."""
    filepath = Path(filepath)
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Read original zip and replace document.xml
    buffer = BytesIO()
    with zipfile.ZipFile(BytesIO(original_zip_bytes), "r") as zf_in:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename == "word/document.xml":
                    zf_out.writestr(item, new_xml)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

    filepath.write_bytes(buffer.getvalue())


def _get_paragraphs(root: etree._Element) -> list[etree._Element]:
    """Get all <w:p> elements from body."""
    body = root.find(W("body"))
    if body is None:
        return []
    return body.findall(f".//{W('p')}")


def _paragraph_text(p: etree._Element) -> str:
    """Get plain text of a paragraph."""
    texts = []
    for t in p.iter(W("t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _find_text_in_paragraph(p: etree._Element, search_text: str) -> Optional[list[tuple]]:
    """Find where search_text appears across runs in a paragraph.

    Returns list of (run_element, start_offset, end_offset) tuples
    that together contain the search_text, or None if not found.
    """
    runs = [r for r in p.findall(f".//{W('r')}") if r.find(W("t")) is not None]
    if not runs:
        return None

    # Build a map: (run_index, char_offset_in_run) for each character
    char_map = []
    for ri, run in enumerate(runs):
        text = _get_run_text(run)
        for ci in range(len(text)):
            char_map.append((ri, ci))

    full_text = "".join(_get_run_text(r) for r in runs)
    pos = full_text.find(search_text)
    if pos == -1:
        return None

    # Map character positions back to runs
    start_ri, start_ci = char_map[pos]
    end_ri, end_ci = char_map[pos + len(search_text) - 1]

    result = []
    for ri in range(start_ri, end_ri + 1):
        run = runs[ri]
        run_text = _get_run_text(run)
        s = start_ci if ri == start_ri else 0
        e = end_ci + 1 if ri == end_ri else len(run_text)
        result.append((run, s, e))

    return result


def track_replace_in_doc(
    filepath: str,
    old_text: str,
    new_text: str,
    author: str = "Av. Yüce Karapazar",
) -> dict:
    """Replace text with tracked changes (delete old + insert new).

    Args:
        filepath: Path to .docx file
        old_text: Text to find and mark as deleted
        new_text: Text to insert as replacement
        author: Author name for the tracked change

    Returns:
        Dict with success status and details
    """
    root, zip_bytes = _load_document_xml(filepath)
    timestamp = _now_iso()
    replacements = 0

    for p in _get_paragraphs(root):
        # Keep searching in the same paragraph until no more matches
        while True:
            match = _find_text_in_paragraph(p, old_text)
            if match is None:
                break

            next_id = _generate_id(root)

            # Get formatting from the first matched run
            first_run = match[0][0]
            rpr = _get_run_rpr(first_run)

            # Build the replacement elements
            del_elem = etree.Element(W("del"))
            del_elem.set(W("id"), str(next_id))
            del_elem.set(W("author"), author)
            del_elem.set(W("date"), timestamp)
            del_run = _make_run(old_text, rpr, is_del=True)
            del_elem.append(del_run)

            ins_elem = etree.Element(W("ins"))
            ins_elem.set(W("id"), str(next_id + 1))
            ins_elem.set(W("author"), author)
            ins_elem.set(W("date"), timestamp)
            ins_run = _make_run(new_text, rpr, is_del=False)
            ins_elem.append(ins_run)

            # Now splice: we need to handle the matched runs
            # Strategy: split first and last runs, remove middle runs,
            # insert del+ins at the right position

            first_run_elem, first_start, first_end = match[0]
            last_run_elem, last_start, last_end = match[-1]
            first_run_text = _get_run_text(first_run_elem)
            last_run_text = _get_run_text(last_run_elem)

            # Find the parent of the first run (should be <w:p> or a tracked change wrapper)
            parent = first_run_elem.getparent()

            # Text before the match in the first run
            before_text = first_run_text[:first_start]
            # Text after the match in the last run
            after_text = last_run_text[last_end:]

            # Determine insertion point (index of first matched run in parent)
            insert_idx = list(parent).index(first_run_elem)

            # Remove all matched runs
            for run_elem, _, _ in match:
                run_parent = run_elem.getparent()
                if run_parent is not None:
                    run_parent.remove(run_elem)

            # Insert: before_text_run, del, ins, after_text_run
            offset = 0
            if before_text:
                before_run = _make_run(before_text, rpr)
                parent.insert(insert_idx + offset, before_run)
                offset += 1

            parent.insert(insert_idx + offset, del_elem)
            offset += 1
            parent.insert(insert_idx + offset, ins_elem)
            offset += 1

            if after_text:
                after_rpr = _get_run_rpr(last_run_elem) if last_run_elem != first_run_elem else rpr
                after_run = _make_run(after_text, after_rpr or rpr)
                parent.insert(insert_idx + offset, after_run)

            replacements += 1

    if replacements == 0:
        return {"success": False, "error": f"Text not found: '{old_text}'"}

    _save_document_xml(filepath, root, zip_bytes)
    return {
        "success": True,
        "replacements": replacements,
        "message": f"Replaced {replacements} occurrence(s) of '{old_text}' with '{new_text}' as tracked change by {author}",
    }


def track_insert_in_doc(
    filepath: str,
    after_text: str,
    insert_text: str,
    author: str = "Av. Yüce Karapazar",
) -> dict:
    """Insert text after a specific string, marked as a tracked insertion.

    Args:
        filepath: Path to .docx file
        after_text: Text to search for; new text will be inserted right after this
        insert_text: Text to insert
        author: Author name for the tracked change

    Returns:
        Dict with success status and details
    """
    root, zip_bytes = _load_document_xml(filepath)
    timestamp = _now_iso()
    insertions = 0

    for p in _get_paragraphs(root):
        match = _find_text_in_paragraph(p, after_text)
        if match is None:
            continue

        next_id = _generate_id(root)

        # Get formatting from the last matched run
        last_run_elem, last_start, last_end = match[-1]
        rpr = _get_run_rpr(last_run_elem)
        last_run_text = _get_run_text(last_run_elem)

        # Build insertion element
        ins_elem = etree.Element(W("ins"))
        ins_elem.set(W("id"), str(next_id))
        ins_elem.set(W("author"), author)
        ins_elem.set(W("date"), timestamp)
        ins_run = _make_run(insert_text, rpr, is_del=False)
        ins_elem.append(ins_run)

        parent = last_run_elem.getparent()
        run_idx = list(parent).index(last_run_elem)

        # If match ends mid-run, split the run
        after_match_text = last_run_text[last_end:]
        if after_match_text:
            # Truncate the current run to end at the match
            before_match_text = last_run_text[:last_end]
            for t in last_run_elem.findall(W("t")):
                last_run_elem.remove(t)
            new_t = etree.SubElement(last_run_elem, W("t"))
            new_t.text = before_match_text
            new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

            # Insert the tracked insertion
            parent.insert(run_idx + 1, ins_elem)

            # Insert remainder run after
            remainder_run = _make_run(after_match_text, rpr)
            parent.insert(run_idx + 2, remainder_run)
        else:
            # Match ends at run boundary, just insert after
            parent.insert(run_idx + 1, ins_elem)

        insertions += 1
        break  # Only first occurrence

    if insertions == 0:
        return {"success": False, "error": f"Text not found: '{after_text}'"}

    _save_document_xml(filepath, root, zip_bytes)
    return {
        "success": True,
        "insertions": insertions,
        "message": f"Inserted '{insert_text}' after '{after_text}' as tracked change by {author}",
    }


def track_delete_in_doc(
    filepath: str,
    text: str,
    author: str = "Av. Yüce Karapazar",
) -> dict:
    """Mark text as deleted (tracked deletion).

    Args:
        filepath: Path to .docx file
        text: Text to mark as deleted
        author: Author name for the tracked change

    Returns:
        Dict with success status and details
    """
    root, zip_bytes = _load_document_xml(filepath)
    timestamp = _now_iso()
    deletions = 0

    for p in _get_paragraphs(root):
        while True:
            match = _find_text_in_paragraph(p, text)
            if match is None:
                break

            next_id = _generate_id(root)

            first_run_elem, first_start, first_end = match[0]
            last_run_elem, last_start, last_end = match[-1]
            rpr = _get_run_rpr(first_run_elem)
            first_run_text = _get_run_text(first_run_elem)
            last_run_text = _get_run_text(last_run_elem)

            # Build deletion element
            del_elem = etree.Element(W("del"))
            del_elem.set(W("id"), str(next_id))
            del_elem.set(W("author"), author)
            del_elem.set(W("date"), timestamp)
            del_run = _make_run(text, rpr, is_del=True)
            del_elem.append(del_run)

            parent = first_run_elem.getparent()
            insert_idx = list(parent).index(first_run_elem)

            before_text = first_run_text[:first_start]
            after_text = last_run_text[last_end:]

            # Remove matched runs
            for run_elem, _, _ in match:
                run_parent = run_elem.getparent()
                if run_parent is not None:
                    run_parent.remove(run_elem)

            offset = 0
            if before_text:
                before_run = _make_run(before_text, rpr)
                parent.insert(insert_idx + offset, before_run)
                offset += 1

            parent.insert(insert_idx + offset, del_elem)
            offset += 1

            if after_text:
                after_rpr = _get_run_rpr(last_run_elem) if last_run_elem != first_run_elem else rpr
                after_run = _make_run(after_text, after_rpr or rpr)
                parent.insert(insert_idx + offset, after_run)

            deletions += 1

    if deletions == 0:
        return {"success": False, "error": f"Text not found: '{text}'"}

    _save_document_xml(filepath, root, zip_bytes)
    return {
        "success": True,
        "deletions": deletions,
        "message": f"Marked {deletions} occurrence(s) of '{text}' as deleted by {author}",
    }


def list_tracked_changes_in_doc(filepath: str) -> dict:
    """List all tracked changes in a document.

    Returns:
        Dict with lists of insertions and deletions
    """
    root, _ = _load_document_xml(filepath)

    insertions = []
    deletions = []

    for ins in root.iter(W("ins")):
        author = ins.get(W("author"), "Unknown")
        date = ins.get(W("date"), "")
        change_id = ins.get(W("id"), "")
        # Get inserted text
        texts = []
        for t in ins.iter(W("t")):
            if t.text:
                texts.append(t.text)
        text = "".join(texts)

        # Find paragraph context
        p = ins.getparent()
        while p is not None and p.tag != W("p"):
            p = p.getparent()
        para_text = _paragraph_text(p) if p is not None else ""

        insertions.append({
            "id": change_id,
            "author": author,
            "date": date,
            "text": text,
            "paragraph_context": para_text[:100],
        })

    for del_elem in root.iter(W("del")):
        author = del_elem.get(W("author"), "Unknown")
        date = del_elem.get(W("date"), "")
        change_id = del_elem.get(W("id"), "")
        # Get deleted text (from <w:delText>)
        texts = []
        for dt in del_elem.iter(W("delText")):
            if dt.text:
                texts.append(dt.text)
        # Also check <w:t> in case of malformed docs
        if not texts:
            for t in del_elem.iter(W("t")):
                if t.text:
                    texts.append(t.text)
        text = "".join(texts)

        p = del_elem.getparent()
        while p is not None and p.tag != W("p"):
            p = p.getparent()
        para_text = _paragraph_text(p) if p is not None else ""

        deletions.append({
            "id": change_id,
            "author": author,
            "date": date,
            "text": text,
            "paragraph_context": para_text[:100],
        })

    return {
        "success": True,
        "insertions": insertions,
        "deletions": deletions,
        "total_insertions": len(insertions),
        "total_deletions": len(deletions),
        "total_changes": len(insertions) + len(deletions),
    }


def accept_tracked_changes_in_doc(
    filepath: str,
    author: Optional[str] = None,
    change_ids: Optional[list[int]] = None,
) -> dict:
    """Accept tracked changes (apply insertions, remove deletions).

    Args:
        filepath: Path to .docx file
        author: If specified, only accept changes by this author
        change_ids: If specified, only accept changes with these IDs

    Returns:
        Dict with success status and count
    """
    root, zip_bytes = _load_document_xml(filepath)
    accepted = 0

    def _should_process(elem):
        if change_ids is not None:
            eid = elem.get(W("id"), "")
            try:
                if int(eid) not in change_ids:
                    return False
            except ValueError:
                return False
        if author is not None:
            if elem.get(W("author"), "") != author:
                return False
        return True

    # Accept insertions: unwrap <w:ins> (keep content)
    for ins in list(root.iter(W("ins"))):
        if not _should_process(ins):
            continue
        parent = ins.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ins)
        # Move children out of <w:ins>
        children = list(ins)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)
        parent.remove(ins)
        accepted += 1

    # Accept deletions: remove <w:del> and its content entirely
    for del_elem in list(root.iter(W("del"))):
        if not _should_process(del_elem):
            continue
        parent = del_elem.getparent()
        if parent is None:
            continue
        parent.remove(del_elem)
        accepted += 1

    # Also remove rPr/del markers (paragraph deletion markers)
    for rpr_del in list(root.iter(W("del"))):
        if not _should_process(rpr_del):
            continue
        parent = rpr_del.getparent()
        if parent is not None:
            parent.remove(rpr_del)
            accepted += 1

    if accepted == 0:
        return {"success": True, "message": "No matching tracked changes found to accept", "accepted": 0}

    _save_document_xml(filepath, root, zip_bytes)
    return {
        "success": True,
        "accepted": accepted,
        "message": f"Accepted {accepted} tracked change(s)",
    }


def reject_tracked_changes_in_doc(
    filepath: str,
    author: Optional[str] = None,
    change_ids: Optional[list[int]] = None,
) -> dict:
    """Reject tracked changes (remove insertions, restore deletions).

    Args:
        filepath: Path to .docx file
        author: If specified, only reject changes by this author
        change_ids: If specified, only reject changes with these IDs

    Returns:
        Dict with success status and count
    """
    root, zip_bytes = _load_document_xml(filepath)
    rejected = 0

    def _should_process(elem):
        if change_ids is not None:
            eid = elem.get(W("id"), "")
            try:
                if int(eid) not in change_ids:
                    return False
            except ValueError:
                return False
        if author is not None:
            if elem.get(W("author"), "") != author:
                return False
        return True

    # Reject insertions: remove <w:ins> and its content
    for ins in list(root.iter(W("ins"))):
        if not _should_process(ins):
            continue
        parent = ins.getparent()
        if parent is None:
            continue
        parent.remove(ins)
        rejected += 1

    # Reject deletions: unwrap <w:del>, convert delText→t (restore original text)
    for del_elem in list(root.iter(W("del"))):
        if not _should_process(del_elem):
            continue
        parent = del_elem.getparent()
        if parent is None:
            continue
        idx = list(parent).index(del_elem)
        children = list(del_elem)
        for i, child in enumerate(children):
            # Convert <w:delText> back to <w:t>
            for dt in child.iter(W("delText")):
                dt.tag = W("t")
            parent.insert(idx + i, child)
        parent.remove(del_elem)
        rejected += 1

    if rejected == 0:
        return {"success": True, "message": "No matching tracked changes found to reject", "rejected": 0}

    _save_document_xml(filepath, root, zip_bytes)
    return {
        "success": True,
        "rejected": rejected,
        "message": f"Rejected {rejected} tracked change(s)",
    }
