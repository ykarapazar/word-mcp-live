"""
Core comment writing functionality for Word documents.

Adds comments to .docx files by manipulating the underlying OOXML structure
directly (zip + lxml), since python-docx doesn't support creating comments.

Reference: Anthropic docx skill scripts/comment.py
"""

import copy
import random
import zipfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Optional

from lxml import etree

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

W = lambda tag: f"{{{WORD_NS}}}{tag}"
W14 = lambda tag: f"{{{W14_NS}}}{tag}"

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

# Minimal comments.xml template (empty container)
COMMENTS_XML_TEMPLATE = b"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 wp14">
</w:comments>
"""


def _generate_hex_id() -> str:
    return f"{random.randint(0, 0x7FFFFFFE):08X}"


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _get_max_comment_id(comments_root: etree._Element) -> int:
    """Find the highest existing comment ID."""
    max_id = -1
    for comment in comments_root.iter(W("comment")):
        cid = comment.get(W("id"))
        if cid is not None:
            try:
                max_id = max(max_id, int(cid))
            except ValueError:
                pass
    return max_id


def _get_max_id_in_doc(root: etree._Element) -> int:
    """Find the highest w:id in the entire document.xml."""
    max_id = 0
    for elem in root.iter():
        val = elem.get(W("id"))
        if val is not None:
            try:
                max_id = max(max_id, int(val))
            except ValueError:
                pass
    return max_id


def _get_run_text(run: etree._Element) -> str:
    parts = []
    for t in run.findall(W("t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _find_text_in_paragraph(p: etree._Element, search_text: str):
    """Find where search_text appears across runs in a paragraph.
    Returns list of (run_element, start_offset, end_offset) or None.
    """
    runs = [r for r in p.findall(f".//{W('r')}") if r.find(W("t")) is not None]
    if not runs:
        return None

    char_map = []
    for ri, run in enumerate(runs):
        text = _get_run_text(run)
        for ci in range(len(text)):
            char_map.append((ri, ci))

    full_text = "".join(_get_run_text(r) for r in runs)
    pos = full_text.find(search_text)
    if pos == -1:
        return None

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


def _get_run_rpr(run: etree._Element):
    rpr = run.find(W("rPr"))
    if rpr is not None:
        return copy.deepcopy(rpr)
    return None


def _make_run(text: str, rpr=None) -> etree._Element:
    r = etree.Element(W("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, W("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _load_zip_part(zip_bytes: bytes, part_name: str) -> Optional[bytes]:
    """Load a single part from a docx zip, or None if missing."""
    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        if part_name in zf.namelist():
            return zf.read(part_name)
    return None


def _get_next_rid(rels_root: etree._Element) -> int:
    """Find the next available rId number."""
    max_rid = 0
    for rel in rels_root.iter("{%s}Relationship" % REL_NS):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return max_rid + 1


def _has_comments_rel(rels_root: etree._Element) -> bool:
    """Check if comments.xml relationship already exists."""
    for rel in rels_root.iter("{%s}Relationship" % REL_NS):
        if rel.get("Type", "") == COMMENTS_REL_TYPE:
            return True
    return False


def add_comment_to_doc(
    filepath: str,
    target_text: str,
    comment_text: str,
    author: str = "Av. Yüce Karapazar",
    initials: str = "AYK",
) -> dict:
    """Add a comment to a Word document anchored to specific text.

    Args:
        filepath: Path to .docx file
        target_text: Text in the document to attach the comment to
        comment_text: The comment content
        author: Comment author name
        initials: Author initials

    Returns:
        Dict with success status and details
    """
    filepath = Path(filepath)
    zip_bytes = filepath.read_bytes()

    # --- Load document.xml ---
    doc_xml_bytes = _load_zip_part(zip_bytes, "word/document.xml")
    if doc_xml_bytes is None:
        return {"success": False, "error": "Cannot find word/document.xml in the docx file"}

    doc_root = etree.fromstring(doc_xml_bytes)

    # --- Find target text in paragraphs ---
    body = doc_root.find(W("body"))
    if body is None:
        return {"success": False, "error": "Document has no body element"}

    paragraphs = body.findall(f".//{W('p')}")
    match = None
    match_para = None
    for p in paragraphs:
        match = _find_text_in_paragraph(p, target_text)
        if match is not None:
            match_para = p
            break

    if match is None:
        return {"success": False, "error": f"Target text not found: '{target_text}'"}

    # --- Load or create comments.xml ---
    comments_bytes = _load_zip_part(zip_bytes, "word/comments.xml")
    if comments_bytes is not None:
        comments_root = etree.fromstring(comments_bytes)
    else:
        comments_root = etree.fromstring(COMMENTS_XML_TEMPLATE)

    # --- Determine comment ID ---
    max_comment_id = _get_max_comment_id(comments_root)
    max_doc_id = _get_max_id_in_doc(doc_root)
    comment_id = max(max_comment_id, max_doc_id) + 1

    timestamp = _now_iso()
    para_id = _generate_hex_id()

    # --- Build comment element in comments.xml ---
    comment_elem = etree.SubElement(comments_root, W("comment"))
    comment_elem.set(W("id"), str(comment_id))
    comment_elem.set(W("author"), author)
    comment_elem.set(W("date"), timestamp)
    comment_elem.set(W("initials"), initials)

    # Comment paragraph
    cp = etree.SubElement(comment_elem, W("p"))
    cp.set(W14("paraId"), para_id)
    cp.set(W14("textId"), "77777777")

    # Annotation reference run
    ann_run = etree.SubElement(cp, W("r"))
    ann_rpr = etree.SubElement(ann_run, W("rPr"))
    ann_style = etree.SubElement(ann_rpr, W("rStyle"))
    ann_style.set(W("val"), "CommentReference")
    etree.SubElement(ann_run, W("annotationRef"))

    # Comment text run
    text_run = etree.SubElement(cp, W("r"))
    text_rpr = etree.SubElement(text_run, W("rPr"))
    sz = etree.SubElement(text_rpr, W("sz"))
    sz.set(W("val"), "20")
    szCs = etree.SubElement(text_rpr, W("szCs"))
    szCs.set(W("val"), "20")
    ct = etree.SubElement(text_run, W("t"))
    ct.text = comment_text
    ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # --- Inject markers into document.xml ---
    # We need:
    #   <w:commentRangeStart w:id="{comment_id}"/>
    #   ... target runs ...
    #   <w:commentRangeEnd w:id="{comment_id}"/>
    #   <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{comment_id}"/></w:r>

    first_run = match[0][0]
    first_start = match[0][1]
    last_run = match[-1][0]
    last_end = match[-1][2]

    parent = first_run.getparent()
    first_idx = list(parent).index(first_run)
    last_idx = list(parent).index(last_run)

    # If the match starts mid-run, split the first run
    first_run_text = _get_run_text(first_run)
    if first_start > 0:
        before_text = first_run_text[:first_start]
        after_text = first_run_text[first_start:]
        rpr = _get_run_rpr(first_run)

        before_run = _make_run(before_text, rpr)
        parent.insert(first_idx, before_run)

        # Update the first run to only contain the matched part
        for t_elem in first_run.findall(W("t")):
            first_run.remove(t_elem)
        new_t = etree.SubElement(first_run, W("t"))
        new_t.text = after_text
        new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        first_idx += 1  # Account for the inserted before_run
        last_idx = list(parent).index(last_run)

    # If the match ends mid-run, split the last run
    last_run_text = _get_run_text(last_run)
    if last_run == first_run:
        # Recalculate because we may have modified first_run
        last_run_text = _get_run_text(last_run)
        effective_end = last_end - first_start if first_start > 0 else last_end
    else:
        effective_end = last_end

    if effective_end < len(last_run_text):
        matched_text = last_run_text[:effective_end]
        remainder_text = last_run_text[effective_end:]
        rpr = _get_run_rpr(last_run)

        for t_elem in last_run.findall(W("t")):
            last_run.remove(t_elem)
        new_t = etree.SubElement(last_run, W("t"))
        new_t.text = matched_text
        new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        last_idx = list(parent).index(last_run)
        remainder_run = _make_run(remainder_text, rpr)
        parent.insert(last_idx + 1, remainder_run)

    # Now insert commentRangeStart before first matched run
    first_idx = list(parent).index(first_run)
    range_start = etree.Element(W("commentRangeStart"))
    range_start.set(W("id"), str(comment_id))
    parent.insert(first_idx, range_start)

    # Insert commentRangeEnd after last matched run
    last_idx = list(parent).index(last_run)
    range_end = etree.Element(W("commentRangeEnd"))
    range_end.set(W("id"), str(comment_id))
    parent.insert(last_idx + 1, range_end)

    # Insert commentReference run after commentRangeEnd
    ref_run = etree.Element(W("r"))
    ref_rpr = etree.SubElement(ref_run, W("rPr"))
    ref_style = etree.SubElement(ref_rpr, W("rStyle"))
    ref_style.set(W("val"), "CommentReference")
    ref_elem = etree.SubElement(ref_run, W("commentReference"))
    ref_elem.set(W("id"), str(comment_id))
    end_idx = list(parent).index(range_end)
    parent.insert(end_idx + 1, ref_run)

    # --- Handle relationships ---
    rels_bytes = _load_zip_part(zip_bytes, "word/_rels/document.xml.rels")
    rels_modified = False
    if rels_bytes is not None:
        rels_root = etree.fromstring(rels_bytes)
        if not _has_comments_rel(rels_root):
            next_rid = _get_next_rid(rels_root)
            new_rel = etree.SubElement(rels_root, "{%s}Relationship" % REL_NS)
            new_rel.set("Id", f"rId{next_rid}")
            new_rel.set("Type", COMMENTS_REL_TYPE)
            new_rel.set("Target", "comments.xml")
            rels_modified = True
    else:
        # Very unlikely but handle: no rels file at all
        rels_root = None

    # --- Handle Content_Types ---
    ct_bytes = _load_zip_part(zip_bytes, "[Content_Types].xml")
    ct_modified = False
    if ct_bytes is not None:
        ct_root = etree.fromstring(ct_bytes)
        has_ct = False
        for override in ct_root.iter("{%s}Override" % CT_NS):
            if override.get("PartName") == "/word/comments.xml":
                has_ct = True
                break
        if not has_ct:
            new_override = etree.SubElement(ct_root, "{%s}Override" % CT_NS)
            new_override.set("PartName", "/word/comments.xml")
            new_override.set("ContentType",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
            ct_modified = True

    # --- Serialize and write back ---
    new_doc_xml = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_comments_xml = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    new_rels_xml = None
    if rels_modified and rels_root is not None:
        new_rels_xml = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    new_ct_xml = None
    if ct_modified and ct_root is not None:
        new_ct_xml = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    buffer = BytesIO()
    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf_in:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
            comments_written = False
            for item in zf_in.infolist():
                if item.filename == "word/document.xml":
                    zf_out.writestr(item, new_doc_xml)
                elif item.filename == "word/comments.xml":
                    zf_out.writestr(item, new_comments_xml)
                    comments_written = True
                elif item.filename == "word/_rels/document.xml.rels" and new_rels_xml is not None:
                    zf_out.writestr(item, new_rels_xml)
                elif item.filename == "[Content_Types].xml" and new_ct_xml is not None:
                    zf_out.writestr(item, new_ct_xml)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

            # If comments.xml didn't exist before, add it
            if not comments_written:
                zf_out.writestr("word/comments.xml", new_comments_xml)

    filepath.write_bytes(buffer.getvalue())

    return {
        "success": True,
        "comment_id": comment_id,
        "author": author,
        "target_text": target_text,
        "comment_text": comment_text,
        "message": f"Added comment #{comment_id} by {author} on text '{target_text[:50]}{'...' if len(target_text) > 50 else ''}'",
    }
