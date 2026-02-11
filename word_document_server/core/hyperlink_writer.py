"""
Core hyperlink functionality for Word documents.

Adds hyperlinks to .docx files by manipulating the underlying OOXML structure
directly (zip + lxml), since python-docx's hyperlink support is limited.
"""

import copy
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Optional

from lxml import etree

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

W = lambda tag: f"{{{WORD_NS}}}{tag}"

HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def _get_run_text(run: etree._Element) -> str:
    parts = []
    for t in run.findall(W("t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _get_run_rpr(run: etree._Element):
    rpr = run.find(W("rPr"))
    if rpr is not None:
        return copy.deepcopy(rpr)
    return None


def _find_text_in_paragraph(p: etree._Element, search_text: str):
    """Find where search_text appears across runs in a paragraph."""
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


def _load_zip_part(zip_bytes: bytes, part_name: str) -> Optional[bytes]:
    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        if part_name in zf.namelist():
            return zf.read(part_name)
    return None


def _get_next_rid(rels_root: etree._Element) -> int:
    max_rid = 0
    for rel in rels_root.iter("{%s}Relationship" % REL_NS):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return max_rid + 1


def add_hyperlink_to_doc(
    filepath: str,
    text: str,
    url: str,
    paragraph_index: Optional[int] = None,
) -> dict:
    """Add a hyperlink to existing text in a Word document.

    Finds the specified text and wraps it in a <w:hyperlink> element
    with blue underline formatting and an external relationship.

    Args:
        filepath: Path to .docx file
        text: Text in the document to convert to a hyperlink
        url: The URL the hyperlink should point to
        paragraph_index: If specified, only search in this paragraph (0-based)

    Returns:
        Dict with success status and details
    """
    filepath = Path(filepath)
    zip_bytes = filepath.read_bytes()

    # --- Load document.xml ---
    doc_xml_bytes = _load_zip_part(zip_bytes, "word/document.xml")
    if doc_xml_bytes is None:
        return {"success": False, "error": "Cannot find word/document.xml"}

    doc_root = etree.fromstring(doc_xml_bytes)
    body = doc_root.find(W("body"))
    if body is None:
        return {"success": False, "error": "Document has no body element"}

    paragraphs = body.findall(f".//{W('p')}")

    # Find target text
    match = None
    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return {"success": False, "error": f"Paragraph index {paragraph_index} out of range (0-{len(paragraphs)-1})"}
        match = _find_text_in_paragraph(paragraphs[paragraph_index], text)
    else:
        for p in paragraphs:
            match = _find_text_in_paragraph(p, text)
            if match is not None:
                break

    if match is None:
        return {"success": False, "error": f"Text not found: '{text}'"}

    # --- Load relationships ---
    rels_bytes = _load_zip_part(zip_bytes, "word/_rels/document.xml.rels")
    if rels_bytes is None:
        return {"success": False, "error": "Cannot find document.xml.rels"}

    rels_root = etree.fromstring(rels_bytes)
    rid = f"rId{_get_next_rid(rels_root)}"

    # Add hyperlink relationship
    new_rel = etree.SubElement(rels_root, "{%s}Relationship" % REL_NS)
    new_rel.set("Id", rid)
    new_rel.set("Type", HYPERLINK_REL_TYPE)
    new_rel.set("Target", url)
    new_rel.set("TargetMode", "External")

    # --- Build hyperlink element ---
    # Collect the matched text and formatting from matched runs
    first_run = match[0][0]
    first_start = match[0][1]
    last_run = match[-1][0]
    last_end = match[-1][2]
    first_run_text = _get_run_text(first_run)
    last_run_text = _get_run_text(last_run)

    parent = first_run.getparent()
    first_idx = list(parent).index(first_run)

    # Text before match in first run
    before_text = first_run_text[:first_start]
    # Text after match in last run
    after_text = last_run_text[last_end:]

    # Get formatting from first run
    rpr = _get_run_rpr(first_run)

    # Create hyperlink run with blue underline style
    hyperlink_elem = etree.Element(W("hyperlink"))
    hyperlink_elem.set("{%s}id" % R_NS, rid)

    h_run = etree.SubElement(hyperlink_elem, W("r"))
    h_rpr = etree.SubElement(h_run, W("rPr"))

    # Add hyperlink style
    h_style = etree.SubElement(h_rpr, W("rStyle"))
    h_style.set(W("val"), "Hyperlink")
    # Blue color
    h_color = etree.SubElement(h_rpr, W("color"))
    h_color.set(W("val"), "0563C1")
    h_color.set(W("themeColor"), "hyperlink")
    # Underline
    h_u = etree.SubElement(h_rpr, W("u"))
    h_u.set(W("val"), "single")

    # Copy other formatting from original (font, size, etc.) but not color/underline
    if rpr is not None:
        for child in rpr:
            tag_local = etree.QName(child).localname
            if tag_local not in ("color", "u", "rStyle"):
                h_rpr.append(copy.deepcopy(child))

    h_t = etree.SubElement(h_run, W("t"))
    h_t.text = text
    h_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Remove all matched runs from parent
    for run_elem, _, _ in match:
        run_parent = run_elem.getparent()
        if run_parent is not None:
            run_parent.remove(run_elem)

    # Re-insert: before_text_run, hyperlink, after_text_run
    insert_idx = first_idx
    offset = 0

    if before_text:
        before_run = etree.Element(W("r"))
        if rpr is not None:
            before_run.append(copy.deepcopy(rpr))
        bt = etree.SubElement(before_run, W("t"))
        bt.text = before_text
        bt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        parent.insert(insert_idx + offset, before_run)
        offset += 1

    parent.insert(insert_idx + offset, hyperlink_elem)
    offset += 1

    if after_text:
        after_rpr = _get_run_rpr(last_run) if last_run != first_run else rpr
        after_run = etree.Element(W("r"))
        if after_rpr is not None:
            after_run.append(copy.deepcopy(after_rpr))
        at = etree.SubElement(after_run, W("t"))
        at.text = after_text
        at.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        parent.insert(insert_idx + offset, after_run)

    # --- Serialize and write back ---
    new_doc_xml = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_rels_xml = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    buffer = BytesIO()
    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf_in:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename == "word/document.xml":
                    zf_out.writestr(item, new_doc_xml)
                elif item.filename == "word/_rels/document.xml.rels":
                    zf_out.writestr(item, new_rels_xml)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

    filepath.write_bytes(buffer.getvalue())

    return {
        "success": True,
        "text": text,
        "url": url,
        "relationship_id": rid,
        "message": f"Added hyperlink to '{text}' pointing to {url}",
    }
