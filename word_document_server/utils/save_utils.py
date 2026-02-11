"""Monkey-patch python-docx Document.save() to preserve custom XML parts.

Problem: python-docx strips parts it doesn't manage (comments.xml, etc.)
Solution: Before save, extract custom parts from the original file. After save,
re-inject them into the new file.
"""

import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

# Parts that python-docx strips on save
CUSTOM_PARTS_TO_PRESERVE = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
]

# Relationship types that accompany comments
COMMENT_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
}

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _extract_custom_parts(zip_bytes: bytes) -> dict | None:
    """Extract custom parts, relationships, and content-type overrides from a docx zip.

    Returns a dict with keys 'parts', 'rels', 'overrides', or None if nothing to preserve.
    """
    parts = {}
    rels = []
    overrides = []

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        namelist = zf.namelist()

        # 1. Extract custom part files
        for part_name in CUSTOM_PARTS_TO_PRESERVE:
            if part_name in namelist:
                parts[part_name] = zf.read(part_name)

        if not parts:
            return None  # Nothing to preserve

        # 2. Extract comment relationships from document.xml.rels
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in namelist:
            rels_root = etree.fromstring(zf.read(rels_path))
            for rel in rels_root.iter(f"{{{REL_NS}}}Relationship"):
                if rel.get("Type", "") in COMMENT_REL_TYPES:
                    rels.append({
                        "Id": rel.get("Id"),
                        "Type": rel.get("Type"),
                        "Target": rel.get("Target"),
                    })

        # 3. Extract comment-related content-type overrides
        if "[Content_Types].xml" in namelist:
            ct_root = etree.fromstring(zf.read("[Content_Types].xml"))
            for override in ct_root.iter(f"{{{CT_NS}}}Override"):
                part_name = override.get("PartName", "")
                if "comment" in part_name.lower():
                    overrides.append({
                        "PartName": part_name,
                        "ContentType": override.get("ContentType"),
                    })

    return {"parts": parts, "rels": rels, "overrides": overrides}


def _reinject_custom_parts(filepath: Path, preserved: dict) -> None:
    """Re-inject preserved custom parts into a saved docx file."""
    zip_bytes = filepath.read_bytes()

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf_in:
        existing_names = set(zf_in.namelist())

        # --- Patch document.xml.rels ---
        rels_path = "word/_rels/document.xml.rels"
        patched_rels_xml = None
        if preserved["rels"] and rels_path in existing_names:
            rels_root = etree.fromstring(zf_in.read(rels_path))

            # Collect existing rIds and rel types
            existing_rids = set()
            existing_types = set()
            for rel in rels_root.iter(f"{{{REL_NS}}}Relationship"):
                existing_rids.add(rel.get("Id", ""))
                existing_types.add(rel.get("Type", ""))

            for rel_info in preserved["rels"]:
                # Skip if this relationship type already exists (don't duplicate)
                if rel_info["Type"] in existing_types:
                    continue

                # Find a non-conflicting rId
                rid = rel_info["Id"]
                if rid in existing_rids:
                    # Generate a new rId
                    max_num = 0
                    for existing_rid in existing_rids:
                        if existing_rid.startswith("rId"):
                            try:
                                max_num = max(max_num, int(existing_rid[3:]))
                            except ValueError:
                                pass
                    rid = f"rId{max_num + 1}"

                new_rel = etree.SubElement(rels_root, f"{{{REL_NS}}}Relationship")
                new_rel.set("Id", rid)
                new_rel.set("Type", rel_info["Type"])
                new_rel.set("Target", rel_info["Target"])
                existing_rids.add(rid)
                existing_types.add(rel_info["Type"])

            patched_rels_xml = etree.tostring(
                rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        # --- Patch [Content_Types].xml ---
        patched_ct_xml = None
        if preserved["overrides"] and "[Content_Types].xml" in existing_names:
            ct_root = etree.fromstring(zf_in.read("[Content_Types].xml"))

            existing_part_names = set()
            for override in ct_root.iter(f"{{{CT_NS}}}Override"):
                existing_part_names.add(override.get("PartName", ""))

            for ov_info in preserved["overrides"]:
                if ov_info["PartName"] not in existing_part_names:
                    new_ov = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
                    new_ov.set("PartName", ov_info["PartName"])
                    new_ov.set("ContentType", ov_info["ContentType"])
                    existing_part_names.add(ov_info["PartName"])

            patched_ct_xml = etree.tostring(
                ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        # --- Rebuild the zip ---
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename == rels_path and patched_rels_xml is not None:
                    zf_out.writestr(item, patched_rels_xml)
                elif item.filename == "[Content_Types].xml" and patched_ct_xml is not None:
                    zf_out.writestr(item, patched_ct_xml)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

            # Re-add custom part files that python-docx stripped
            for part_name, part_bytes in preserved["parts"].items():
                if part_name not in existing_names:
                    zf_out.writestr(part_name, part_bytes)

    filepath.write_bytes(buffer.getvalue())


def install_save_hook() -> None:
    """Monkey-patch docx.document.Document.save to preserve custom XML parts.

    Only intercepts file-path saves (not stream saves).
    Safe to call multiple times — will not double-patch.
    """
    import docx.document

    # Guard against double-patching
    if hasattr(docx.document.Document.save, "_custom_parts_hooked"):
        return

    _original_save = docx.document.Document.save

    def _hooked_save(self, path_or_stream):
        # Only intercept when saving to a file path (str or Path)
        if isinstance(path_or_stream, (str, Path)):
            filepath = Path(path_or_stream)
            preserved = None

            # Extract custom parts from the existing file before python-docx overwrites it
            if filepath.exists():
                try:
                    preserved = _extract_custom_parts(filepath.read_bytes())
                except Exception:
                    preserved = None

            # Let python-docx do its normal save
            _original_save(self, path_or_stream)

            # Re-inject custom parts if we had any
            if preserved is not None:
                try:
                    _reinject_custom_parts(filepath, preserved)
                except Exception:
                    pass  # Fail silently — better to have a saved doc without comments than crash
        else:
            # Stream-based save — don't interfere
            _original_save(self, path_or_stream)

    _hooked_save._custom_parts_hooked = True
    docx.document.Document.save = _hooked_save
