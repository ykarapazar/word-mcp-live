import threading
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

from word_document_server.core.tracked_changes import track_replace_in_doc

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _make_docx(path: Path, text: str) -> None:
    doc = Document()
    doc.add_paragraph(text)
    doc.save(path)


def _replace(path: Path, old_text: str, new_text: str) -> dict:
    """Run track_replace_in_doc in a daemon thread so that a hang fails the test instead of blocking it."""
    result = {}
    worker = threading.Thread(
        target=lambda: result.update(track_replace_in_doc(str(path), old_text, new_text, author="Tester")),
        daemon=True,
    )
    worker.start()
    worker.join(timeout=10)
    assert not worker.is_alive(), "track_replace_in_doc did not return (infinite loop)"
    return result


def _texts(path: Path) -> tuple[str, str, str]:
    """Visible text (w:t), deleted text (w:delText) and inserted text (w:ins) of the first paragraph."""
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    p = root.find(f".//{{{W_NS}}}p")
    visible = "".join(p.xpath(".//w:t/text()", namespaces={"w": W_NS}))
    deleted = "".join(p.xpath(".//w:delText/text()", namespaces={"w": W_NS}))
    inserted = "".join(p.xpath(".//w:ins//w:t/text()", namespaces={"w": W_NS}))
    return visible, deleted, inserted


def test_new_text_containing_old_text(tmp_path: Path):
    path = tmp_path / "doc.docx"
    _make_docx(path, "The amount shall be paid.")

    result = _replace(path, "shall be paid.", "shall be paid. A cap applies.")

    assert result["success"] and result["replacements"] == 1
    assert _texts(path) == (
        "The amount shall be paid. A cap applies.",
        "shall be paid.",
        "shall be paid. A cap applies.",
    )


def test_new_text_containing_old_text_twice_in_paragraph(tmp_path: Path):
    path = tmp_path / "doc.docx"
    _make_docx(path, "Pay A. Pay B.")

    result = _replace(path, "Pay", "Pay now")

    assert result["success"] and result["replacements"] == 2
    assert _texts(path)[0] == "Pay now A. Pay now B."


def test_all_occurrences_replaced(tmp_path: Path):
    path = tmp_path / "doc.docx"
    _make_docx(path, "cat, cat and cat")

    result = _replace(path, "cat", "dog")

    assert result["success"] and result["replacements"] == 3
    assert _texts(path)[0] == "dog, dog and dog"
