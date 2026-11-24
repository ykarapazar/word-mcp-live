import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from word_document_server.utils.document_utils import find_and_replace_text


def test_simple_single_run_replace(tmp_path):
    doc = Document()
    p = doc.add_paragraph("Hello World")
    doc_path = str(tmp_path / "test1.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    count = find_and_replace_text(doc, "World", "Universe")
    assert count == 1
    assert doc.paragraphs[0].text == "Hello Universe"


def test_fragmented_multi_run_replace(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    r1 = p.add_run("The ")
    r2 = p.add_run("quick ")
    r3 = p.add_run("brown ")
    r4 = p.add_run("fox")
    doc_path = str(tmp_path / "test2.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    # Target phrase spans across r2, r3, r4
    count = find_and_replace_text(doc, "quick brown fox", "lazy dog")
    assert count == 1
    full_text = "".join(t.text for t in doc.paragraphs[0]._element.xpath(".//w:t") if t.text)
    assert full_text == "The lazy dog"


def test_table_cell_replace(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).paragraphs[0].text = "Header Cell"
    table.cell(1, 1).paragraphs[0].text = "Target Data"
    doc_path = str(tmp_path / "test3.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    count = find_and_replace_text(doc, "Target Data", "Updated Data")
    assert count == 1
    assert doc.tables[0].cell(1, 1).paragraphs[0].text == "Updated Data"


def test_textbox_replace(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    # Inject a Word DrawingML/txbxContent textbox into a paragraph
    textbox_xml = parse_xml(
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '  <w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        '             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        '             xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '    <wps:wsp>'
        '      <wps:txbx>'
        '        <w:txbxContent>'
        '          <w:p>'
        '            <w:r>'
        '              <w:t>Textbox Header Text</w:t>'
        '            </w:r>'
        '          </w:p>'
        '        </w:txbxContent>'
        '      </wps:txbx>'
        '    </wps:wsp>'
        '  </w:drawing>'
        '</w:r>'
    )
    p._element.append(textbox_xml)
    doc_path = str(tmp_path / "test4.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    count = find_and_replace_text(doc, "Textbox Header Text", "Replaced Textbox Content")
    assert count == 1
    all_t = [t.text for t in doc._element.xpath(".//w:txbxContent//w:t") if t.text]
    assert "Replaced Textbox Content" in all_t


def test_multiple_occurrences_in_single_paragraph(tmp_path):
    doc = Document()
    p = doc.add_paragraph("item and item and item")
    doc_path = str(tmp_path / "test5.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    count = find_and_replace_text(doc, "item", "element")
    assert count == 3
    assert doc.paragraphs[0].text == "element and element and element"


def test_skip_toc_paragraphs(tmp_path):
    doc = Document()
    p1 = doc.add_paragraph("TOC Header")
    # Set paragraph style to TOC 1 if available or inject pStyle
    p1_pPr = p1._element.get_or_add_pPr()
    pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="TOC 1"/>')
    p1_pPr.append(pStyle)

    p2 = doc.add_paragraph("Normal Header")

    doc_path = str(tmp_path / "test6.docx")
    doc.save(doc_path)

    doc = Document(doc_path)
    count = find_and_replace_text(doc, "Header", "Section")
    # Only p2 should be replaced, p1 should be skipped because it's a TOC style
    assert count == 1
    assert "Section" in doc.paragraphs[1].text
    assert "Header" in "".join(doc.paragraphs[0]._element.xpath(".//w:t/text()"))
