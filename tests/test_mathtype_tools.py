import pytest

from word_document_server.tools import mathtype_tools


def test_unified_equation_tools_delegate_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload)) or {"ok": True},
    )

    digest = "a" * 64
    assert mathtype_tools.word_live_list_equations("paper.docx") == {"ok": True}
    assert mathtype_tools.word_live_get_equation("omml:1.0:5:1") == {"ok": True}
    assert mathtype_tools.word_live_dump_equation_document(
        "dump.txt", "paper.docx"
    ) == {"ok": True}
    assert mathtype_tools.word_live_delete_equation("omml:1.0:5:1") == {"ok": True}
    assert mathtype_tools.word_live_replace_equation_tex(
        "omml:1.0:5:1", r"\frac{a}{b}", digest, filename="paper.docx"
    ) == {"ok": True}
    assert mathtype_tools.word_live_insert_equation_tex(
        r"\frac{a}{b}", 120, "display_numbered", filename="paper.docx"
    ) == {"ok": True}

    assert calls == [
        ("list_all_equations", {"filename": "paper.docx"}),
        (
            "get_any_equation",
            {"equation_id": "omml:1.0:5:1", "filename": None},
        ),
        (
            "dump_equation_document",
            {"output_path": "dump.txt", "filename": "paper.docx", "timeout": 600},
        ),
        (
            "delete_any_equation",
            {
                "equation_id": "omml:1.0:5:1",
                "revision_mode": "auto",
                "filename": None,
            },
        ),
        (
            "replace_any_equation_tex",
            {
                "equation_id": "omml:1.0:5:1",
                "tex": r"\frac{a}{b}",
                "expected_mathml_sha256": digest,
                "revision_mode": "auto",
                "filename": "paper.docx",
                "timeout": 120,
            },
        ),
        (
            "insert_equation_tex",
            {
                "tex": r"\frac{a}{b}",
                "range_start": 120,
                "layout": "display_numbered",
                "revision_mode": "auto",
                "filename": "paper.docx",
                "timeout": 120,
            },
        ),
    ]


def test_unified_replace_requires_expected_hash():
    with pytest.raises(ValueError, match="expected_mathml_sha256"):
        mathtype_tools.word_live_replace_equation_tex(
            "omml:1.0:5:1", "x", ""
        )


def test_list_mathtype_equations_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"equations": []},
    )

    result = mathtype_tools.word_live_list_mathtype_equations("paper.docx")

    assert result == {"equations": []}
    assert calls == [("list_equations", {"filename": "paper.docx"})]


def test_get_mathtype_equation_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"mathml": "<math />"},
    )

    result = mathtype_tools.word_live_get_mathtype_equation(
        "ole:1:42:Equation.DSMT4", "paper.docx"
    )

    assert result == {"mathml": "<math />"}
    assert calls == [
        (
            "get_equation",
            {
                "equation_id": "ole:1:42:Equation.DSMT4",
                "filename": "paper.docx",
            },
        )
    ]


def test_replace_mathtype_equation_requires_expected_hash(monkeypatch):
    monkeypatch.setattr(
        mathtype_tools, "invoke_bridge", lambda command, **payload: payload
    )

    with pytest.raises(ValueError, match="expected_mathml_sha256"):
        mathtype_tools.word_live_replace_mathtype_equation(
            "ole:1:42:Equation.DSMT4",
            '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>',
            "",
        )


def test_replace_mathtype_equation_sends_mathml_and_hash(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"verified": True},
    )
    mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>y</mi></math>'
    digest = "a" * 64

    result = mathtype_tools.word_live_replace_mathtype_equation(
        "ole:1:42:Equation.DSMT4", mathml, digest, "paper.docx"
    )

    assert result == {"verified": True}
    assert calls == [
        (
            "replace_equation",
            {
                "equation_id": "ole:1:42:Equation.DSMT4",
                "mathml": mathml,
                "expected_mathml_sha256": digest,
                "filename": "paper.docx",
            },
        )
    ]


def test_replace_tex_requires_expected_hash():
    with pytest.raises(ValueError, match="expected_mathml_sha256"):
        mathtype_tools.word_live_replace_mathtype_equation_tex(
            "ole:inline:1.0:10:1:Equation.DSMT4", "x^2", ""
        )


def test_replace_tex_passes_hash_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"equation_id": "new"},
    )

    digest = "a" * 64
    result = mathtype_tools.word_live_replace_mathtype_equation_tex(
        "eq-1", r"\alpha", digest, filename="paper.docx"
    )

    assert result == {"equation_id": "new"}
    assert calls == [
        (
            "replace_equation_tex",
            {
                "equation_id": "eq-1",
                "tex": r"\alpha",
                "expected_mathml_sha256": digest,
                "revision_mode": "auto",
                "filename": "paper.docx",
                "timeout": 120,
            },
        )
    ]


def test_probe_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload)) or {"formats": []},
    )

    result = mathtype_tools.word_live_probe_mathtype_equation("eq-1", "paper.docx")

    assert result == {"formats": []}
    assert calls == [("probe_equation", {"equation_id": "eq-1", "filename": "paper.docx"})]


def test_all_mathtype_tools_registered_with_docstring_description():
    import pathlib

    source = pathlib.Path("word_document_server/main.py").read_text(encoding="utf-8")
    for name in (
        "word_live_list_equations",
        "word_live_get_equation",
        "word_live_dump_equation_document",
        "word_live_delete_equation",
        "word_live_replace_equation_tex",
        "word_live_list_mathtype_equations",
        "word_live_get_mathtype_equation",
        "word_live_dump_mathtype_equations",
        "word_live_dump_mathtype_document",
        "word_live_delete_mathtype_equation",
        "word_live_replace_mathtype_equation_tex",
        "word_live_probe_mathtype_equation",
        "word_live_replace_mathtype_equation",
    ):
        assert f"def {name}(" in source, name
        assert f"description=mathtype_tools.{name}.__doc__" in source, name


def test_packaged_addin_scripts_exist():
    import pathlib

    bridge_dir = pathlib.Path("word_document_server/mathtype_bridge")
    for script in (
        "build_mathtype_bridge.ps1",
        "install_mathtype_addin.ps1",
        "uninstall_mathtype_addin.ps1",
    ):
        assert (bridge_dir / script).is_file(), script
