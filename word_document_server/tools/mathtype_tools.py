"""Live tools for semantic access to real MathType OLE equations."""

from word_document_server.core.mathtype_bridge import invoke_bridge


def word_live_list_equations(filename: str | None = None) -> dict:
    """List MathType OLE and Word OMML equations through one live interface."""
    return invoke_bridge("list_all_equations", filename=filename)


def word_live_get_equation(equation_id: str, filename: str | None = None) -> dict:
    """Read one MathType OLE or Word OMML equation as MathML plus a hash."""
    return invoke_bridge(
        "get_any_equation", equation_id=equation_id, filename=filename
    )


def word_live_dump_equation_document(
    output_path: str, filename: str | None = None
) -> dict:
    """Dump document text and every MathType/OMML equation in one agent read."""
    return invoke_bridge(
        "dump_equation_document",
        output_path=output_path,
        filename=filename,
        timeout=600,
    )


def word_live_delete_equation(
    equation_id: str,
    revision_mode: str = "auto",
    filename: str | None = None,
) -> dict:
    """Delete one MathType OLE or Word OMML equation in one undo record.

    revision_mode: "auto" follows the document's Track Changes state (a tracked
    document keeps the equation visible as a pending deletion), "track" forces a
    tracked revision, "suppress" applies the edit as if already approved.
    """
    return invoke_bridge(
        "delete_any_equation",
        equation_id=equation_id,
        revision_mode=revision_mode,
        filename=filename,
    )


def word_live_replace_equation_tex(
    equation_id: str,
    tex: str,
    expected_mathml_sha256: str,
    revision_mode: str = "auto",
    filename: str | None = None,
) -> dict:
    """Replace a MathType OLE or Word OMML equation from TeX through one interface.

    The equation id selects the correct storage engine. The required hash prevents
    stale writes; failures restore and verify the original equation with Word Undo.
    revision_mode: "auto" follows the document's Track Changes state (pending
    revisions await manual approval), "track" forces tracking, "suppress" applies
    the edit as if already approved. The result's revisions_pending flag reports
    whether an approval step is outstanding.
    """
    if not expected_mathml_sha256:
        raise ValueError("expected_mathml_sha256 is required; call the get tool first")
    return invoke_bridge(
        "replace_any_equation_tex",
        equation_id=equation_id,
        tex=tex,
        expected_mathml_sha256=expected_mathml_sha256,
        revision_mode=revision_mode,
        filename=filename,
        timeout=120,
    )


def word_live_insert_equation_tex(
    tex: str,
    range_start: int,
    layout: str = "inline",
    revision_mode: str = "auto",
    filename: str | None = None,
) -> dict:
    """Insert a new MathType equation from TeX at a main-story character offset.

    Offsets come from the dump tools. layout "inline" places the equation in the
    surrounding text; "display" gives it its own centered paragraph;
    "display_numbered" uses MathType's tabbed display layout with a right-side
    SEQ MTEqn number. revision_mode is as in the replace tool.
    """
    return invoke_bridge(
        "insert_equation_tex",
        tex=tex,
        range_start=range_start,
        layout=layout,
        revision_mode=revision_mode,
        filename=filename,
        timeout=120,
    )


def word_live_list_mathtype_equations(filename: str | None = None) -> dict:
    """List MathType OLE equations in an open Word document without editing it."""
    return invoke_bridge("list_equations", filename=filename)


def word_live_get_mathtype_equation(
    equation_id: str, filename: str | None = None
) -> dict:
    """Read one MathType OLE equation as validated MathML and a concurrency hash."""
    return invoke_bridge(
        "get_equation", equation_id=equation_id, filename=filename
    )


def word_live_dump_mathtype_equations(
    output_path: str, filename: str | None = None
) -> dict:
    """Dump every MathType equation (id, context, canonical MathML) to a text file.

    One bridge call for the whole document; read the file with offset/grep
    instead of calling the get tool once per equation.
    """
    return invoke_bridge(
        "dump_equations", output_path=output_path, filename=filename, timeout=600
    )


def word_live_dump_mathtype_document(
    output_path: str, filename: str | None = None
) -> dict:
    """Dump the full main-story text with inline [eq N] markers plus a MathML appendix.

    One call gives an agent the whole document and every MathType equation's
    position, id, and canonical MathML.
    """
    return invoke_bridge(
        "dump_document", output_path=output_path, filename=filename, timeout=600
    )


def word_live_delete_mathtype_equation(
    equation_id: str,
    revision_mode: str = "auto",
    filename: str | None = None,
) -> dict:
    """Delete one MathType equation inside a single Word undo record.

    revision_mode: "auto" follows the document's Track Changes state, "track"
    forces a tracked revision, "suppress" applies the edit as already approved.
    """
    return invoke_bridge(
        "delete_equation",
        equation_id=equation_id,
        revision_mode=revision_mode,
        filename=filename,
    )


def word_live_replace_mathtype_equation_tex(
    equation_id: str,
    tex: str,
    expected_mathml_sha256: str,
    revision_mode: str = "auto",
    filename: str | None = None,
) -> dict:
    """Replace one MathType equation from TeX via MathType's Toggle TeX (no popups).

    Requires the equation's current mathml_sha256 (from the get/dump tools) so a
    stale request cannot overwrite a newer edit. Deletes the old OLE object,
    inserts $tex$ at the same position, converts it back to a MathType equation,
    and returns the new equation_id plus read-back MathML for verification. On
    any failure the original equation is restored. The returned equation_id
    differs from the input. revision_mode: "auto" follows the document's Track
    Changes state, "track" forces a tracked revision, "suppress" applies the
    edit as already approved.
    """
    if not expected_mathml_sha256:
        raise ValueError("expected_mathml_sha256 is required; call the get tool first")
    return invoke_bridge(
        "replace_equation_tex",
        equation_id=equation_id,
        tex=tex,
        expected_mathml_sha256=expected_mathml_sha256,
        revision_mode=revision_mode,
        filename=filename,
        timeout=120,
    )


def word_live_probe_mathtype_equation(
    equation_id: str, filename: str | None = None
) -> dict:
    """Enumerate the OLE clipboard formats one MathType equation offers (diagnostic)."""
    return invoke_bridge(
        "probe_equation", equation_id=equation_id, filename=filename
    )


def word_live_replace_mathtype_equation(
    equation_id: str,
    mathml: str,
    expected_mathml_sha256: str,
    filename: str | None = None,
) -> dict:
    """Replace one MathType equation and verify the saved MathML by reading it back."""
    if not expected_mathml_sha256:
        raise ValueError("expected_mathml_sha256 is required; call the get tool first")
    return invoke_bridge(
        "replace_equation",
        equation_id=equation_id,
        mathml=mathml,
        expected_mathml_sha256=expected_mathml_sha256,
        filename=filename,
    )
