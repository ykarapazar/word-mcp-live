"""Tests for document resolution in word_document_server.core.word_com.

Regression cover for the wrong-document write bug: with two same-named
documents open in different folders, the old single-pass loop matched on
basename BEFORE full path, so asking for C:\\B\\Lettre.docx returned
C:\\A\\Lettre.docx — a silent write into the wrong file.

These tests use fake COM objects, so they run anywhere (no Word, no Windows).
"""
import os

import pytest

from word_document_server.core.word_com import (
    find_document,
    find_document_for_write,
    validate_live_filename,
)

DUPONT = r"C:\Clients\Dupont\Lettre.docx"
TREMBLAY = r"C:\Clients\Tremblay\Lettre.docx"


class FakeDoc:
    def __init__(self, full_name):
        self.FullName = full_name
        self.Name = os.path.basename(full_name)


class FakeDocuments:
    def __init__(self, docs):
        self._docs = docs

    @property
    def Count(self):
        return len(self._docs)

    def __call__(self, index):  # Word's Documents(i) is 1-based
        return self._docs[index - 1]


class FakeApp:
    def __init__(self, docs, active_index=0):
        self.Documents = FakeDocuments(docs)
        self.ActiveDocument = docs[active_index] if docs else None


@pytest.fixture
def ambiguous_app():
    """Two documents sharing the basename 'Lettre.docx'. Dupont opened first."""
    return FakeApp([FakeDoc(DUPONT), FakeDoc(TREMBLAY)])


# --- the core bug -----------------------------------------------------------

def test_full_path_wins_over_basename(ambiguous_app):
    """The regression. Old code returned DUPONT when asked for TREMBLAY."""
    assert find_document(ambiguous_app, TREMBLAY).FullName == TREMBLAY
    assert find_document(ambiguous_app, DUPONT).FullName == DUPONT


def test_ambiguous_basename_raises_and_lists_candidates(ambiguous_app):
    """A bare ambiguous basename must fail loudly, not silently pick the first."""
    with pytest.raises(ValueError) as exc:
        find_document(ambiguous_app, "Lettre.docx")
    msg = str(exc.value)
    assert "Ambiguous" in msg
    assert DUPONT in msg and TREMBLAY in msg  # both candidates surfaced


def test_unambiguous_basename_still_resolves():
    """Common case must not regress: one match by basename is fine."""
    app = FakeApp([FakeDoc(r"C:\Clients\Dupont\Requete.docx")])
    assert find_document(app, "Requete.docx").FullName == r"C:\Clients\Dupont\Requete.docx"


def test_matching_is_case_and_separator_insensitive(ambiguous_app):
    assert find_document(ambiguous_app, "C:/CLIENTS/tremblay/LETTRE.DOCX").FullName == TREMBLAY


def test_missing_document_raises(ambiguous_app):
    with pytest.raises(ValueError, match="is not open in Word"):
        find_document(ambiguous_app, r"C:\Clients\Other\Absent.docx")


def test_no_documents_open_raises():
    with pytest.raises(ValueError, match="No documents are open"):
        find_document(FakeApp([]), "x.docx")


# --- write guard ------------------------------------------------------------

def test_write_without_filename_is_refused(ambiguous_app):
    """A mutating call must never fall back to ActiveDocument implicitly."""
    with pytest.raises(ValueError, match="filename is required"):
        find_document_for_write(ambiguous_app, None)


def test_write_with_full_path_resolves_correctly(ambiguous_app):
    assert find_document_for_write(ambiguous_app, TREMBLAY).FullName == TREMBLAY


def test_read_without_filename_still_uses_active_document(ambiguous_app):
    """Read tools keep the permissive behaviour."""
    assert find_document(ambiguous_app, None).FullName == DUPONT


def test_escape_hatch_restores_implicit_active_writes(ambiguous_app, monkeypatch):
    monkeypatch.setenv("WORD_MCP_ALLOW_ACTIVE_WRITES", "1")
    assert find_document_for_write(ambiguous_app, None).FullName == DUPONT


# --- path hardening (upstream issue #15) ------------------------------------

@pytest.mark.parametrize("hostile", [
    r"..\..\Windows\System32\evil.docx",
    "../../etc/passwd.docx",
    r"\\evil-server\share\x.docx",
    "//evil-server/share/x.docx",
    "nul\x00byte.docx",
])
def test_hostile_paths_rejected(hostile):
    with pytest.raises(ValueError):
        validate_live_filename(hostile)


@pytest.mark.parametrize("benign", [
    "Lettre.docx",
    r"C:\Clients\Dupont\Lettre.docx",
    "C:/Clients/Dupont/Lettre.docx",
    None,
])
def test_benign_paths_accepted(benign):
    assert validate_live_filename(benign) == benign
