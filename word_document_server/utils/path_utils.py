"""Monkey-patch python-docx's PhysPkgReader to give clear errors for locked files.

When Word has a file open with an exclusive lock, python-docx's is_zipfile check
silently fails (returns False) because open() raises PermissionError. This patch
detects that scenario and gives a helpful error pointing to word_live_* tools.
"""

import os


def install_path_hook() -> None:
    """Patch PhysPkgReader.__new__ to detect locked files.

    Safe to call multiple times -- will not double-patch.
    """
    from docx.opc.phys_pkg import PhysPkgReader
    from docx.opc.exceptions import PackageNotFoundError

    if getattr(PhysPkgReader.__new__, "_locked_file_hooked", False):
        return

    _orig_new = PhysPkgReader.__new__

    def _patched_new(cls, pkg_file, *args, **kwargs):
        if isinstance(pkg_file, str) and not os.path.isdir(pkg_file):
            if not os.path.exists(pkg_file):
                raise PackageNotFoundError(
                    f"Package not found at '{pkg_file}'"
                )
            try:
                with open(pkg_file, "rb"):
                    pass  # Just test readability
            except PermissionError:
                raise PackageNotFoundError(
                    f"File locked (probably open in Word): '{pkg_file}'. "
                    f"Use word_live_* tools for Word-open files."
                )
        return _orig_new(cls, pkg_file, *args, **kwargs)

    _patched_new._locked_file_hooked = True
    PhysPkgReader.__new__ = _patched_new
