"""Process boundary for the Windows MathType OLE bridge."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import Any


class MathTypeBridgeError(RuntimeError):
    """A structured error returned by, or raised while invoking, the bridge."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


def bridge_executable_path() -> Path:
    return Path(__file__).resolve().parents[1] / "bin" / "MathTypeBridge.exe"


def invoke_bridge(
    command: str,
    *,
    executable: Path | str | None = None,
    timeout: int = 30,
    **payload: Any,
) -> dict[str, Any]:
    """Send one request to the native bridge and return its result object."""
    if os.name != "nt":
        raise MathTypeBridgeError(
            "windows_required",
            "MathType OLE access requires Windows desktop Word and MathType.",
        )

    executable_path = Path(executable) if executable else bridge_executable_path()
    if not executable_path.is_file():
        build_script = (
            Path(__file__).resolve().parents[1]
            / "mathtype_bridge"
            / "build_mathtype_bridge.ps1"
        )
        raise MathTypeBridgeError(
            "bridge_not_built",
            f"MathTypeBridge.exe is missing. Run: pwsh -File \"{build_script}\"",
        )

    request = {"command": command, **payload}
    try:
        completed = subprocess.run(
            [str(executable_path)],
            input=json.dumps(request, ensure_ascii=False),
            capture_output=True,
            encoding="utf-8",
            timeout=timeout,
            check=False,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired as exc:
        raise MathTypeBridgeError(
            "bridge_timeout", f"MathType bridge timed out after {timeout} seconds."
        ) from exc

    try:
        response = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        detail = completed.stderr.strip() or completed.stdout.strip() or "no output"
        raise MathTypeBridgeError(
            "invalid_bridge_response",
            f"MathType bridge returned invalid JSON: {detail}",
        ) from exc

    if not isinstance(response, dict) or not isinstance(response.get("ok"), bool):
        raise MathTypeBridgeError(
            "invalid_bridge_response", "MathType bridge response has an invalid schema."
        )

    if not response["ok"]:
        error = response.get("error") or {}
        raise MathTypeBridgeError(
            str(error.get("code") or "bridge_error"),
            str(error.get("message") or "MathType bridge failed."),
        )

    result = response.get("result")
    if not isinstance(result, dict):
        raise MathTypeBridgeError(
            "invalid_bridge_response", "MathType bridge result must be a JSON object."
        )
    return result
