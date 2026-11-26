"""Tests for MCP stdio handshake, schema hygiene, and stdout isolation."""

import json
import os
import subprocess
import sys

from word_document_server.core import protection, tables


def _run_stdio_handshake(extra_env: dict[str, str] | None = None) -> tuple[int, list[dict], str]:
    """Spawn the MCP server over stdio, run initialize + tools/list, and return (rc, json_lines, stderr)."""
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)

    proc = subprocess.Popen(
        [sys.executable, "-m", "word_document_server.main"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    assert proc.stdin is not None
    assert proc.stdout is not None
    assert proc.stderr is not None

    try:
        init_req = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "pytest", "version": "1.0"},
            },
        }
        proc.stdin.write((json.dumps(init_req) + "\n").encode())
        proc.stdin.flush()
        init_line = proc.stdout.readline().decode()

        proc.stdin.write((json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n").encode())
        proc.stdin.write((json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/list"}) + "\n").encode())
        proc.stdin.flush()
        tools_line = proc.stdout.readline().decode()

        proc.stdin.close()
        rc = proc.wait(timeout=10)
        remaining_out = proc.stdout.read().decode().strip()
        stderr = proc.stderr.read().decode(errors="replace")
    except Exception:
        proc.kill()
        proc.wait()
        raise

    assert not remaining_out, f"Unexpected trailing stdout data: {remaining_out!r}"
    messages = [json.loads(line) for line in (init_line, tools_line) if line.strip()]
    return rc, messages, stderr


def test_stdio_handshake_clean_exit_and_typed_schemas():
    """Server returns valid JSON-RPC on stdout, exits 0 on EOF, and exposes typed schemas."""
    rc, messages, stderr = _run_stdio_handshake({"MCP_ENABLE_LIVE_TOOLS": "true"})
    assert rc == 0, f"Expected exit code 0, got {rc}; stderr={stderr}"
    assert len(messages) == 2

    tools = messages[1]["result"]["tools"]
    assert len(tools) == 120
    by_name = {tool["name"]: tool for tool in tools}

    assert by_name["word_live_insert_paragraphs"]["inputSchema"]["properties"]["paragraphs"]["items"]["type"] == "string"
    assert by_name["word_live_add_table"]["inputSchema"]["properties"]["data"]["items"]["items"]["type"] == "string"
    assert by_name["word_live_format_table"]["inputSchema"]["properties"]["column_widths"]["items"]["type"] == "number"
    assert by_name["word_live_setup_heading_numbering"]["inputSchema"]["properties"]["h1_paragraphs"]["items"]["type"] == "integer"
    assert by_name["word_live_setup_heading_numbering"]["inputSchema"]["properties"]["heading_map"]["additionalProperties"]["type"] == "integer"
    assert by_name["word_live_apply_list"]["inputSchema"]["properties"]["number_format"]["additionalProperties"]["type"] == "string"
    assert by_name["word_live_apply_list"]["inputSchema"]["properties"]["start_at"]["additionalProperties"]["type"] == "integer"


def test_stdio_handshake_disable_live_tools():
    """Setting MCP_ENABLE_LIVE_TOOLS=false exposes only the 75 cross-platform tools."""
    rc, messages, stderr = _run_stdio_handshake({"MCP_ENABLE_LIVE_TOOLS": "false"})
    assert rc == 0, f"Expected exit code 0, got {rc}; stderr={stderr}"
    tools = messages[1]["result"]["tools"]
    names = {tool["name"] for tool in tools}
    assert len(tools) == 75
    assert "word_screen_capture" not in names
    assert not any(name.startswith("word_live_") for name in names)


def test_error_logs_go_to_stderr_not_stdout(capsys, tmp_path):
    """Error handlers in core modules must never pollute stdout."""
    assert tables.set_cell_shading(None, fill_color="FFFFFF") is False
    assert protection.add_protection_info(str(tmp_path / "no_dir" / "doc.docx"), "password", "hash") is False
    assert protection.add_protection_info(str(tmp_path / "missing.docx"), "password", "hash", raw_password="pw") is False

    captured = capsys.readouterr()
    assert captured.out == ""
    assert "Error setting cell shading" in captured.err
    assert "Protection error:" in captured.err
    assert "Encryption error:" in captured.err
