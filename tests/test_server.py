from pathlib import Path

import pytest

from excel_mcp.server import parse_config


@pytest.mark.parametrize("args, expected", [
    ([], False),
    (["--read-only"], True),
    (["--read-only=true"], True),
    (["--read-only=false"], False),
])
def test_read_only_flag(args, expected):
    config, _ = parse_config(args)
    assert config.read_only is expected


def test_environment_variables_are_ignored(monkeypatch):
    """Settings come only from flags, so nothing in the environment can change them."""
    monkeypatch.setenv("EXCEL_MCP_READ_ONLY", "1")
    monkeypatch.setenv("EXCEL_MCP_MAX_FILE_MB", "1")
    config, _ = parse_config([])
    assert config.read_only is False and config.max_file_mb == 50


def test_allowed_dirs_expand_home(tmp_path):
    config, _ = parse_config(["--allow-dir", "~", "--allow-dir", str(tmp_path)])
    assert config.allowed_dirs == [Path.home().resolve(), tmp_path.resolve()]


def test_missing_allowed_dir_is_an_error(tmp_path):
    with pytest.raises(SystemExit):
        parse_config(["--allow-dir", str(tmp_path / "nope")])


def test_plugin_launcher_starts_server(tmp_path):
    """run_server.py (what the plugin runs) must start the server from src/ with no PYTHONPATH."""
    import asyncio
    import os
    import sys

    from mcp import Client, StdioServerParameters

    root = Path(__file__).resolve().parent.parent
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    params = StdioServerParameters(
        command=sys.executable,
        args=[str(root / "run_server.py"), "--allow-dir", str(tmp_path), "--read-only",
              "--log-file", str(tmp_path / "test.log")],
        env=env,
        cwd=str(tmp_path),
    )

    async def go():
        async with Client(params) as client:
            return {t.name for t in (await client.list_tools()).tools}

    assert asyncio.run(go()) == {
        "read_spreadsheet", "get_sheet_data", "get_formulas", "search", "get_summary_stats", "apply_filter",
    }
