from pathlib import Path

import pytest

from excel_mcp.server import parse_config


@pytest.fixture(autouse=True)
def clean_env(monkeypatch):
    for k in ("EXCEL_MCP_ALLOWED_DIRS", "EXCEL_MCP_READ_ONLY", "EXCEL_MCP_MAX_FILE_MB", "EXCEL_MCP_LOG_FILE"):
        monkeypatch.delenv(k, raising=False)


@pytest.mark.parametrize("args, env, expected", [
    ([], None, False),
    (["--read-only"], None, True),
    (["--read-only=true"], None, True),
    (["--read-only=false"], None, False),
    (["--read-only=false"], "1", True),   # env can lock it down...
    ([], "true", True),
    (["--read-only=true"], "0", True),    # ...but can't unlock it
])
def test_read_only_sources(monkeypatch, args, env, expected):
    if env is not None:
        monkeypatch.setenv("EXCEL_MCP_READ_ONLY", env)
    config, _ = parse_config(args)
    assert config.read_only is expected


def test_allowed_dirs_expand_home(tmp_path):
    config, _ = parse_config(["--allow-dir", "~", "--allow-dir", str(tmp_path)])
    assert config.allowed_dirs == [Path.home().resolve(), tmp_path.resolve()]


def test_missing_allowed_dir_is_an_error(tmp_path):
    with pytest.raises(SystemExit):
        parse_config(["--allow-dir", str(tmp_path / "nope")])
