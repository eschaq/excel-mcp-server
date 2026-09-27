"""Entry point: ``python -m excel_mcp.server`` or the ``dws-spreadsheet-mcp`` script.

Runs over stdio only; the server opens no network connections.

Settings come only from command-line flags; the server reads no environment variables:
  --allow-dir DIR      directory files may live in, repeatable (default: home dir)
  --read-only[=BOOL]   disable and hide all write tools
  --max-file-mb N      refuse files larger than this (default: 50)
  --log-file PATH      operation log (default: ~/.excel-mcp/excel-mcp.log)
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from excel_mcp.excel_ops import Config
from excel_mcp.tools import build_server


def _truthy(v: str | None) -> bool:
    return (v or "").strip().lower() in {"1", "true", "yes", "on"}


def parse_config(argv: list[str] | None = None) -> tuple[Config, Path]:
    parser = argparse.ArgumentParser(prog="dws-spreadsheet-mcp", description="Excel MCP server (stdio)")
    parser.add_argument("--allow-dir", action="append", dest="allow_dirs", metavar="DIR")
    parser.add_argument("--read-only", nargs="?", const="true", default="", metavar="BOOL")
    parser.add_argument("--max-file-mb", type=float)
    parser.add_argument("--log-file")
    args = parser.parse_args(argv)

    dirs = [Path(d).expanduser() for d in args.allow_dirs or []] or [Path.home()]
    for d in dirs:
        if not d.expanduser().is_dir():
            parser.error(f"allowed directory does not exist: {d}")

    config = Config(
        allowed_dirs=dirs,
        read_only=_truthy(args.read_only),
        max_file_mb=args.max_file_mb or 50.0,
    )
    log_file = Path(args.log_file) if args.log_file else Path.home() / ".excel-mcp" / "excel-mcp.log"
    return config, log_file.expanduser()


def setup_logging(log_file: Path) -> None:
    # stdout carries the MCP protocol, so logs go to stderr (shown in Claude Desktop's
    # MCP log) and to a file for an audit trail of every file operation.
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stderr)]
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    except OSError as e:
        print(f"excel-mcp: cannot open log file {log_file}: {e}", file=sys.stderr)
    logger = logging.getLogger("excel_mcp")
    logger.setLevel(logging.INFO)
    logger.propagate = False  # the SDK configures the root logger too; avoid double lines
    for h in handlers:
        h.setFormatter(fmt)
        logger.addHandler(h)


def main(argv: list[str] | None = None) -> None:
    config, log_file = parse_config(argv)
    setup_logging(log_file)
    logging.getLogger("excel_mcp").info(
        "starting: allowed_dirs=%s read_only=%s max_file_mb=%g",
        [str(d) for d in config.allowed_dirs], config.read_only, config.max_file_mb,
    )
    build_server(config).run("stdio")


if __name__ == "__main__":
    main()
