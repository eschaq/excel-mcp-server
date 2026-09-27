"""Plugin entry point: start the MCP server from this plugin folder's own source.

The plugin starts this file through uv in frozen mode, which provides exactly the dependencies
pinned in uv.lock. The server's code is in src/excel_mcp/ next to this file and is imported from there,
so nothing outside the plugin folder is executed.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from excel_mcp.server import main  # noqa: E402

if __name__ == "__main__":
    main()
