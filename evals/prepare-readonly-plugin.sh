#!/usr/bin/env bash
# Generate the test copy of the plugin used by the security-read-only-mode eval case.
#
# It is the same server as the real plugin, started with --read-only. It's generated (and
# gitignored) rather than committed so the repository holds exactly one plugin.json.
# Run this once before `claude plugin eval`.
set -euo pipefail
dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/security-read-only-mode/readonly-plugin/.claude-plugin"
mkdir -p "$dir"
cat > "$dir/plugin.json" <<'JSON'
{
  "name": "dws-spreadsheet-readonly",
  "version": "0.1.0",
  "description": "Eval-only copy of dws-spreadsheet whose server runs with --read-only.",
  "author": { "name": "DWS Build", "email": "datawisdomsolutions@gmail.com" },
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run", "--frozen", "--project", "${CLAUDE_PLUGIN_ROOT}/../../..",
        "python", "-m", "excel_mcp.server", "--read-only"
      ],
      "env": {
        "PYTHONPATH": "${CLAUDE_PLUGIN_ROOT}/../../../src",
        "UV_PROJECT_ENVIRONMENT": "${CLAUDE_PLUGIN_DATA}/venv"
      }
    }
  }
}
JSON
echo "wrote $dir/plugin.json"
