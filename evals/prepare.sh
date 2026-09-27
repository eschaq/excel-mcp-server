#!/usr/bin/env bash
# Generate the files the eval suite needs but the repository doesn't commit. Run once before
# `claude plugin eval` (and again after changing the generators), with the project's dev
# environment active so `python` has openpyxl.
#
# 1. evals/fixtures/eval_workbook.xlsx, built by make_eval_workbook.py. It's generated rather than
#    committed so the plugin ships no binary files.
# 2. The test copy of the plugin used by the security-read-only-mode case: the same server,
#    started with --read-only. It's generated so the repository holds exactly one plugin.json.
set -euo pipefail
evals="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

python "$evals/fixtures/make_eval_workbook.py"
echo "wrote $evals/fixtures/eval_workbook.xlsx"

dir="$evals/security-read-only-mode/readonly-plugin/.claude-plugin"
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
        "${CLAUDE_PLUGIN_ROOT}/../../../run_server.py", "--read-only"
      ]
    }
  }
}
JSON
echo "wrote $dir/plugin.json"
