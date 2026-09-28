#!/usr/bin/env bash
# Generate the files the eval suite needs but the repository doesn't commit. Run once before
# `claude plugin eval` (and again after changing the generators), with the project's dev
# environment active so `python` has openpyxl.
#
# 1. evals/fixtures/eval_workbook.xlsx: generated rather than committed so the plugin ships no
#    binary files.
# 2. The read-only copy of the plugin used by the security-read-only-mode case, derived from the
#    real plugin.json: generated so the repository holds exactly one plugin.json.
set -euo pipefail
evals="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

python "$evals/fixtures/make_eval_workbook.py"
echo "wrote $evals/fixtures/eval_workbook.xlsx"
python "$evals/fixtures/make_readonly_plugin.py"
