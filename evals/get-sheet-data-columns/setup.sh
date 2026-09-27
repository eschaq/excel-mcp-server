#!/usr/bin/env bash
# Copy the shared eval workbook into the run's empty workspace.
set -euo pipefail
cp "$(dirname "${BASH_SOURCE[0]}")/../fixtures/eval_workbook.xlsx" .
