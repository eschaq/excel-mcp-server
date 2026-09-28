"""Generate the files the eval suite needs but the repository doesn't commit.

Run once before `claude plugin eval` (and again after changing the generators), using the
project's dev environment so openpyxl is available:

    python evals/prepare.py

It writes:
1. evals/fixtures/eval_workbook.xlsx, generated rather than committed so the plugin ships no
   binary files.
2. The read-only copy of the plugin used by the security-read-only-mode case, derived from the
   real plugin.json and generated so the repository holds exactly one plugin.json.
"""

import importlib.util
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, FIXTURES / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    workbook = FIXTURES / "eval_workbook.xlsx"
    _load("make_eval_workbook").build(workbook)
    print(f"wrote {workbook}")
    _load("make_readonly_plugin").main()


if __name__ == "__main__":
    main()
