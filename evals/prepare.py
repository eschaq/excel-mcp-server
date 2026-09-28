"""Generate the files the eval suite needs but the repository doesn't commit.

See "Evals" in the README for when to run it. It writes:
1. evals/fixtures/eval_workbook.xlsx, generated rather than committed so the plugin ships no
   binary files.
2. The read-only copy of the plugin used by the security-read-only-mode case, derived from the
   real plugin.json and generated so the repository holds exactly one plugin.json.
"""

import sys
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(FIXTURES))

import make_eval_workbook  # noqa: E402
import make_readonly_plugin  # noqa: E402


def main() -> None:
    workbook = FIXTURES / "eval_workbook.xlsx"
    make_eval_workbook.build(workbook)
    print(f"wrote {workbook}")
    make_readonly_plugin.main()


if __name__ == "__main__":
    main()
