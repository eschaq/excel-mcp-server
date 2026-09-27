import shutil
from pathlib import Path

import pytest

from excel_mcp.excel_ops import Config, ExcelOps

FIXTURES = Path(__file__).parent / "fixtures"


@pytest.fixture
def workdir(tmp_path: Path) -> Path:
    """A sandbox directory holding fresh copies of the fixture files."""
    for name in ("sample.xlsx", "sample.xls"):
        shutil.copy(FIXTURES / name, tmp_path / name)
    (tmp_path / "sales.csv").write_text(
        "Region,Units,Price\nWest,10,2.5\nEast,4,10\nWest,7,6\n", encoding="utf-8"
    )
    return tmp_path


@pytest.fixture
def ops(workdir: Path) -> ExcelOps:
    return ExcelOps(Config(allowed_dirs=[workdir]))


@pytest.fixture
def ro_ops(workdir: Path) -> ExcelOps:
    return ExcelOps(Config(allowed_dirs=[workdir], read_only=True))
