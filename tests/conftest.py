import importlib.util
import shutil
from pathlib import Path

import pytest

from excel_mcp.excel_ops import Config, ExcelOps

FIXTURES = Path(__file__).parent / "fixtures"


def _load_generator():
    spec = importlib.util.spec_from_file_location("make_sample", FIXTURES / "make_sample.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="session")
def fixture_files(tmp_path_factory) -> Path:
    """Build the sample workbooks once per test session (they aren't committed)."""
    gen = _load_generator()
    d = tmp_path_factory.mktemp("fixtures")
    gen.build(d / "sample.xlsx")
    gen.build_xls(d / "sample.xls")
    return d


@pytest.fixture
def workdir(tmp_path: Path, fixture_files: Path) -> Path:
    """A sandbox directory holding fresh copies of the fixture files."""
    for name in ("sample.xlsx", "sample.xls"):
        shutil.copy(fixture_files / name, tmp_path / name)
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
