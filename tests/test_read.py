import pytest

from excel_mcp.excel_ops import Condition, Config, ExcelOps, ExcelOpsError


# ---- read_spreadsheet ------------------------------------------------------ #

def test_read_spreadsheet_lists_sheets_and_dimensions(ops):
    info = ops.read_spreadsheet("sample.xlsx", preview_rows=2)
    assert info["format"] == "xlsx"
    assert [s["name"] for s in info["sheets"]] == ["Sales", "Inventory"]
    sales = info["sheets"][0]
    assert (sales["rows"], sales["columns"], sales["used_range"]) == (9, 7, "A1:G9")
    assert sales["preview"][0][:3] == ["Region", "Rep", "Product"]
    assert sales["preview"][1][5] == "2026-01-05"  # midnight datetimes render as dates
    assert len(sales["preview"]) == 2


def test_read_spreadsheet_csv(ops):
    info = ops.read_spreadsheet("sales.csv")
    assert info["sheets"][0]["name"] == "sales"
    assert info["sheets"][0]["preview"][1] == ["West", 10, 2.5]


def test_read_spreadsheet_legacy_xls(ops):
    info = ops.read_spreadsheet("sample.xls")
    sales = info["sheets"][0]
    assert sales["name"] == "Sales"
    assert sales["preview"][1][:4] == ["West", "Alice", "Widget", 10]
    assert sales["preview"][1][5] == "2026-01-05"


# ---- get_sheet_data -------------------------------------------------------- #

def test_get_sheet_data_records_carry_row_numbers(ops):
    data = ops.get_sheet_data("sample.xlsx")
    assert data["columns"][0] == "Region"
    assert data["total_rows"] == 7  # the blank row 8 is skipped; the Total row is kept
    first = data["rows"][0]
    assert first["_row"] == 2 and first["Rep"] == "Alice" and first["Units"] == 10
    assert data["rows"][-1]["_row"] == 9


def test_get_sheet_data_columns_and_paging(ops):
    data = ops.get_sheet_data("sample.xlsx", sheet="sales", columns=["rep", "Units"], offset=1, limit=2)
    assert data["columns"] == ["Rep", "Units"]
    assert data["rows"] == [{"_row": 3, "Rep": "Bob", "Units": 4}, {"_row": 4, "Rep": "Carol", "Units": 7}]
    assert data["truncated"] is True


def test_get_sheet_data_without_header(ops):
    data = ops.get_sheet_data("sample.xlsx", sheet="Inventory", header_row=0)
    assert data["columns"] == ["A", "B", "C"]
    assert data["rows"][0] == {"_row": 1, "A": "SKU", "B": "Product", "C": "In Stock"}


def test_get_sheet_data_unknown_sheet_and_column(ops):
    with pytest.raises(ExcelOpsError, match="Available sheets"):
        ops.get_sheet_data("sample.xlsx", sheet="Nope")
    with pytest.raises(ExcelOpsError, match="Available columns"):
        ops.get_sheet_data("sample.xlsx", columns=["Nope"])


# ---- get_summary_stats ----------------------------------------------------- #

def test_summary_stats(ops):
    stats = ops.get_summary_stats("sales.csv")
    units = stats["numeric_columns"]["Units"]
    assert units["count"] == 3 and units["min"] == 4 and units["max"] == 10
    assert units["mean"] == pytest.approx(7.0)
    assert units["median"] == 7
    assert units["stdev"] == pytest.approx(3.0)
    assert units["sum"] == 21
    assert stats["non_numeric_columns"]["Region"] == {
        "count": 3, "missing": 0, "unique": 2, "top": "West", "top_count": 2,
    }


def test_summary_stats_missing_values_and_dates(ops):
    stats = ops.get_summary_stats("sample.xlsx", columns=["Units", "Date"])
    assert stats["numeric_columns"]["Units"]["missing"] == 2  # Erin's blank + the Total row formula
    assert stats["non_numeric_columns"]["Date"] == {
        "count": 6, "missing": 1, "earliest": "2026-01-05", "latest": "2026-03-20",
    }


# ---- get_formulas ---------------------------------------------------------- #

def test_get_formulas(ops):
    out = ops.get_formulas("sample.xlsx", "G2:G3")
    assert out["formulas"] == [
        {"cell": "G2", "formula": "=D2*E2", "cached_value": None},
        {"cell": "G3", "formula": "=D3*E3", "cached_value": None},
    ]
    assert ops.get_formulas("sample.xlsx")["formula_count"] == 8


def test_get_formulas_rejects_csv(ops):
    with pytest.raises(ExcelOpsError, match="xlsx"):
        ops.get_formulas("sales.csv")


# ---- search ---------------------------------------------------------------- #

def test_search_text_across_sheets(ops):
    out = ops.search("sample.xlsx", "widget")
    assert [(m["sheet"], m["cell"]) for m in out["matches"]] == [
        ("Sales", "C2"), ("Sales", "C5"), ("Inventory", "B2"),
    ]


def test_search_numeric_exact_and_regex(ops):
    assert [m["cell"] for m in ops.search("sample.xlsx", "10", sheet="Sales", mode="exact")["matches"]] == [
        "D2", "E3", "E7",
    ]
    out = ops.search("sample.xlsx", r"^[WG]-\d$", mode="regex")
    assert [m["value"] for m in out["matches"]] == ["W-1", "G-2"]


def test_search_limit(ops):
    out = ops.search("sample.xlsx", "e", limit=2)
    assert out["match_count"] == 2 and out["truncated"] is True


# ---- apply_filter ---------------------------------------------------------- #

def _reps(result):
    return [r["Rep"] for r in result["rows"]]


def test_filter_all_conditions(ops):
    out = ops.apply_filter("sample.xlsx", [
        Condition("Region", "eq", "west"),
        Condition("Units", "gt", 8),
    ])
    assert _reps(out) == ["Alice"] and out["match_count"] == 1


def test_filter_any_and_operators(ops):
    out = ops.apply_filter("sample.xlsx", [
        Condition("Region", "in", ["North", "South"]),
        Condition("Product", "startswith", "gad"),
    ], match="any")
    assert _reps(out) == ["Bob", "Dave", "Erin"]
    assert _reps(ops.apply_filter("sample.xlsx", [Condition("Units", "is_null")])) == ["Erin", None]
    assert _reps(ops.apply_filter("sample.xlsx", [Condition("Units", "between", [4, 7])])) == ["Bob", "Carol"]


def test_filter_dates(ops):
    out = ops.apply_filter("sample.xlsx", [Condition("Date", "gte", "2026-02-14")], columns=["Rep", "Date"])
    assert out["rows"] == [
        {"_row": 5, "Rep": "Dave", "Date": "2026-02-14"},
        {"_row": 6, "Rep": "Alice", "Date": "2026-03-02"},
        {"_row": 7, "Rep": "Erin", "Date": "2026-03-20"},
    ]


def test_filter_validation(ops):
    with pytest.raises(ExcelOpsError, match="between"):
        ops.apply_filter("sample.xlsx", [Condition("Units", "between", 5)])
    with pytest.raises(ExcelOpsError, match="Available columns"):
        ops.apply_filter("sample.xlsx", [Condition("Nope", "eq", 1)])


# ---- security -------------------------------------------------------------- #

def test_rejects_paths_outside_sandbox(ops, workdir):
    with pytest.raises(ExcelOpsError, match="Access denied"):
        ops.read_spreadsheet("../outside.xlsx")
    with pytest.raises(ExcelOpsError, match="Access denied"):
        ops.read_spreadsheet(str(workdir.parent / "outside.xlsx"))


def test_rejects_unsupported_extension(ops, workdir):
    (workdir / "notes.txt").write_text("hi")
    with pytest.raises(ExcelOpsError, match="Unsupported file type"):
        ops.read_spreadsheet("notes.txt")


def test_rejects_missing_file(ops):
    with pytest.raises(ExcelOpsError, match="File not found"):
        ops.read_spreadsheet("missing.xlsx")


def test_rejects_oversized_file(workdir):
    tiny = ExcelOps(Config(allowed_dirs=[workdir], max_file_mb=0.001))
    with pytest.raises(ExcelOpsError, match="over the"):
        tiny.read_spreadsheet("sample.xlsx")
