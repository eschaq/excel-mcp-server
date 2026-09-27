"""MCP tool definitions. Each tool is a thin wrapper over ``ExcelOps``."""

from __future__ import annotations

import functools
import logging
from typing import Annotated, Any, Callable, Literal

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import BaseModel, Field

from excel_mcp import __version__
from excel_mcp.excel_ops import Condition, Config, ExcelOps, ExcelOpsError, FilterOp

log = logging.getLogger("excel_mcp")

INSTRUCTIONS = """\
Read, analyze and edit local spreadsheet files (.xlsx, .xlsm, .xls, .csv).
Start with read_spreadsheet to see sheets, sizes and a preview, then use get_sheet_data,
apply_filter, search or get_summary_stats. Data rows come back as records with a "_row" key
holding the Excel row number, so you can write back to the right cell. Writes (write_cell,
write_range, create_sheet, create_workbook) only work on .xlsx/.xlsm files; a string value
starting with "=" is stored as a formula. Formula results are not recalculated until the file
is opened in Excel, so newly written formulas read back as empty values: use get_formulas to see them.
"""

Path_ = Annotated[str, Field(description="Path to the spreadsheet file (absolute, or relative to the first allowed directory)")]
Sheet = Annotated[str | None, Field(description="Sheet name. Defaults to the first sheet.")]
HeaderRow = Annotated[int, Field(description="1-based row holding column names; 0 means no header (columns named A, B, C...)")]
Columns = Annotated[list[str] | None, Field(description="Only return these columns. Defaults to all.")]
CellValue = str | int | float | bool | None

READ = ToolAnnotations(readOnlyHint=True, destructiveHint=False, idempotentHint=True, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=True, openWorldHint=False)


class FilterCondition(BaseModel):
    column: str = Field(description="Column name (case-insensitive)")
    op: FilterOp = Field(description="Comparison operator")
    value: Any = Field(
        default=None,
        description="Value to compare against. A list for 'in'/'not_in', [low, high] for 'between', "
        "omitted for 'is_null'/'not_null'. Dates as ISO strings, e.g. '2026-01-31'.",
    )
    case_sensitive: bool = Field(default=False, description="Case-sensitive text comparison")


def _tool_errors(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Surface anticipated errors to the model as readable tool errors."""

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except ExcelOpsError as e:
            log.info("%s rejected: %s", fn.__name__, e)
            raise ToolError(str(e)) from e

    return wrapper


def build_server(config: Config) -> MCPServer:
    ops = ExcelOps(config)
    server = MCPServer(name="excel", title="Excel", version=__version__, instructions=INSTRUCTIONS)

    def tool(annotations: ToolAnnotations):
        return lambda fn: server.tool(annotations=annotations)(_tool_errors(fn))

    # ---- read tools -------------------------------------------------------- #

    @tool(READ)
    def read_spreadsheet(
        path: Path_,
        preview_rows: Annotated[int, Field(description="Rows to preview per sheet (max 50)")] = 5,
    ) -> dict:
        """Open a spreadsheet and list its sheets with dimensions and a preview of the first rows."""
        return ops.read_spreadsheet(path, preview_rows)

    @tool(READ)
    def get_sheet_data(
        path: Path_,
        sheet: Sheet = None,
        header_row: HeaderRow = 1,
        columns: Columns = None,
        offset: Annotated[int, Field(description="Data rows to skip, for paging")] = 0,
        limit: Annotated[int, Field(description="Max rows to return (max 5000)")] = 500,
    ) -> dict:
        """Return a sheet's rows as JSON records, optionally limited to some columns and paged."""
        return ops.get_sheet_data(path, sheet, header_row, columns, offset, limit)

    @tool(READ)
    def get_formulas(
        path: Path_,
        cell_range: Annotated[str | None, Field(description="Range like 'A1:D20'. Defaults to the whole sheet.")] = None,
        sheet: Sheet = None,
    ) -> dict:
        """List the formulas (not values) in a range, with each cell's last cached result. .xlsx/.xlsm only."""
        return ops.get_formulas(path, cell_range, sheet)

    @tool(READ)
    def search(
        path: Path_,
        query: Annotated[str, Field(description="Text, number or regex to look for")],
        sheet: Annotated[str | None, Field(description="Sheet to search. Defaults to all sheets.")] = None,
        mode: Annotated[Literal["contains", "exact", "regex"], Field(description="How to match")] = "contains",
        case_sensitive: bool = False,
        limit: Annotated[int, Field(description="Max matches to return")] = 100,
    ) -> dict:
        """Find cells whose value matches a text, numeric or regex pattern. Returns cell addresses."""
        return ops.search(path, query, sheet, mode, case_sensitive, limit)

    @tool(READ)
    def get_summary_stats(
        path: Path_,
        sheet: Sheet = None,
        header_row: HeaderRow = 1,
        columns: Columns = None,
    ) -> dict:
        """Summary statistics per column: count, missing, min, max, mean, median, stdev and sum for
        numeric columns; count, unique and most common value for text columns."""
        return ops.get_summary_stats(path, sheet, header_row, columns)

    @tool(READ)
    def apply_filter(
        path: Path_,
        conditions: Annotated[list[FilterCondition], Field(description="Conditions to test each row against")],
        sheet: Sheet = None,
        header_row: HeaderRow = 1,
        match: Annotated[Literal["all", "any"], Field(description="Rows must match all conditions, or any")] = "all",
        columns: Columns = None,
        limit: Annotated[int, Field(description="Max rows to return (max 5000)")] = 500,
    ) -> dict:
        """Filter rows by column conditions (eq, ne, gt, gte, lt, lte, between, contains, not_contains,
        startswith, endswith, regex, in, not_in, is_null, not_null) and return the matching rows."""
        conds = [Condition(**c.model_dump()) for c in conditions]
        return ops.apply_filter(path, conds, sheet, header_row, match, columns, limit)

    # ---- write tools (not registered in read-only mode) --------------------- #

    if config.read_only:
        return server

    @tool(WRITE)
    def write_cell(
        path: Path_,
        cell: Annotated[str, Field(description="Cell address, e.g. 'B3'")],
        value: Annotated[CellValue, Field(description="Value to write. Strings starting with '=' become formulas; null clears the cell.")],
        sheet: Sheet = None,
    ) -> dict:
        """Write one value or formula to a cell in an .xlsx/.xlsm file. Returns the previous value."""
        return ops.write_cell(path, cell, value, sheet)

    @tool(WRITE)
    def write_range(
        path: Path_,
        start_cell: Annotated[str, Field(description="Top-left cell of the block, e.g. 'A2'")],
        values: Annotated[list[list[CellValue]], Field(description="2D array of rows to write")],
        sheet: Sheet = None,
    ) -> dict:
        """Write a 2D block of values (rows of cells) starting at a cell in an .xlsx/.xlsm file."""
        return ops.write_range(path, start_cell, values, sheet)

    @tool(WRITE)
    def create_sheet(
        path: Path_,
        name: Annotated[str, Field(description="New sheet name (max 31 chars)")],
        headers: Annotated[list[str] | None, Field(description="Optional header row for the new sheet")] = None,
        index: Annotated[int | None, Field(description="Position among sheets (0 = first). Defaults to last.")] = None,
    ) -> dict:
        """Add a new sheet to an existing .xlsx/.xlsm workbook."""
        return ops.create_sheet(path, name, headers, index)

    @tool(WRITE)
    def create_workbook(
        path: Annotated[str, Field(description="Where to create the .xlsx file")],
        sheet_name: str = "Sheet1",
        headers: Annotated[list[str] | None, Field(description="Optional header row (bold, frozen)")] = None,
        overwrite: Annotated[bool, Field(description="Replace the file if it already exists")] = False,
    ) -> dict:
        """Create a new .xlsx workbook, optionally with a header row."""
        return ops.create_workbook(path, sheet_name, headers, overwrite)

    return server
