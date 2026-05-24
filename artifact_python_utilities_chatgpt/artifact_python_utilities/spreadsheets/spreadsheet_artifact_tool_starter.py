"""Minimal artifact_tool spreadsheet starter utilities.

These helpers are intended as copy/paste scaffolding for creating or editing .xlsx
files with artifact_tool. They assume artifact_tool is installed in the runtime.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable, Sequence

from artifact_tool import Blob, SpreadsheetFile, Workbook


def create_workbook_with_sheet(sheet_name: str = "Sheet1"):
    """Create a new workbook with one worksheet."""
    wb = Workbook.create()
    sheet = wb.worksheets.add(sheet_name)
    return wb, sheet


def import_workbook(path: str | Path):
    """Load an existing .xlsx workbook from disk."""
    return SpreadsheetFile.import_xlsx(Blob.load(str(path)))


def export_workbook(wb, path: str | Path) -> None:
    """Save a workbook to .xlsx."""
    SpreadsheetFile.export_xlsx(wb).save(str(path))


def write_table(sheet, top_left_range: str, rows: Sequence[Sequence[Any]]) -> None:
    """Write a 2D row matrix into a sheet starting at a range such as 'A1'."""
    sheet.get_range(top_left_range).write(rows)


def style_basic_header(sheet, header_range: str, fill: str = "#1F4E79") -> None:
    """Apply a simple professional header style to a range like 'A1:H1'."""
    rng = sheet.get_range(header_range)
    rng.format = {
        "fill": fill,
        "font": {"bold": True, "color": "#FFFFFF"},
        "horizontal_alignment": "center",
        "vertical_alignment": "center",
        "row_height": 22,
        "wrap_text": True,
    }


def add_dropdown(sheet, target_range: str, values: Iterable[str]) -> None:
    """Add list data validation to a range."""
    sheet.get_range(target_range).data_validation = {
        "rule": {"type": "list", "values": list(values)}
    }


def scan_formula_errors(wb, max_results: int = 300):
    """Return an inspect result for common Excel formula errors."""
    return wb.inspect(
        {
            "kind": "match",
            "search_term": r"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
            "options": {"use_regex": True, "max_results": max_results},
            "summary": "formula error scan",
        }
    )


def render_range(wb, sheet_name: str, cell_range: str = "A1:H20", scale: int = 2):
    """Render a sheet range for visual verification."""
    return wb.render({"sheet_name": sheet_name, "range": cell_range, "scale": scale})


if __name__ == "__main__":
    wb, sheet = create_workbook_with_sheet("Dashboard")
    rows = [
        ["Metric", "Value", "Status"],
        ["Open tasks", 12, "In progress"],
        ["Blocked tasks", 2, "Watch"],
    ]
    write_table(sheet, "A1", rows)
    style_basic_header(sheet, "A1:C1")
    add_dropdown(sheet, "C2:C50", ["Not started", "In progress", "Watch", "Done"])
    sheet.get_range("A1:C50").format.autofit_columns()
    export_workbook(wb, "/mnt/data/example_output.xlsx")
