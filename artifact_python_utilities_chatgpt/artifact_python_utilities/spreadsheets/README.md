# Spreadsheet Python utilities

The spreadsheet workflow primarily uses the installed `artifact_tool` Python API rather than a large local script folder.

Included here:

- `artifact_tool_quick_start.md` — compact API reference for workbook creation, import/export, formulas, formatting, validation, charts, rendering, and inspection.
- `spreadsheet_artifact_tool_starter.py` — small helper scaffold with common create/import/export/write/style/validation/verify functions.

Minimal pattern:

```python
from artifact_tool import Workbook, SpreadsheetFile

wb = Workbook.create()
sheet = wb.worksheets.add("Dashboard")
sheet.get_range("A1:B2").values = [["Metric", "Value"], ["Open", 12]]
SpreadsheetFile.export_xlsx(wb).save("/mnt/data/output.xlsx")
```
