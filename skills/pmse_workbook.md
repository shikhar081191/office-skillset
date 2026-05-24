# Pattern: PMSE Comparison Workbook

## When to use
Use for an editable Excel comparison of candidate models across asset cohorts,
with consistent number formats and a summary chart.

## Example call
```python
from xlsx_patterns import pmse_comparison_workbook

b = pmse_comparison_workbook(
    title="Treasury Spread Model PMSE",
    cohorts=["Govt Bonds", "Treasury Bills", "OTR Bills"],
    models=["Current", "Model A", "Model B"],
    values=[[0.033, 0.011, 0.004], [0.602, 0.224, 0.141], [0.400, 0.182, 0.116]],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("PMSE Comparison")
book_path = project.output_path("pmse_comparison.xlsx")
b.save(book_path, final=True, report_path=project.qa_path(book_path.name))
```
