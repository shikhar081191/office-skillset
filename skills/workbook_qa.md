# Workflow Skill: Workbook Quality Gate

## When to use
Use after creating an Excel workbook for analysis, model comparison, scenario
inputs, validation tracking, or reporting.

## Required workflow
Final workbook scripts must save with the built-in structural audit:

```python
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("PMSE Dashboard")
book_path = project.output_path("pmse_dashboard.xlsx")
b.save(book_path, final=True, report_path=project.qa_path(book_path.name))
```

The gate blocks unfinished placeholder content and invalid formula references,
and flags empty sheets or unusually wide columns. Review calculation logic,
units, number formats, and input/formula colour conventions before delivery.
