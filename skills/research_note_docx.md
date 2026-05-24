# Pattern: Research Note Document

## When to use
Use for a polished Word note containing an executive summary, findings,
supporting tables, methodology and next steps. It is also suitable as a source
document that may later be converted into a slide deck.

## Example call
```python
from docx_patterns import research_note_document

b = research_note_document(
    title="Treasury Spread Model Review",
    subtitle="Portfolio Risk Modelling | May 2026",
    executive_summary="Model B materially improves bill PMSE without sacrificing stability.",
    findings=["Bills PMSE improves by 46 percentage points.", "OOS checks remain stable."],
    data_tables=[{
        "title": "PMSE Comparison",
        "data": [["Cohort", "Current", "Model B"], ["Bills", "60.2%", "14.1%"]],
        "col_widths": [2.5, 1.5, 1.5],
    }],
    next_steps=["Complete sub-period validation.", "Seek RQA sign-off."],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Treasury Note")
note_path = project.output_path("treasury_note.docx")
b.save(note_path, final=True, report_path=project.qa_path(note_path.name))
```
