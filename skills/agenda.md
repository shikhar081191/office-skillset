# Skill: Agenda Slide

## Pattern name
`agenda_slide`

## When to use
- Opening a multi-section deck where the audience needs orientation
- Returning to a running agenda between major sections (mark the current item as `active`)
- The deck has 3 or more distinct sections

## Content rules
| Field | Guidance |
|---|---|
| `title` | `"Agenda"` or `"Today's Discussion"` |
| `items` | List of up to 8 agenda items |

Each item dict:
| Key | Required | Guidance |
|---|---|---|
| `label` | Yes | Section name (≤6 words) |
| `description` | No | One-line elaboration shown smaller beneath the label (≤10 words) |
| `active` | No | Set `True` for the current section when reusing the slide between sections |

## Layout
- Numbered badges on the left; label and optional description on the right
- Active item: highlighted row, filled badge, bold label
- Inactive items: muted colour — they exist but don't compete with the active section

## Distillation rules
- Agenda items should match section divider titles exactly
- Descriptions are optional — omit if the label is self-explanatory
- Maximum 8 items; if the source has more, combine minor sub-sections

## Example
```python
from patterns import agenda_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Review")
b = agenda_slide(
    title="Agenda",
    items=[
        {"label": "Executive Summary",    "description": "Key findings and recommendation"},
        {"label": "Model Performance",    "description": "PMSE, Sharpe, and stress tests", "active": True},
        {"label": "Sensitivity Analysis", "description": "Factor and macro scenario analysis"},
        {"label": "Implementation Plan",  "description": "Timeline and resource requirements"},
        {"label": "Appendix"},
    ],
)
deck_path = project.output_path("agenda.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```
