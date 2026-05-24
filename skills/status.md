# Pattern: Status

## When to use
Use when you have 4–8 workstreams, projects, or initiatives and need to show their
health at a glance. Classic use: programme update for a steering committee or weekly
status deck. The 2-column grid scales to 8 cards before it becomes cluttered.

## Function signature
```python
from patterns import status_slide

status_slide(
    title: str,                # slide title, top-left, small — include time period
    workstreams: list[dict],   # 4–8 workstream dicts (see below)
    builder=None,              # pass existing PptxBuilder to append; else creates new
    palette=None,              # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

Each workstream dict:

| Key    | Type | Rules                                              |
|--------|------|----------------------------------------------------|
| name   | str  | Workstream name — title case, ≤ 30 chars           |
| status | str  | One-line status, < 80 chars (see shaping rules)   |
| rag    | str  | "green", "amber", or "red"                         |
| owner  | str  | 1–2 character initial (e.g. "AK", "S")             |

## Example call
```python
from patterns import status_slide

b = status_slide(
    title="Q2 2025 Programme Status",
    workstreams=[
        {"name": "Data Pipeline",    "status": "On track; UAT starts 3 Jun",       "rag": "green", "owner": "AK"},
        {"name": "Risk Models",      "status": "2-week delay — vendor data late",   "rag": "amber", "owner": "SP"},
        {"name": "Compliance",       "status": "Legal sign-off still outstanding",  "rag": "red",   "owner": "TJ"},
        {"name": "Client Reporting", "status": "MVP live; feedback round complete", "rag": "green", "owner": "LM"},
    ]
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Q2 Status")
deck_path = project.output_path("q2_status.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. **Extract workstreams**: identify each distinct workstream/project/squad from the user's text.
2. **One-line status**: distil to key fact + implication. Avoid "we are" or "the team is".
   - Bad: "The team is currently working on the data pipeline migration and it is going well"
   - Good: "Migration 70% complete; UAT scheduled for 3 Jun"
3. **RAG assignment**:
   - green = on track, no blockers
   - amber = at risk, delay possible, or waiting on external dependency
   - red = blocked, off-track, or escalation required
4. **Owner**: first letter of last name, or first+last initials. One or two characters only.
5. **Volume**: keep to 6 cards maximum. 8 is the absolute ceiling. If there are more
   workstreams, group smaller ones under a theme card.
6. **Title**: always include the time period (e.g. "Q2 2025 Status", "Week 22 Update").
