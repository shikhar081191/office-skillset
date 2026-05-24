# Pattern: Timeline

## When to use
Use when you need to show a project roadmap, delivery schedule, or sequence of
milestones across time. Best for 4–8 milestones. Labels alternate above/below the
spine so neighbouring items don't overlap. Past milestones are grey, future ones are
in accent colour, and the current one is called out with a ring.

## Function signature
```python
from patterns import timeline_slide

timeline_slide(
    title: str,               # slide title
    milestones: list[dict],   # 4–8 milestone dicts, ordered earliest → latest
    builder=None,             # pass existing PptxBuilder to append; else creates new
    palette=None,             # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

Each milestone dict:

| Key   | Type | Rules                                             |
|-------|------|---------------------------------------------------|
| date  | str  | Short date string, e.g. "Jan 25", "Q3 2025"      |
| label | str  | Milestone name — 1–3 words, title case            |
| state | str  | "past", "current", or "future"                    |

## Example call
```python
from patterns import timeline_slide

b = timeline_slide(
    title="Programme Delivery Roadmap — 2025",
    milestones=[
        {"date": "Feb 25", "label": "Kick-off",           "state": "past"},
        {"date": "Apr 25", "label": "Design Freeze",      "state": "past"},
        {"date": "Jun 25", "label": "Build Complete",     "state": "current"},
        {"date": "Aug 25", "label": "UAT",                "state": "future"},
        {"date": "Oct 25", "label": "Go-Live",            "state": "future"},
        {"date": "Dec 25", "label": "Lessons Learned",    "state": "future"},
    ]
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Roadmap")
deck_path = project.output_path("roadmap.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. **Milestone count**: 5–6 is ideal for readability. Minimum 3. Maximum 8 before
   labels start to overlap.
2. **Labels**: noun phrases only — not sentences. "Design Freeze" not "Design phase is frozen".
3. **Dates**: short and consistent. Pick one format and use it throughout
   ("Jan 25" or "Q1 25" or "Jan 2025" — never mix formats).
4. **State assignment**: exactly one milestone should be "current". It marks where the
   project is today. If the user does not specify, infer from context.
5. **Ordering**: always chronological left to right. Never reverse or skip around.
6. **Spacing**: milestones are evenly spaced visually regardless of actual calendar
   distance. If exact spacing matters, note it in the slide title or a footnote added
   via a follow-up `content_slide`.
