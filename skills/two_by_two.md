# Pattern: Two By Two

## When to use
Use for prioritisation and risk mapping where each candidate has two scored
dimensions: impact versus effort, severity versus probability, or benefit
versus implementation complexity. It is a decision aid, not a precise plot.

## Function signature
```python
from patterns import two_by_two_slide

two_by_two_slide(
    title: str,
    x_label: str,
    y_label: str,
    x_low: str,
    x_high: str,
    y_low: str,
    y_high: str,
    items: list[dict],
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import two_by_two_slide

b = two_by_two_slide(
    title="Validation Priorities: Effort vs Risk Reduction",
    x_label="Implementation effort", y_label="Risk reduction",
    x_low="Low effort", x_high="High effort",
    y_low="Low impact", y_high="High impact",
    items=[
        {"name": "Drift alerts", "x": 0.22, "y": 0.78, "descriptor": "Quick win"},
        {"name": "DCC rebuild", "x": 0.81, "y": 0.86, "descriptor": "Strategic"},
    ],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Validation Priorities")
deck_path = project.output_path("validation_priorities.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Convert each item into x/y scores between 0 and 1 using one consistent judgement scale.
2. Keep item names short; use a descriptor only when it clarifies the decision.
3. Plot no more than roughly eight items on a single slide.
4. Choose axis direction deliberately so the preferred quadrant is immediately obvious.
5. Describe the quadrant implication in accompanying narrative when choices are contested.
