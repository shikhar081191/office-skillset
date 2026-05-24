# Pattern: Waterfall

## When to use
Use to show which positive and negative effects compose a net result, such as
factor contributions to PMSE improvement or P&L attribution. It works best
when the audience needs to see what offsets the gain and where the total lands.

## Function signature
```python
from patterns import waterfall_slide

waterfall_slide(
    title: str,
    items: list[dict],
    total_label: str = "Total",
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import waterfall_slide

b = waterfall_slide(
    title="Contribution to PMSE Improvement vs Baseline",
    items=[
        {"name": "Rates factors", "value": 12.5},
        {"name": "Credit spreads", "value": 8.2},
        {"name": "Volatility regime", "value": 5.1},
        {"name": "FX carry noise", "value": -2.8},
    ],
    total_label="Net improvement (bps)",
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("PMSE Contribution")
deck_path = project.output_path("pmse_contribution.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Express every contribution in the same unit and keep the sign meaningful.
2. Limit to roughly 4-8 drivers; group immaterial residuals into an "Other" row.
3. Sort by narrative order or materiality, not alphabetically.
4. Make the total label explicit about its metric and unit.
5. Use negative items only for offsets, costs or degradation.
