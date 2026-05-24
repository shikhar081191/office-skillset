# Pattern: Diverging Bar

## When to use
Use to compare two positive measurements across the same categories when the
direction and size of the difference matter. Typical applications include
current versus candidate model PMSE, volatility by sub-period or exposure by scenario.

## Function signature
```python
from patterns import diverging_bar_slide

diverging_bar_slide(
    title: str,
    row_labels: list[str],
    left_label: str,
    right_label: str,
    left_values: list[float],
    right_values: list[float],
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import diverging_bar_slide

b = diverging_bar_slide(
    title="PMSE Comparison by Asset Cohort",
    row_labels=["Bills", "UST 2Y", "UST 10Y", "IG Credit", "Equity Beta"],
    left_label="Current Model", right_label="Model B",
    left_values=[0.74, 0.65, 0.59, 0.68, 0.77],
    right_values=[0.36, 0.44, 0.48, 0.52, 0.61],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("PMSE Comparison")
deck_path = project.output_path("pmse_comparison.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Use paired positive values in the same unit for every row.
2. Keep category ordering stable or sort by absolute gap to support the message.
3. Limit to approximately 5-9 rows so labels and bar ends remain readable.
4. Use labels that clearly identify the baseline and comparison case.
5. Select this pattern when paired comparison is the story; use waterfall for contribution.
