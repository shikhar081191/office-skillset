# Pattern: Heat Map

## When to use
Use for a two-dimensional numeric comparison where intensity communicates the
story quickly: PMSE by model and cohort, factor exposure by portfolio, or
scenario sensitivity by horizon. It is best when rows and columns are both
meaningful categories and viewers should find concentrations or outliers.

## Function signature
```python
from patterns import heat_map_slide

heat_map_slide(
    title: str,
    row_labels: list[str],
    col_labels: list[str],
    values: list[list[float]],
    show_values: bool = True,
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import heat_map_slide

b = heat_map_slide(
    title="PMSE by Asset Cohort and Candidate Model",
    row_labels=["Rates", "Credit", "Equities", "FX"],
    col_labels=["Baseline", "Model A", "Model B", "Ensemble"],
    values=[[0.62, 0.55, 0.48, 0.44],
            [0.71, 0.61, 0.52, 0.46],
            [0.81, 0.69, 0.58, 0.50],
            [0.49, 0.43, 0.39, 0.35]],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("PMSE Heat Map")
deck_path = project.output_path("pmse_heat_map.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Use 3-8 categories on each axis; large matrices become unreadable on a slide.
2. Ensure every row has exactly one value per column and all values are numeric.
3. Use consistent units; do not mix PMSE, volatility and returns in one grid.
4. Keep values visible when exact comparison matters; hide them only for a pure pattern read.
5. Name the title around the interpretation, metric and comparison population.
