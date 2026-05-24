# Pattern: Annotation Chart

## When to use
Use when a chart image is already available and the slide should direct attention
to events, turning points or regime changes. It is suited to rolling risk metrics,
shock histories, factor returns and backtest diagnostics.

## Function signature
```python
from patterns import annotation_chart_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Correlation Review")

annotation_chart_slide(
    title: str,
    chart_path: str,
    annotations: list[dict],
    headline: str | None = None,
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import annotation_chart_slide

b = annotation_chart_slide(
    title="Rolling Equity-Credit Correlation: Regime Signals",
    chart_path=project.asset_path("correlation_chart.png"),
    annotations=[
        {"label": "Liquidity shock", "x": 0.42, "y": 0.36},
        {"label": "Policy pivot", "x": 0.78, "y": 0.20},
    ],
    headline="Correlation risk re-entered the stress band",
)
deck_path = project.output_path("annotated_correlation.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Use a resolved PNG or JPG chart file before constructing the slide.
2. Coordinates are relative to the chart image bounds and must be between 0 and 1.
3. Limit annotations to the 2-4 events needed for the conclusion.
4. Labels should name the event or insight in a few words, not restate the headline.
5. The optional headline should state the finding, not describe the chart type.
