# Skill: Line Chart Slide

## Pattern name
`line_chart_slide`

## When to use
- Showing how a metric **changes over time** (time series)
- Comparing trend lines across multiple scenarios or models
- Visualising cumulative performance, rolling statistics, or index levels

**Do not use** for categorical comparisons — use `bar_chart_slide` instead.

## Content rules
| Field | Guidance |
|---|---|
| `title` | Chart heading — include the metric name and date range |
| `categories` | Ordered list of time-period labels (dates, quarters, months) |
| `series` | List of data series dicts |
| `footnote` | Optional source or disclaimer |

Each series dict:
| Key | Required | Guidance |
|---|---|---|
| `name` | Yes | Series label shown in the legend |
| `values` | Yes | List of floats, one per time period (must match category count) |

## Layout
- Full-width native PowerPoint chart (not an image)
- Single series: no legend; multiple series: legend at bottom
- Gridlines on value axis only; no gridlines on category axis

## Distillation rules
- Use consistent period labels (e.g. `"Q1 24"`, `"Q2 24"`) — not mixed formats
- Truncate to 20 data points maximum for legibility on a slide
- If the source table has weekly data, aggregate to monthly or quarterly
- Call `b.ai_label()` if any value was estimated or interpolated

## Example
```python
from patterns import line_chart_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Portfolio Review")
b = line_chart_slide(
    title="Cumulative Return vs Benchmark — 2024",
    categories=["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    series=[
        {"name": "Portfolio",  "values": [0.8, 1.5, 2.1, 1.9, 3.2, 4.0, 3.8, 4.9, 5.4, 6.1, 7.2, 8.0]},
        {"name": "Benchmark",  "values": [0.5, 1.0, 1.6, 1.4, 2.5, 3.1, 2.9, 3.6, 4.0, 4.8, 5.5, 6.2]},
    ],
    footnote="Source: Bloomberg. Returns in USD, gross of fees. Benchmark: Bloomberg US Agg.",
)
deck_path = project.output_path("cumulative_return.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```
