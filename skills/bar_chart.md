# Skill: Bar Chart Slide

## Pattern name
`bar_chart_slide`

## When to use
- Comparing a metric **across categories** (products, cohorts, periods, regions)
- Showing period-over-period change where each period is discrete
- Visualising ranked output — which category performs best/worst

**Do not use** for time series with many data points (>12 periods) — use `line_chart_slide` instead.

## Content rules
| Field | Guidance |
|---|---|
| `title` | Chart heading — include the metric name and time period |
| `categories` | List of category labels (≤12 items) |
| `series` | List of data series dicts |
| `footnote` | Optional source or caveat |

Each series dict:
| Key | Required | Guidance |
|---|---|---|
| `name` | Yes | Series label shown in the legend |
| `values` | Yes | List of floats, one per category (must match category count) |

## Layout
- Full-width native PowerPoint chart (not an image)
- Single series: no legend; multiple series: legend shown at bottom
- Gridlines on value axis only; no gridlines on category axis

## Distillation rules
- Extract the numbers directly from tables in the source — do not estimate
- Round to 2 decimal places maximum for readability
- If the source has >12 categories, aggregate minor ones into "Other"
- Call `b.ai_label()` if any value was inferred or interpolated

## Example
```python
from patterns import bar_chart_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Factor Attribution")
b = bar_chart_slide(
    title="Factor Return Attribution — Q2 2025",
    categories=["Equity Beta", "Duration", "Credit", "FX Carry", "Residual"],
    series=[
        {"name": "Return (bps)", "values": [32, -8, 15, 6, -4]},
    ],
    footnote="Source: MSCI Barra. Returns gross of fees.",
)
deck_path = project.output_path("factor_attribution.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```

## Multi-series example
```python
b = bar_chart_slide(
    title="PMSE by Model Version — Holdout Set",
    categories=["Short (<1Y)", "Medium (1–5Y)", "Long (>5Y)", "All"],
    series=[
        {"name": "Model v2", "values": [0.048, 0.061, 0.072, 0.060]},
        {"name": "Model v3", "values": [0.031, 0.049, 0.058, 0.046]},
    ],
    footnote="Lower PMSE is better. Holdout period: Jan–Mar 2025.",
)
```
