# Skill: KPI Dashboard Slide

## Pattern name
`kpi_dashboard_slide`

## When to use
- 3–6 headline metrics need to appear on one slide
- Audience wants a scorecard snapshot before detail slides
- Metrics have a direction signal (up/down vs prior period)

## Content rules
| Field | Guidance |
|---|---|
| `title` | Dashboard heading (≤8 words) |
| `kpis` | List of up to 6 KPI dicts |

Each KPI dict:
| Key | Required | Guidance |
|---|---|---|
| `name` | Yes | Short metric label (≤4 words) |
| `value` | Yes | The formatted number/text to show large (e.g. `"$2.3B"`, `"94%"`) |
| `delta` | No | Change vs prior period (e.g. `"+12bps"`, `"-3pp"`) |
| `trend` | No | `"up"`, `"down"`, or `"flat"` — drives the arrow colour |

## Layout
- Up to 3 cards per row; 2 rows for 4–6 KPIs
- Cards are centred on the canvas; equal width and height (symmetry enforced by the pattern)
- Trend up → green arrow; down → red; flat → grey

## Distillation rules
- Pull only headline figures, not every metric in the source
- Prefer values the audience will act on (e.g. P&L, AUM, Sharpe, coverage ratio)
- delta should be the most recent change, not a cumulative change

## Example
```python
from patterns import kpi_dashboard_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Q2 Dashboard")
b = kpi_dashboard_slide(
    title="Q2 2025 — Portfolio at a Glance",
    kpis=[
        {"name": "AUM",          "value": "$4.7B",   "delta": "+$0.3B",  "trend": "up"},
        {"name": "Sharpe Ratio", "value": "1.12",    "delta": "+0.08",   "trend": "up"},
        {"name": "Vol (ann.)",   "value": "8.4%",    "delta": "+1.2pp",  "trend": "down"},
        {"name": "Max Drawdown", "value": "-4.1%",   "delta": "-0.8pp",  "trend": "up"},
        {"name": "IR",           "value": "0.71",    "delta": "—",       "trend": "flat"},
        {"name": "Coverage",     "value": "98.3%",   "delta": "+0.5pp",  "trend": "up"},
    ],
)
deck_path = project.output_path("q2_dashboard.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```
