# Skill: Assumption Table Slide

## Pattern name
`assumption_table_slide`

## When to use
- Analysis rests on explicit model or economic assumptions
- Audience (e.g. risk committee, board) needs to validate or challenge inputs
- Documenting stress-test parameters, macro scenarios, or pricing assumptions

## Content rules
| Field | Guidance |
|---|---|
| `title` | `"Key Assumptions"` or `"Model Parameters"` |
| `assumptions` | List of up to 12 assumption rows |
| `footnote` | Optional source or disclaimer line |

Each assumption dict:
| Key | Required | Guidance |
|---|---|---|
| `assumption` | Yes | What is being assumed (≤10 words) |
| `value` | Yes | The specific value, range, or qualitative level |
| `source` | No | Where the assumption comes from (e.g. `"Bloomberg", "MSCI", "Internal"`) |
| `sensitivity` | No | `"high"`, `"medium"`, or `"low"` — colour-coded in the table |

## Layout
- Assumption column is widest; colour-coded sensitivity cell (red/amber/green)
- Alternating row shading for readability
- Up to 12 rows — if more, split into two slides by theme

## Distillation rules
- Pull only the assumptions that materially affect the conclusion
- `value` should be a number or short phrase, not a sentence
- Sensitivity = how much the conclusion changes if this assumption is wrong

## Example
```python
from patterns import assumption_table_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Scenario Analysis")
b = assumption_table_slide(
    title="Macro Assumptions — Stress Scenario",
    assumptions=[
        {"assumption": "GDP growth (2025)",  "value": "-1.5%",      "source": "IMF WEO",   "sensitivity": "high"},
        {"assumption": "Rate path (10Y UST)", "value": "+75bps",     "source": "Fed dots",  "sensitivity": "high"},
        {"assumption": "IG spread widening",  "value": "+120bps",    "source": "Internal",  "sensitivity": "medium"},
        {"assumption": "Equity drawdown",     "value": "-20%",       "source": "Internal",  "sensitivity": "high"},
        {"assumption": "FX (EURUSD)",         "value": "0.95",       "source": "Bloomberg", "sensitivity": "low"},
        {"assumption": "Recovery rate",       "value": "40%",        "source": "S&P LossStats", "sensitivity": "medium"},
    ],
    footnote="* Assumptions represent a 1-in-20 adverse scenario. Not a forecast.",
)
deck_path = project.output_path("assumptions.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```
