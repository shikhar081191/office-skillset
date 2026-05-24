# Pattern: Chart and Context

## When to use
Use when you have a chart or visualisation and want to pair it with a structured
narrative. The chart fills the left 60% of the slide; the right 40% provides
headline, supporting bullets, and a bold "so what" takeaway. Ideal for: market
commentary, model output review, performance attribution, research findings.

## Function signature
```python
from patterns import chart_context_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Credit Spreads")

chart_context_slide(
    title: str,               # small slide title top-left
    chart_path: str | None,   # file path to .png/.jpg image; None shows a placeholder
    headline: str,            # 10–15 word bold summary of what the chart shows
    bullets: list[str],       # 2–3 supporting observations (one sentence each)
    so_what: str,             # single action/implication line in accent colour
    builder=None,             # pass existing PptxBuilder to append; else creates new
    palette=None,             # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

## Example call
```python
from patterns import chart_context_slide

b = chart_context_slide(
    title="Credit Spread Dynamics — Q2 2025",
    chart_path=project.asset_path("spread_chart.png"),   # or None for placeholder
    headline="IG spreads have tightened 18bps while HY has widened",
    bullets=[
        "IG OAS moved from 112bps to 94bps — driven by tech sector compression",
        "HY OAS widened 35bps on energy and consumer discretionary stress",
        "BB/BBB crossover ratio now at 0.72, lowest since Jan 2023",
    ],
    so_what="Rotate 10% from HY into IG: risk/reward has shifted materially",
)
deck_path = project.output_path("credit_spreads.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Generating the chart image
Save your chart as a .png before calling this function:
```python
import matplotlib.pyplot as plt
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Credit Spreads")
fig, ax = plt.subplots(figsize=(8, 5))
ax.plot(dates, values)
fig.savefig(project.asset_path("spread_chart.png"), dpi=150, bbox_inches="tight")
plt.close()
```

## Content-shaping rules
1. **Headline**: one declarative sentence stating the chart's main message.
   "What does the chart show?" — not "Chart of spreads over time".
   Good: "IG spreads have tightened 18bps while HY has widened"
2. **Bullets**: 2–3 is the limit. Each bullet expands on one aspect of the headline.
   Start with the observation, end with the interpretation or magnitude.
3. **So what**: the action or implication that follows from the chart.
   One sentence. Should feel actionable: "Reduce", "Increase", "Monitor", "Alert".
4. **chart_path=None**: pass None (or omit) to show a grey placeholder during drafting.
   Replace with the actual path once the chart is generated.
5. **Image dimensions**: portrait charts (taller than wide) may not fill the panel well.
   For best results, generate charts at approximately 8:5 (width:height) aspect ratio.
