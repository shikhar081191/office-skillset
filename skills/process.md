# Pattern: Process

## When to use
Use when you need to show a linear sequence of steps — a workflow, lifecycle,
methodology, or procedure. Best for 3–5 steps. Ideal for: trade lifecycles,
onboarding flows, approval processes, model development pipelines.

## Function signature
```python
from patterns import process_slide

process_slide(
    title: str,               # slide title
    steps: list[dict],        # 3–5 step dicts (see below)
    highlight: int = None,    # 0-based index of the step to call out in accent colour
    footnote: str = None,     # optional small italic note below the strip
    builder=None,             # pass existing PptxBuilder to append; else creates new
    palette=None,             # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

Each step dict:

| Key         | Type | Rules                                        |
|-------------|------|----------------------------------------------|
| name        | str  | Step name — 2–4 words, title case            |
| description | str  | One sentence describing what happens, ≤ 90 chars |

## Example call
```python
from patterns import process_slide

b = process_slide(
    title="Model Development Lifecycle",
    steps=[
        {"name": "Data Sourcing",   "description": "Identify and licence raw data feeds from approved vendors"},
        {"name": "Feature Eng.",    "description": "Transform and normalise inputs; remove look-ahead bias"},
        {"name": "Model Training",  "description": "Train on 10yr in-sample period with walk-forward CV"},
        {"name": "Backtesting",     "description": "Out-of-sample test on 3yr hold-out; track key metrics"},
        {"name": "Production",      "description": "Deploy via model registry with A/B shadow run"},
    ],
    highlight=2,   # highlight "Model Training"
    footnote="* All backtests use point-in-time data to prevent look-ahead contamination.",
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Model Lifecycle")
deck_path = project.output_path("model_lifecycle.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. **Step count**: 3 is clean; 4–5 is the sweet spot; never exceed 5 (boxes become too narrow).
2. **Step names**: short noun phrases — "Risk Review", "Execution", "Settlement". Not sentences.
3. **Descriptions**: one active sentence per step. Start with a verb ("Validate", "Route", "Match").
4. **Highlight**: use to call out the step that is the current focus, the problem area, or
   the step you want the audience to remember. Omit if all steps are equal weight.
5. **Footnote**: use only for qualifications or scope notes that apply to the whole process
   (e.g. "* Applies to vanilla instruments only"). One sentence maximum.
6. **Sequence**: steps must read left-to-right in natural flow. If the user gives a circular
   or branching process, pick the dominant linear path and note exceptions in the footnote.
