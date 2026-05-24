# Skill: Recommendation Slide

## Pattern name
`recommendation_slide`

## When to use
- The output of the analysis is a clear **go / no-go / choose-option decision**
- The deck owner wants the conclusion visible before the supporting detail
- The audience needs to leave with one unambiguous action

## Content rules
| Field | Guidance |
|---|---|
| `title` | Section or topic heading (≤8 words) |
| `recommendation` | One crisp sentence — the decision itself (≤20 words). Must start with a verb: *Approve*, *Reject*, *Proceed with*, *Select*. |
| `rationale` | 3–5 bullets, each ≤12 words. Evidence, not opinion. |
| `caveats` | Optional. 1–3 conditions or risks that qualify the recommendation. |

## Distillation rules
- The `recommendation` string is the headline — it goes large on the dark left panel. Never paste a paragraph here.
- Rationale bullets should each stand alone; the reader should understand each without the others.
- If the source document has a *Conclusion* or *Recommendation* section, extract the first sentence as `recommendation` and subsequent supporting points as `rationale`.

## Example
```python
from patterns import recommendation_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Credit Model Review")
b = recommendation_slide(
    title="Credit Model v3 — Decision",
    recommendation="Approve Credit Model v3 for production deployment in Q3.",
    rationale=[
        "PMSE improved 18% vs baseline on holdout set",
        "Stress-test performance within regulatory tolerance",
        "Independent model validation sign-off received 14 May",
        "Parallel run confirms no material change to capital requirements",
    ],
    caveats=[
        "Monitor vintage drift monthly for the first two quarters",
        "Recalibration required if macro scenario shifts by >2σ",
    ],
)
deck_path = project.output_path("credit_model_decision.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```
