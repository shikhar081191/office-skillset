# Pattern: Scorecard

## When to use
Use for selection decisions across criteria: model choice, vendor review,
framework sign-off or control assessment. The pattern supports numeric scoring
or RAG statuses and adds a compact aggregate row.

## Function signature
```python
from patterns import scorecard_slide

scorecard_slide(
    title: str,
    criteria: list[str],
    options: list[str],
    scores: list[list[float | str]],
    weights: list[float] | None = None,
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import scorecard_slide

b = scorecard_slide(
    title="Candidate Model Selection Scorecard",
    criteria=["PMSE accuracy", "Stability", "Explainability", "Runtime", "Governance"],
    options=["Baseline", "Model B", "Ensemble"],
    scores=[[2, 4, 5], [4, 4, 4], [5, 4, 3], [5, 4, 2], [5, 4, 3]],
    weights=[0.35, 0.20, 0.15, 0.10, 0.20],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Model Scorecard")
deck_path = project.output_path("model_scorecard.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Use either numeric scores throughout or RAG strings throughout; never mix modes.
2. Numeric scores should share a scale, normally 1-5 with higher meaning better.
3. Provide weights only when the decision method supports them; weights must sum to 1.
4. Keep criteria mutually distinct and concise.
5. Use 2-5 options; beyond that the comparison should move to an appendix or workbook.
