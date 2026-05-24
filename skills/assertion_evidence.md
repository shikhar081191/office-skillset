# Pattern: Assertion Evidence

## When to use
Use when one research conclusion should dominate a slide and a short evidence
list should substantiate it. This is useful for executive findings, sign-off
recommendations and narrative turning points in technical decks.

## Function signature
```python
from patterns import assertion_evidence_slide

assertion_evidence_slide(
    title: str,
    assertion: str,
    evidence: list[dict],
    builder=None,
    palette=None,
) -> PptxBuilder
```

## Example call
```python
from patterns import assertion_evidence_slide

b = assertion_evidence_slide(
    title="Core Finding",
    assertion="Bills drive almost all of the PMSE improvement",
    evidence=[
        {"stat": "46pp", "text": "lower PMSE in short-duration sovereign cohorts"},
        {"stat": "p<0.01", "text": "effect persists after regime controls"},
        {"stat": "82%", "text": "of aggregate gain is concentrated in bill exposures"},
    ],
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Core Finding")
deck_path = project.output_path("core_finding.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. Assertion copy should be a single declarative conclusion of about 10-20 words.
2. Include 3-4 pieces of evidence; each should be independently understandable.
3. Put the strongest numeric evidence first.
4. Use `stat` only for a short number or significance marker, not a sentence.
5. Avoid hedged language unless uncertainty is itself the decision-relevant finding.
