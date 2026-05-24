# Workflow Skill: Slide Quality Gate

## When to use
Use after every PowerPoint deck build, whether the user requested validation or
not. This is an internal delivery step: the user supplies content and receives
a checked `.pptx` file without needing to run a separate command.

## Required workflow
1. Build the deck from the user's text using `patterns.py` and/or `PptxBuilder`.
2. Save final deliverables with QA enabled and a JSON record:

```python
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Review")
b.save(
    project.output_path("model_review.pptx"),
    final=True,
    report_path=project.qa_path("model_review.pptx"),
)
```

3. Read the JSON report even when saving passes. Interpret every warning as a
   required review item; fix it when it indicates a real layout, readability,
   contrast or density problem.
4. Regenerate and rerun QA after each repair round until no actionable warning
   remains.
5. For a user-facing refined deck, run the rendered-review helper. It records
   whether previews and a rendered overflow smoke test ran, or which optional
   runtime is unavailable:

```python
from artifact_utilities import ArtifactUtilities

rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
```

6. Only report the refined `.pptx` path after this repair loop. Never deliver
   a Word-source scaffold as the finished deck.

## Always-on checks
- No slide object extends outside the slide canvas.
- No visible placeholder or unfinished drafting text remains in final output.
- Warnings identify likely text crowding, very small type, low solid-fill text
  contrast, low-resolution images, overly dense tables, and uneven repeated
  panel spacing or dimensions.

## Judgement and rendered QA status
Warnings are prompts for the assistant to improve a deck where they indicate a
real visual issue; they are not automatic failures because some layouts are
intentionally compact. Always create the rendered-review status file for a
final refined deck. If the optional rendering runtime is installed, inspect
the individual PNG slides before delivery. If it is unavailable, the status
file must say so explicitly; structural QA still runs.
