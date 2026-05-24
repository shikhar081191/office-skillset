# Workflow Skill: Advanced Artifact Utilities

## When to use
Use when a task needs a retained advanced capability beyond the lightweight
builders: DOCX templates, comments, tracked revisions, redaction, accessibility,
table conversion, rendered slide/document QA, font checks, montages or an
artifact-tool example.

## Required workflow
1. Create or reuse a `ProjectWorkspace`.
2. List supported utilities or runtime availability when uncertain:

```bash
python artifact_utilities.py list
```

3. Use `ArtifactUtilities(project)` for project-aware helper functions, or run
   any registered utility ID through the facade:

```python
from artifact_utilities import ArtifactUtilities
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Review")
tools = ArtifactUtilities(project)
tools.run("docx.a11y_audit", [project.output_path("note.docx")], check=False)
```

4. Save intermediate conversions in `working/`, final deliverables in
   `outputs/`, and utility logs/previews in `qa/`.
5. If an optional utility is unavailable, report the required runtime clearly;
   do not silently skip a requested rendered or artifact-tool operation.

For a Word-to-PowerPoint request, these utilities do not replace refinement:
the DOCX scaffold remains internal, and the assistant must create a patterned
refined deck and repair actionable PPTX QA findings before delivery.

## Optional visual QA
`slides.render_slides`, `slides.slides_test`, `docx.render_docx` and
`docx.render_and_diff` require the optional rendering runtime. Structural QA
from `pptx_qa.py`, `docx_qa.py` and `xlsx_qa.py` remains mandatory regardless.

For final refined PowerPoint decks, call:

```python
rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
```

This is workflow-safe for non-technical users: it writes a QA status report
even when rendering cannot run, and automatically creates previews, a montage
and a rendered overflow smoke test when its runtime is available.
