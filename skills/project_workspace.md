# Workflow Skill: Project Workspace

## When to use
Use for any real team deliverable: a deck, document, workbook, or a linked set
of artifacts. Do not ask a non-technical user to manage these folders manually;
create them as part of fulfilling the request.

## Folder contract
```text
projects/<project_name>/
  inputs/    User-provided DOCX, XLSX, data extracts and briefs
  assets/    Charts, figures and referenced images
  working/   Extracted context, decision log and mandatory slide plan
  outputs/   Internal scaffolds and final delivered artifacts
  qa/        Automatic QA reports
  scripts/   Project-specific refined build scripts when needed
  project.json  Folder map and workflow run manifest
```

## Required workflow
Initialize or reuse a workspace:

```python
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Treasury Spread Model")
deck_path = project.output_path("results_recommendations.pptx")
qa_path = project.qa_path(deck_path.name)
```

Use `project.ingest_input(path)` when a user supplies a source file outside the
project directory. Save final artifacts to `outputs/` and checks to `qa/`.
Reserve top-level `output/` for pattern demos and backwards-compatible scripts.

For the Word intake stage, use `project.record_workflow(...)` so `project.json`
records the scaffold, extracted context and QA evidence. For an AI-authored
refined deck, use `project.register_output(...)` after its QA repair loop so
the user-facing artifact and its `working/slide_plan.md` evidence are clearly
recorded. Include the rendered-review status report alongside structural QA:

```python
from artifact_utilities import ArtifactUtilities

rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
project.register_output(
    deck_path,
    qa_path,
    patterns=[...],
    working_files=[project.working_path("slide_plan.md")],
    qa_report_paths=[rendered_qa_path],
)
```
