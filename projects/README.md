# Project Workspaces

For the plain-language user journey, see [`../EXPLANATION.md`](../EXPLANATION.md).
For recommended prompts and user best practices, see
[`../BEST_PRACTICES.md`](../BEST_PRACTICES.md).

Each real deliverable should live inside one project folder:

```text
projects/<project_name>/
  inputs/    Source Word files, spreadsheets, data extracts and briefs
  assets/    Charts, images and figures used in outputs
  working/   Extracted context, decision log and AI slide plan
  outputs/   Internal scaffolds and final refined deliverables
  qa/        Automatic quality-check reports
  scripts/   Optional project-specific refined build scripts
  project.json  Folder map and generation-run manifest
```

The person requesting work only needs to provide their material in chat or in
`inputs/`. The assistant creates the remaining folders, may create an internal
scaffold from a Word source, then creates and delivers the refined checked output.

For the Word intake stage, which creates a checked internal scaffold for
AI-assisted refinement:

```bash
python deck.py model_brief.docx --project "Model Brief" --no-open
```

The scaffold is not a final presentation. Before delivery, the assistant must
read `skills/PATTERN_CATALOG.md`, record the visual choices and rationale in
`working/slide_plan.md`, build the refined deck, read and repair its QA
findings, record rendered-review status with
`ArtifactUtilities(project).review_rendered_pptx(deck_path)`, and register the
final output plus both QA records in `project.json`.

Use `output/` only for reusable pattern demos and technical smoke tests.
