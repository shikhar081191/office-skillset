# Resource Integration And Expansion Plan

The toolkit has two complementary layers:

1. The core, lightweight Python API: `create_pptx.py`, `patterns.py`,
   `create_docx.py`, `create_xlsx.py`, and Office package helpers.
2. The retained supplemental bundle in `artifact_python_utilities_chatgpt/`,
   which provides advanced QA, OOXML and optional artifact-tool workflows.

No supplemental skill or utility is removed. The current folder remains a
reference and executable source for future integrations.

`artifact_utilities.py` is the supported facade over every Python utility in
the retained bundle. It discovers all DOCX, slides and spreadsheet utility
entry points, provides project-aware execution logs, and reports optional
runtime readiness. See `docs/ADVANCED_UTILITY_CATALOG.md` for the full list.

## Project Workspace Layer

`project_workspace.py` is the user-deliverable organization layer. Real work
is grouped beneath `projects/<project_name>/` with separate `inputs/`,
`assets/`, `working/`, `outputs/` and `qa/` folders. `project.json` now also
records generation runs, including source inputs, outputs, QA evidence,
palette and patterns used.

Top-level `output/` remains for technical demos and compatibility use.

## Presentation Layer

`patterns.py` is the stable AI-facing pattern catalog. It now contains research
visuals for comparisons, findings, heat maps, annotations and prioritisation.
All patterns follow `PptxBuilder.palette`, enabling a palette supplied from a
project, client or brand configuration without changing slide logic.
`skills/PATTERN_CATALOG.md` is the mandatory assistant-facing selection
registry: before refined authoring, the chat agent evaluates the complete
pattern set and records its chosen pattern, source evidence and rationale in
`working/slide_plan.md`.

`pptx_qa.py` is now the always-on delivery gate for `PptxBuilder.save()`. It
uses only core Python dependencies to catch off-canvas objects and visible
placeholder content before delivery, while flagging likely legibility issues
for correction, including uneven aligned repeated panels. For a final refined
deck, the AI assistant must read the report, repair actionable warnings,
regenerate, and repeat before it supplies the file to the user. Final scripts
should call:

```python
deck_path = project.output_path("deck.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
```

The retained slide utilities are exposed through `artifact_utilities.py` as the
richer rendered-inspection layer:

- `slides/container_tools/render_slides.py` for rendered slide previews.
- `slides/container_tools/create_montage.py` for contact sheets.
- `slides/container_tools/slides_test.py` for overflow smoke tests.
- `slides/artifact_tool_examples/` as optional alternative-engine examples.

These tools currently carry additional renderer/runtime assumptions. Final
refined builds record a rendered-review attempt; when dependencies are absent
the resulting JSON explicitly reports the skipped phase while the mandatory
core `python-pptx` gate still applies.

## Word Source And Output Layer

`deck.py` is the user-friendly intake command for a Word-sourced presentation.
It calls `artifact_workflow.py`, which ingests the supplied Word file into a
project, preserves structured JSON/markdown extraction under `working/`,
creates a checked internal scaffold and writes the intake generation record.
The scaffold is not the delivered deck. The AI assistant must use the source
context and skills to produce a patterned refined deck, run the QA repair
loop, and register that refined output and its `working/slide_plan.md` evidence
in `project.json`. `source_docx.py`
performs extraction and `build_deck_from_docx.py` remains the lower-level
scaffold builder for custom assistant authoring.

`DocxBuilder` remains the basic authoring API, with `docx_patterns.py` adding
research-note and model-validation memo patterns. `docx_qa.py` is an automatic
structural gate on saves. The retained DOCX tools are registered in
`artifact_utilities.py` for template-driven whitepapers and short notes:

- template style application and content controls;
- comments and tracked changes;
- captions, cross-references and navigation;
- accessibility, style and privacy audits;
- redaction and render/diff verification.

Rendered-page inspection remains the richer delivery gate when its optional
runtime is available.

## Spreadsheet Layer

`XlsxBuilder` remains the default lightweight workbook engine.
`xlsx_patterns.py` now supplies factor summary and PMSE comparison workbooks,
and `xlsx_qa.py` runs automatically on saves. Further patterns to add are:

- scenario shock input/output model;
- portfolio exposure heat map;
- model comparison scorecard and validation tracker.

The retained `spreadsheets/` artifact-tool starter is cataloged through
`artifact_utilities.py` as an optional richer rendering/inspection route; it
should be enabled deliberately when the team standardises that runtime.

## Suggested Future Structure

The existing top-level imports should stay compatible while the library grows.
When the next expansion warrants packaging, add internal modules such as:

```text
office_utils/
  pptx/
  docx/
  xlsx/
  shared/
tools/
  slides/
  docx/
```

Top-level files can then act as compatibility wrappers, preventing existing
Copilot/Windsurf examples and user scripts from breaking.
