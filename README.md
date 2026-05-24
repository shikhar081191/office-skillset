# Office File Toolkit

Core Python builders plus nineteen reusable PowerPoint slide patterns for
research storytelling, model comparison and executive communication.

Pure-Python builders for .pptx, .docx, and .xlsx — no npm, no Node.js,
no LibreOffice for core generation. Plus nineteen reusable PowerPoint slide patterns ready to call
from an AI assistant or a script.

---

## Setup

```bash
pip install -r requirements.txt
```

---

## One-liner workflow (AI-assisted)

```text
paste bullets or a table into your AI assistant
  or attach a Word file containing notes and/or data tables
  -> it reads skills/ and picks the right pattern
    -> for chat content it writes a script using patterns.py
    -> for a Word source it runs deck.py to extract context and make an internal scaffold
      -> it plans the story and builds a refined patterned deck
        -> it reads the QA report, fixes errors and actionable warnings, and rebuilds
          -> it records rendered-review status and inspects previews when available
            -> the final refined file appears in projects/<project>/outputs/
```

The AI instructions live in `.github/copilot-instructions.md` (GitHub Copilot)
and `.windsurfrules` (Windsurf / Cursor) — both contain identical content.
For a non-technical overview, read [`EXPLANATION.md`](EXPLANATION.md). For
prompt examples and practical user guidance, read
[`BEST_PRACTICES.md`](BEST_PRACTICES.md).
For completed capabilities and the recommended expansion roadmap, read
[`REMAINING_PLAN.md`](REMAINING_PLAN.md).

---

## Project workspaces

For real team deliverables, the toolkit now groups source material and results
under a single project folder:

```text
projects/<project_name>/
  inputs/                  Supplied Word files, spreadsheets and data
  assets/                  Charts and images used in deliverables
  working/
    decision_log.md        ← Human-readable log of AI decisions (read this!)
    slide_plan.md          ← Pattern-by-pattern plan with evidence and rationale
    *_source.md            Extracted content from the Word source
    *_source.json          Structured extraction for downstream scripts
    user_instructions.txt  Audience, decision, and editorial direction
  outputs/                 Internal scaffold plus delivered refined files
  qa/                      Automatic quality reports (JSON)
  project.json             Folder contract plus reproducible workflow run history
```

```python
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Treasury Spread Model")
deck_path = project.output_path("results_recommendations.pptx")
```

The assistant creates and uses this structure automatically. The top-level
`output/` directory remains available for slide-pattern demos and older scripts.
Project-specific refinement scripts belong inside the related project, for
example `projects/q2_model_review/scripts/`. Reusable scenario builds and
sample inputs belong under `examples/`.

---

## Word brief to presentation

A user can provide a `.docx` file containing unstructured thinking, talking
points, and tables. `deck.py` is the intake command used by the AI assistant:
it creates a checked internal scaffold, extracted context and decision log.
That scaffold is not the presentation delivered to the user; the assistant
must then build and QA a refined visual deck:

```bash
python deck.py model_brief.docx
python deck.py model_brief.docx --audience "risk committee" --decision "approve model v2"
python deck.py model_brief.docx --palette midnight_executive --instructions "Lead with the recommendation."
```

For full control over project naming or output path use `artifact_workflow.py` directly:

```bash
python artifact_workflow.py --source model_brief.docx --output-type pptx --project "Model Brief"
python artifact_workflow.py --source model_brief.docx --project "Model Brief" --palette blackrock --audience "board" --decision "approve budget" --instructions "Emphasise the PMSE recommendation."
```

```python
from source_docx import read_source_docx, source_to_markdown

source = read_source_docx(project.input_path("model_brief.docx"))
assistant_context = source_to_markdown(source)
```

This is intentionally a two-stage assistant workflow rather than a rigid
converter: the first pass preserves source structure; the AI assistant then
reads `skills/PATTERN_CATALOG.md`, rewrites long notes into concise slide
language, records `working/slide_plan.md` with its pattern choices and reasons,
and builds the refined presentation with the best storytelling patterns. It reads
the refined QA report and repairs real layout, contrast and density warnings
before reporting the final PPTX path.

The intake files under `working/` retain complete captured narrative and table
content in source order, plus detected header/footer and text-box text. The
original Word file is also copied into `inputs/`. The internal scaffold
paginates long material across continuation slides; it must not silently trim
paragraphs, rows or columns merely to fit a preview slide.

The lower-level checked starter builder remains available for custom assistant
scripts or backwards-compatible use:

```bash
python build_deck_from_docx.py model_brief.docx --project "Model Brief"
```

The catalog at [`skills/PATTERN_CATALOG.md`](skills/PATTERN_CATALOG.md) lists
every supported slide pattern and the utility-selection cues the AI must
consider before it writes a refined deck.

To initialize a project before adding input files:

```bash
python project_workspace.py "Treasury Spread Model"
```

---

## Available patterns

| Function               | Pattern            | When to reach for it                                |
|------------------------|--------------------|-----------------------------------------------------|
| `status_slide()`       | Status grid        | 4–8 workstream RAG cards for programme updates      |
| `process_slide()`      | Process strip      | 3–5 step linear flow with optional highlighted step |
| `timeline_slide()`     | Timeline           | Roadmap or milestone schedule, past/current/future  |
| `results_slide()`      | Results comparison | Model/scenario table with best and worst highlights |
| `chart_context_slide()`| Chart + context    | Chart image paired with headline, bullets, so-what  |
| `numbers_slide()`      | Big numbers        | 3–4 oversized stats on dark background              |

Additional research-oriented patterns:

| Function | Pattern | When to reach for it |
|----------|---------|----------------------|
| `heat_map_slide()` | Heat map | Numeric matrix such as PMSE by model and cohort |
| `waterfall_slide()` | Waterfall | Positive/negative contributions to a net result |
| `scorecard_slide()` | Scorecard | Criteria-based model or vendor selection |
| `annotation_chart_slide()` | Annotation chart | Chart with event callouts and a finding |
| `two_by_two_slide()` | Two by two | Impact/effort or probability/severity prioritisation |
| `assertion_evidence_slide()` | Assertion + evidence | One conclusion supported by concise facts |
| `diverging_bar_slide()` | Diverging bar | Paired comparison across cohorts or models |
| `recommendation_slide()` | Recommendation | Decision slide — dark panel with rationale and caveats |
| `kpi_dashboard_slide()` | KPI dashboard | 3–6 metric cards with value, delta, and trend arrow |
| `agenda_slide()` | Agenda | Numbered section list with active-item highlight |
| `assumption_table_slide()` | Assumption table | Model/scenario assumptions with sensitivity colour coding |
| `bar_chart_slide()` | Bar chart | Native clustered column chart for category comparisons |
| `line_chart_slide()` | Line chart | Native line chart for time-series trends |

All patterns default to the **BlackRock palette**. Every pattern uses the active
builder palette, so a whole deck can switch styling without changes to pattern code.
Pass a named palette to `PptxBuilder`, pass an external palette dictionary, or
register a reusable team palette:

```python
from create_pptx import PptxBuilder, register_palette

register_palette("team_brand", {
    "primary": "113355",
    "secondary": "29A3A3",
    "accent": "F2C14E",
    "text_dark": "13202B",
    "text_light": "FFFFFF",
    "negative": "B42318",   # optional semantic override
})
b = PptxBuilder(palette="team_brand")
```

---

## Quick example

```python
from patterns import status_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Q2 Status")

b = status_slide(
    title="Q2 2025 Programme Status",
    workstreams=[
        {"name": "Data Pipeline",    "status": "On track; UAT starts 3 Jun",       "rag": "green", "owner": "AK"},
        {"name": "Risk Models",      "status": "2-week delay — vendor data late",   "rag": "amber", "owner": "SP"},
        {"name": "Compliance",       "status": "Legal sign-off still outstanding",  "rag": "red",   "owner": "TJ"},
    ]
)
deck_path = project.output_path("q2_status.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

Run the full demo (one slide per pattern):

```bash
python patterns.py
```

Output files appear in `output/`.

---

## Combining patterns into one deck

```python
from patterns import status_slide, timeline_slide, numbers_slide, BLACKROCK
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Q2 Review")
b = PptxBuilder(palette=BLACKROCK)
b.title_slide("Q2 2025 Review", subtitle="Programme Risk & Delivery")

status_slide("Status", workstreams=[...], builder=b)
timeline_slide("Roadmap", milestones=[...], builder=b)
numbers_slide("At a Glance", stats=[...], builder=b)

deck_path = project.output_path("q2_review.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

---

## Automatic PowerPoint quality gate

For end users, there is no extra step: provide the content in chat and receive
the checked presentation. Internally, `PptxBuilder.save()` always runs the
lightweight structural audit in `pptx_qa.py`. Final-delivery scripts pass
`final=True` and write a QA record alongside the deck:

```python
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

The mandatory gate blocks slides with off-canvas objects or unfinished
placeholder text, and flags likely text crowding, overlap, sparse layouts,
small type, low contrast, low-resolution images, overly dense tables, or
uneven spacing/dimensions in aligned repeated panels.
The assistant must inspect those warnings, fix any that indicate a real visual
problem, and rerun QA before delivering a refined deck. A scaffold is never a
finished Word-to-deck deliverable.

Three optional badges can be added to any slide after building it:

```python
b.ai_label()                        # red badge: AI-synthesised content — verify before use
b.source_label("Bloomberg, Q2 25")  # grey italic: data attribution bottom-left
b.speaker_notes("Stress the out-of-time validation result here.")  # notes pane
```
Final refined builds also record the rendered-review phase:

```python
from artifact_utilities import ArtifactUtilities
rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
```

When its optional LibreOffice/Poppler runtime is installed, this creates slide
PNGs, a montage and a rendered overflow smoke test for assistant inspection.
Otherwise it records why rendered inspection was skipped; it does not hide the
missing capability from the user.

---

## Other builders

| Import                          | What it does                            |
|---------------------------------|-----------------------------------------|
| `from create_docx import DocxBuilder` | Create Word documents             |
| `from docx_patterns import research_note_document` | Research-note DOCX pattern |
| `from create_pptx import PptxBuilder` | Create PowerPoint presentations   |
| `from create_xlsx import XlsxBuilder` | Create Excel workbooks            |
| `from xlsx_patterns import pmse_comparison_workbook` | PMSE workbook pattern |
| `from source_docx import read_source_docx` | Read Word input for deck authoring |
| `from project_workspace import ProjectWorkspace` | Group inputs, outputs and QA by project |
| `from decision_log import DecisionLog` | Append-friendly log of AI decisions readable by the user |
| `python deck.py brief.docx` | AI intake: extract content and build an internal scaffold for refinement |
| `python artifact_workflow.py --source brief.docx --output-type pptx --project "Brief"` | Lower-level DOCX intake/scaffold controller |
| `python build_deck_from_docx.py brief.docx` | Build a checked starter deck from Word |
| `python unpack.py file.docx out/`     | Unzip Office file to editable XML |
| `python pack.py out/ file.docx`       | Rezip XML back to Office file     |

DOCX and XLSX builders now mirror the PPTX delivery convention:

```python
doc_path = project.output_path("research_note.docx")
doc.save(doc_path, final=True, report_path=project.qa_path(doc_path.name))
book_path = project.output_path("pmse.xlsx")
book.save(book_path, final=True, report_path=project.qa_path(book_path.name))
```

## Retained advanced utilities

The supplemental utilities in `artifact_python_utilities_chatgpt/` are retained
and now exposed through [artifact_utilities.py](artifact_utilities.py). This
provides one catalog and execution facade for advanced DOCX editing/auditing,
PPTX rendering and slide QA utilities, table conversion, montages, and optional
spreadsheet/presentation `artifact_tool` examples.

```bash
python artifact_utilities.py list
python artifact_utilities.py --project "Model Review" run docx.a11y_audit -- projects/model_review/outputs/note.docx
```

Core DOCX utility execution now includes `lxml` in `requirements.txt`.
Rendered slide/document QA and artifact-tool examples remain optional because
they need additional runtimes. Install Python render dependencies with
`requirements-optional-rendering.txt` and configure LibreOffice/Poppler when
the team wants full rendered QA.

See `docs/ADVANCED_UTILITY_CATALOG.md` for the complete integrated utility
catalog and `docs/RESOURCE_INTEGRATION.md` for the overall architecture.

---

## Adding a new pattern

See `skills/ADDING_A_SKILL.md` for the four-step process:
write the function → add a demo → create the skill file → update this table.
