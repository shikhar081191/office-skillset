# Office File Toolkit — AI Assistant Instructions

These instructions guide AI assistants (GitHub Copilot, Windsurf, Cursor, etc.)
when working with .docx, .pptx, and .xlsx files in this repo.
All workflows use pure Python — no npm, no Node.js, no LibreOffice required.

---

## HARD CONSTRAINTS — read these before doing anything else

**NEVER do any of the following, regardless of what the user asks:**

- **NEVER call `patterns.py` functions or `PptxBuilder` directly as the first step** when the user supplies a `.docx` source file. Always run `deck.py` first to produce the structural first pass.
- **NEVER write a build script from scratch** when a Word source file exists. Run `deck.py`; then write a refinement script only to replace `content_slide` calls with proper patterns.
- **NEVER use `content_slide` as a default for narrative sections.** Every section must map to the best-fit pattern from the skills table (Step 1). Only fall back to `content_slide` if no pattern fits and the section cannot be distilled.
- **NEVER discard source content during Word intake.** Keep the copied DOCX and complete captured context; an internal scaffold may paginate long content but must not silently trim paragraphs, table rows or columns.
- **NEVER start refined deck code** before (a) reading `skills/PATTERN_CATALOG.md` and the relevant detailed skill files, (b) evaluating pattern choices for each planned slide, and (c) recording the plan in `working/slide_plan.md` and `decision_log.md`.
- **NEVER deliver a deck to the user** before the refined deck has passed its QA repair loop (`b.save(..., final=True, report_path=...)`) and recorded rendered-review status with `ArtifactUtilities(project).review_rendered_pptx(deck_path)`. The Word-source scaffold is internal only. Fix errors and every actionable warning before reporting the refined output path.
- **NEVER invent project folder paths.** All outputs go inside a `ProjectWorkspace` — never write to top-level `output/` for real deliverables.

**ALWAYS do the following:**

- **ALWAYS use `deck.py` as the single entry point** when the user provides a `.docx` source. Do not bypass it.
- **ALWAYS read `skills/PATTERN_CATALOG.md`** before selecting layouts. It is the complete AI decision catalog of supported presentation patterns and supporting utilities.
- **ALWAYS read `skills/REFERENCE_GALLERY.md`** before selecting layouts or setting font sizes. It contains annotated examples of good slides — match their font hierarchy, space coverage, and colour use. Key benchmarks: 44 pt for deck titles, 26–28 pt for slide titles, 13 pt for card headings, 11–12 pt for body text. Never use font sizes smaller than these.
- **ALWAYS produce a slide plan** (source evidence → assertion → selected pattern → rationale → inputs/assets → inference flag) and record it via `DecisionLog.record_slide_plan()` before writing refined deck code. This creates `working/slide_plan.md` and records the decision in `decision_log.md`.
- **ALWAYS apply `register_output()`** at the end of any manual refinement script so the output is recorded in `project.json`.
- **ALWAYS read `working/user_instructions.txt`** if it exists before building any slides — it overrides default pattern choices.
- **ALWAYS read the refined deck QA report, interpret warnings, revise the code/content for real visual issues, and rerun QA before delivery.**
- **ALWAYS run `ArtifactUtilities(project).review_rendered_pptx(deck_path)` on the final refined deck.** If rendering is available, inspect the generated slide PNGs; if unavailable, retain the explicit skipped report rather than pretending a rendered review occurred.

---

## One-liner workflow for PowerPoint

```text
user pastes bullets or a table
  OR provides a Word file with notes and/or tables
  -> AI reads skills/PATTERN_CATALOG.md and relevant detailed skills
    -> for chat content, AI writes and runs a script using patterns.py
    -> for a Word source, AI runs `deck.py` to create a checked scaffold
      -> AI writes `working/slide_plan.md`, selects patterns, builds the refined deck
        -> AI fixes QA errors/actionable warnings
          -> AI records rendered-review status and inspects previews when available
            -> passing refined file appears in `projects/<project>/outputs/`
```

**Always follow these steps, including QA and any needed repair before delivery.**

---

## Brief-based builds (no Word source file)

When the user provides a plain-text brief or a filled `BRIEF_TEMPLATE.md`
instead of a `.docx`, skip the `deck.py` intake step and follow
`AI_INSTRUCTIONS.md` directly — it has the complete brief → deck workflow.

| File | Purpose |
|------|---------|
| `AI_INSTRUCTIONS.md` | Step-by-step brief → deck workflow (parse → plan → build → QA) |
| `STORY_TEMPLATES.md` | Per-slide content rules and code scaffolds for 5 deck types |
| `BRIEF_TEMPLATE.md` | Standard brief format; users paste this + their notes into chat |

---

## Step 1 — Inventory the pattern choices first

Before writing refined deck code, read these three files in order:

1. **`skills/REFERENCE_GALLERY.md`** — visual examples of approved good slides.
   Use the font sizes, colour palettes, and layout density shown there as your
   target standard. Do not produce slides with smaller fonts or more vacant space
   than the examples show.

2. **`skills/PATTERN_CATALOG.md`** — all available patterns and selection cues.

Then read the detailed skill files relevant to plausible choices:

| File                       | Pattern             | Use when…                                    |
|----------------------------|---------------------|----------------------------------------------|
| `skills/status.md`         | Status grid         | Workstream RAG dashboard, programme status   |
| `skills/process.md`        | Process strip       | Step-by-step flow, lifecycle, methodology    |
| `skills/timeline.md`       | Timeline            | Roadmap, milestones, delivery schedule       |
| `skills/results.md`        | Results comparison  | Model comparison, scenario analysis, A/B     |
| `skills/chart_context.md`  | Chart + context     | Chart with headline, bullets, so-what        |
| `skills/numbers.md`        | Big numbers         | 3–4 headline stats, executive summary        |

Additional research patterns:

| File | Pattern | Use when... |
|------|---------|-------------|
| `skills/heat_map.md` | Heat map | Numeric matrix, PMSE/exposure concentration |
| `skills/waterfall.md` | Waterfall | Contribution to P&L or model improvement |
| `skills/scorecard.md` | Scorecard | Model/vendor selection criteria |
| `skills/annotation_chart.md` | Annotation chart | Chart events and regime callouts |
| `skills/two_by_two.md` | Two by two | Impact/effort or risk prioritisation |
| `skills/assertion_evidence.md` | Assertion + evidence | One conclusion supported by facts |
| `skills/diverging_bar.md` | Diverging bar | Paired metric comparison by cohort |
| `skills/slide_qa.md` | Quality gate | Mandatory check after every PPTX build |
| `skills/docx_to_deck.md` | Word source to deck | Source DOCX has ideas, text or tables |
| `skills/project_workspace.md` | Project workspace | Group real inputs, outputs and QA files |
| `skills/artifact_workflow.md` | Unified workflow | Manage Word intake, refinement and QA-repair delivery |
| `skills/advanced_artifact_utilities.md` | Advanced utilities | Templates, audits, redlines, render QA or conversions |
| `skills/layout_density.md` | Layout density | Sparse slides, text-only font sizing, visual-to-text balance |
| `skills/layout_symmetry.md` | Layout symmetry | Text overflow, shape overlap, auto-fit, AI content labeling |
| `skills/recommendation.md` | Recommendation | Decision slide with dark panel and rationale |
| `skills/kpi_dashboard.md` | KPI dashboard | 3–6 headline metric cards |
| `skills/agenda.md` | Agenda | Numbered section list with active-item highlight |
| `skills/assumption_table.md` | Assumption table | Model/scenario assumptions with sensitivity coding |
| `skills/bar_chart.md` | Bar chart | Category comparison with native pptx chart |
| `skills/line_chart.md` | Line chart | Time-series trend with native pptx chart |
| `skills/PATTERN_CATALOG.md` | Mandatory catalog | Complete pattern and utility selection surface |

Visual-bar patterns (full-width header bar, prm-deck-kit style — see `AI_INSTRUCTIONS.md`):

| Pattern function | When to use | Cap |
|-----------------|-------------|-----|
| `three_column_card_slide()` | Team slides, 3-feature breakdown — only when content is genuinely 3 parallel items with real body text in all cards | Max 1 per deck |
| `two_column_contrast_slide()` | Problem vs solution, before vs after — one use early in deck to frame the narrative | Max 1 per deck |
| `numbered_steps_slide()` | Sequential pipeline, onboarding flow, governance process with 3–6 distinct steps | Max 1 per deck |
| `callout_bar_slide()` | Closing statement — final or penultimate slide only; not a generic bullet slide with a tagline | Max 1 per deck |

**No single visual-bar pattern may appear more than once per deck.** Combined visual-bar
patterns should not exceed half the content slides (title excluded) — for an 8-slide deck
that means a maximum of 3–4 visual-bar slides. Prefer data-rich patterns
(`numbers_slide`, `kpi_dashboard_slide`, `results_slide`, `heat_map_slide`,
`bar_chart_slide`) when the content is quantitative. Use `assertion_evidence_slide`
or `recommendation_slide` for single-argument slides rather than card layouts.

If no pattern fits, fall back to `PptxBuilder` directly (see PPTX section below).

---

## Step 2 — Match content to pattern

Apply the content-shaping rules from the relevant skill file:
- Extract the right fields from the user's raw text
- Distil status lines to < 80 chars
- Assign RAG: green = on track, amber = at risk, red = blocked
- Pick the right column indices for win_col / worst_col
- Set highlight= for the step the user wants to call out
- When input is a `.docx`, run `artifact_workflow.py`; it extracts context,
  builds a checked first deck and records the workflow run
- Use `build_deck_from_docx.py` as the lower-level builder only when a custom
  assistant script needs direct control of the sequence

**Never use `content_slide` as the default for narrative text.** For every
section — whether the input is a Word file or pasted chat content — pick the
best-fit pattern from the table above. Only fall back to `content_slide` if
no pattern fits and the content cannot be distilled into any of them.

---

## Step 3 — Write and run the script

When the user supplies a Word source for a deck, begin with the intake entry point:

```bash
python deck.py brief.docx --audience "risk committee" --decision "approve model v2"
python deck.py brief.docx --palette midnight_executive --instructions "Lead with the recommendation."
```

This wraps `artifact_workflow.py`, prints readable QA output, and opens the checked
scaffold automatically. It is intake for refinement, not the finished editorial deck.
For full control over project naming or output path:

```bash
python artifact_workflow.py --source model_brief.docx --output-type pptx --project "Model Review" --audience "board" --decision "approve budget" --instructions "Emphasise the recommendation."
```

**This produces a structural first pass only — every narrative section becomes
a bullet slide.** After running it:
1. If `working/user_instructions.txt` exists, read it first and apply the editorial direction throughout.
2. Read the extracted context in `working/` — sections marked `[DISTIL]` must be rewritten into concise slide copy, not copied verbatim.
3. Read `skills/PATTERN_CATALOG.md`, evaluate the plausible patterns for each
   audience-facing slide and produce a plan before writing any refined code.
4. **Record the slide plan using `DecisionLog.record_slide_plan()`**. It writes
   `working/slide_plan.md` and appends the plan to `working/decision_log.md`:
```python
from decision_log import DecisionLog
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Project Name", create=False)
log = DecisionLog(project.working_path("decision_log.md"))
log.record_slide_plan([
    {"section": "Introduction", "source": "Executive Summary",
     "pattern": "recommendation_slide", "assertion": "Approve model v3 for Q3 deployment",
     "rationale": "A committee decision is explicitly requested",
     "inputs": "Recommendation; validation evidence", "inference": "No"},
    {"section": "Model Performance", "source": "Performance table",
     "pattern": "results_slide", "assertion": "PMSE improved on every cohort",
     "rationale": "Exact cross-model values matter",
     "inputs": "PMSE table", "inference": "No"},
])
```
5. Write a refined build script that implements the plan with proper patterns.
6. Run its QA report, fix errors and actionable warnings, regenerate and repeat until the final visual deck is suitable for delivery.
7. Run `ArtifactUtilities(project).review_rendered_pptx(deck_path)`; inspect generated PNG slides if available, or retain its explicit optional-runtime skip report.

It records the source, output, QA report, palette and patterns used in the
project manifest. For content supplied directly in chat, write a short
pattern-based build script.

Write a short Python script (10–30 lines) and run it:

```python
from patterns import status_slide   # or process_slide, timeline_slide, etc.
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Q2 Programme Status")

b = status_slide(
    title="Q2 2025 Programme Status",
    workstreams=[
        {"name": "Data Pipeline",  "status": "On track; UAT 3 Jun",  "rag": "green", "owner": "AK"},
        {"name": "Risk Models",    "status": "2w delay, vendor data", "rag": "amber", "owner": "SP"},
    ]
)
deck_path = project.output_path("q2_status.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
```

Run with: `python <script_name>.py`

---

## Step 4 — Validate, fix and report

Final slide scripts must save inside a project workspace with
`b.save(..., final=True, report_path=project.qa_path(...))`.
The save performs the core structural and legibility audit automatically. If
the quality gate fails, fix the script, regenerate, and rerun it. It also
flags uneven repeated panel sizing/spacing as a layout-symmetry warning. For a
final refined deck, record rendered-review status through
`ArtifactUtilities(project).review_rendered_pptx(deck_path)` and inspect slide
PNGs when generated. Tell the user the output path only after this review loop.

---

## PPTX — PptxBuilder (direct use)

```python
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace

b = PptxBuilder(palette="midnight_executive")
b.title_slide("My Presentation", subtitle="Subtitle here")
b.content_slide("Slide Title", body=["Bullet one", "Bullet two"])  # last resort only — prefer a named pattern
b.two_column_slide("Comparison", "Option A", ["Fast", "Cheap"], "Option B", ["Reliable"])
b.stat_slide("Key Numbers", [("$2.3B", "AUM"), ("0.94", "Sharpe")])
b.section_divider("Appendix")
project = ProjectWorkspace("My Presentation")
deck_path = project.output_path("output.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

**Available palettes:** `prm` (navy/teal/amber — recommended for financial decks),
`blackrock`, `midnight_executive`, `coral_energy`, `teal_trust`,
`charcoal_minimal`, `warm_terracotta`

BlackRock is the default. For project- or client-specific colours, pass a
palette dictionary to `PptxBuilder` or call `register_palette()` from
`create_pptx.py`. Patterns always follow the active builder palette.

**Design rules:**
- Every content slide must have ≥1 visual element (table, chart image, or structural pattern); read `skills/layout_density.md` before building any slide
- Text-only fallback is only acceptable with body font ≥18pt — never small text on an empty canvas
- `content_slide` auto-scales font by bullet count (≤3 → 20pt, ≤5 → 16pt, 6+ → 14pt); if the result is still sparse, switch to a visual pattern
- One colour dominates (60–70% visual weight)
- Dark backgrounds for title + conclusion, light for content
- Never centre body text — only centre titles and stat numbers
- Title font: 22–44pt bold; body 18–22pt for text-only slides, 9–14pt when visuals carry the canvas
- **Use `key_message` on every slide where the slide has a clear single-sentence takeaway** — it renders as a bold banner and prevents visible blank space below the title
- **Fill all optional pattern fields when content is available** — `tag` in cards, `so_what` in chart_context, `footnote` in numbered_steps; blank optional fields leave empty space
- **A slide with large empty areas means the wrong pattern was chosen** — merge the content, switch to a richer pattern, or drop the slide
- **No single visual-bar pattern may appear more than once per deck; combined visual-bar patterns should not exceed half the content slides**
- **Card and contrast patterns require dense body text** — `three_column_card_slide` and `two_column_contrast_slide` need ≥40 words per card/panel. If you cannot write that much real content for each section, use `assertion_evidence_slide` or `callout_bar_slide` instead. Short cards always look vacant.
- **Step descriptions must be full sentences** — `numbered_steps_slide` step descriptions must be ≥10 words each. Fragment descriptions (3–5 words) leave bare rows. Use `process_slide` instead if descriptions are short.
- **Font size is set by the pattern, not by you** — do not cram content to fit; split or switch patterns instead. Small fonts on slides are always a sign the wrong pattern was chosen or too much content was forced into one slide.

---

## DOCX — DocxBuilder

```python
from create_docx import DocxBuilder
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Report")
b = DocxBuilder()
b.heading("Report Title", level=1)
b.paragraph("Body text here.", font_size=12)
b.bullet_list(["Point one", "Point two", "Point three"])
b.table(
    data=[["Header A", "Header B"], ["Row 1A", "Row 1B"]],
    header_row=True,
    col_widths=[2.5, 2.5],
    header_color="D5E8F0",
)
doc_path = project.output_path("output.docx")
b.save(doc_path, final=True, report_path=project.qa_path(doc_path.name))
```

Rules: never use unicode bullets directly — use `bullet_list()`. Never use `\n`
for line breaks — use separate paragraphs. Page size is US Letter automatically.
Read `skills/research_note_docx.md` for the reusable research-note pattern and
`skills/document_qa.md` for the mandatory DOCX check.

---

## XLSX — XlsxBuilder

```python
from create_xlsx import XlsxBuilder, FMT_CURRENCY, FMT_PCT, FMT_INT
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Workbook")
b = XlsxBuilder()
s = b.add_sheet("Model")
s.headers(1, ["Factor", "Return", "Vol", "Sharpe"], col_width=14)
s.input_cell(2, 2, 0.032, FMT_PCT)        # blue — hardcoded input
s.formula_cell(2, 4, "=B2/C2", FMT_PCT)   # black — calculated
s.freeze_panes("B2")
book_path = project.output_path("output.xlsx")
b.save(book_path, final=True, report_path=project.qa_path(book_path.name))
```

Financial colour conventions (always follow):
- Blue `0000FF` — hardcoded inputs
- Black `000000` — all formulas
- Green `008000` — cross-sheet links
- Yellow fill `FFFF00` — key assumptions

Read `skills/pmse_workbook.md` for the reusable research workbook pattern and
`skills/workbook_qa.md` for the mandatory XLSX check.

---

## Retained advanced resources

Do not delete or replace `artifact_python_utilities_chatgpt/`. Every retained
Python utility is exposed through `artifact_utilities.py`. Use
`python artifact_utilities.py list` to inspect capabilities and runtime status,
and read `docs/ADVANCED_UTILITY_CATALOG.md` when using templates, comments,
redlines, DOCX accessibility/privacy tools, table conversion, montages,
rendered QA or optional artifact-tool workflows.

For PowerPoint delivery, `pptx_qa.py` is the mandatory dependency-light gate.
Every refined delivery records a rendered-review attempt through
`ArtifactUtilities(project).review_rendered_pptx(deck_path)`; it creates
previews and a rendered overflow smoke test when its optional runtime is
installed, otherwise it records exactly why this phase was skipped.
`docx_qa.py` and `xlsx_qa.py` provide the corresponding structural gates for
Word and Excel deliverables.

When a user explicitly requests advanced visual rendering or artifact-tool
functionality and the catalog reports a missing optional runtime, explain what
must be installed rather than silently omitting that operation.

## General rules

- Use `pathlib.Path` over `os.path`
- Real deliverables go in `projects/<project>/outputs/`, with source material
  in `inputs/`, supporting figures in `assets/`, and reports in `qa/`
- Reserve top-level `output/` for reusable demos and compatibility smoke tests
- Never calculate in Python what should be an Excel formula
- Never hardcode colours in patterns; always reference builder palette keys,
  including semantic roles such as `surface`, `border`, `positive`, `negative`
  and `highlight`
- Use `defusedxml` for XML parsing
- Call `b.ai_label()` immediately after any slide containing synthesised, inferred, or hallucinated content not present in the user's source; see `skills/layout_symmetry.md`
- Call `b.source_label("Source name")` when slide content originates from a named external source
- Call `b.speaker_notes("...")` to record presenter talking points, anticipated questions, or verbal caveats on any slide where the presenter needs guidance
- Pass `auto_fit=True` to `b._add_text_box()` in custom scripts whenever text length is uncertain — never force text into a box by dropping font below 8pt
- Read `skills/layout_symmetry.md` whenever building slides with multiple text boxes to avoid overflow and overlap
- When the user supplies `--audience` or `--decision` context, shape every slide's assertion and emphasis toward that audience and toward making that decision legible
