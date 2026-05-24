# Artifact Skills Summary: PPTX, Excel, and DOCX

This is a practical summary of the available artifact-generation skills and workflows for PowerPoint decks, Excel workbooks, and Word documents.

## Repository Integration Status

The retained Python bundle described in this summary is exposed in this repo
through `artifact_utilities.py`. Run `python artifact_utilities.py list` to
discover all supported retained entry points and see which optional render or
artifact-tool runtimes are available. Project-aware runs write logs and
previews into `projects/<project_name>/qa/`.

See `docs/ADVANCED_UTILITY_CATALOG.md` for the complete catalog and usage
examples. Core DOCX utility paths are supported in the normal environment;
rendered DOCX/PPTX checks and artifact-tool examples require optional runtimes.

Current repository convention: real deliverables live under
`projects/<project_name>/outputs/`, with QA evidence under `qa/`. For a Word
source used to make a presentation, an extracted/scaffold PPTX is internal
only; the AI assistant must produce the refined patterned deck, interpret and
repair actionable QA findings, record rendered-review status through
`ArtifactUtilities(project).review_rendered_pptx(deck_path)`, and then deliver
the refined PPTX.
Before refined PPTX authoring, the assistant must read
`skills/PATTERN_CATALOG.md` and create `working/slide_plan.md`, recording the
source evidence, selected pattern and rationale for each audience-facing slide.

## 1. PowerPoint / PPTX Skill

### What it is used for

Use the slides/PPTX workflow when the task involves creating or editing presentation decks, visual aids, posters, demo screens, pitch decks, charts, or other slide-based artifacts.

### Main creation/editing options

The default slide workflow uses `PptxBuilder` and the research patterns in
`patterns.py`. Retained rendering and artifact-tool examples are available
through `artifact_utilities.py` when their optional runtimes are configured.

### Typical capabilities

I can create or modify decks with:

- Executive-style presentation layouts.
- Pitch decks and hackathon decks.
- One-page strategy slides.
- Visual demo flows and mock product screens.
- Charts, shapes, callouts, diagrams, and structured layouts.
- Theme-consistent styling, spacing, and typography.
- Exportable `.pptx` deliverables.

### Default workflow

1. Understand the slide purpose, audience, decision and desired style.
2. When input is a Word source, run `deck.py` to capture the source and create
   an internal scaffold and decision log.
3. Read `skills/PATTERN_CATALOG.md`, create `working/slide_plan.md`, select
   suitable patterns, distil the text, and build a refined deck in the project
   workspace.
4. Save the refined `.pptx` with `final=True` and a JSON QA report.
5. Interpret the QA findings, repair actionable warnings and regenerate until
   suitable for delivery.
6. Run `ArtifactUtilities(project).review_rendered_pptx(deck_path)`; inspect
   generated PNG slides when available, or retain its explicit skipped report.
7. Return the refined presentation from `projects/<project>/outputs/`.

### Important behavior

For any user request involving slides, presentations, or PPTX files, the
presentation workflow applies. A Word-source scaffold must never be returned
as the finished deck.

---

## 2. Excel / Spreadsheet Skill

### What it is used for

Use the spreadsheet workflow when the task involves creating, editing, analyzing, formatting, or validating spreadsheets such as `.xlsx` or `.xls` files.

Typical use cases include:

- Creating new workbooks.
- Editing uploaded Excel files.
- Building trackers, dashboards, templates, tables, or calculators.
- Adding formulas, formatting, charts, validation, and conditional formatting.
- Aggregating, filtering, pivoting, or analyzing tabular data.
- Recalculating formulas and checking for spreadsheet errors.

### Required tool approach

For this repository, the core workbook route uses `XlsxBuilder`,
`xlsx_patterns.py` and `xlsx_qa.py`. The retained artifact-tool workbook
starter is optional and available only when its runtime is configured.

### Main APIs/concepts

Common operations include:

- Create a workbook.
- Add worksheets.
- Import an existing `.xlsx` file.
- Write values, headers, formulas, and tables.
- Format cells, ranges, rows, and columns.
- Add data validation for editable categorical fields.
- Add conditional formatting.
- Add charts or visual summaries where appropriate.
- Inspect key ranges for values and formulas.
- Scan for formula errors such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`.
- Render a compact preview of important ranges when useful.
- Export the final workbook as `.xlsx`.

### Default workflow

1. Create or import the workbook.
2. Build the workbook quickly: sheets, sections, headers, data, formulas, and key visuals.
3. Apply professional formatting: widths, row heights, headers, borders, number formats, and spacing.
4. Add validation or conditional formatting when useful.
5. Inspect important ranges for correctness.
6. Scan formulas for errors.
7. Render a small preview if visual verification is needed.
8. Export the `.xlsx` file to `projects/<project>/outputs/`.
9. Return the workbook path with a brief summary.

### Quality rules

Spreadsheet outputs should be readable, bounded, and formula-driven where logic is expected. Derived values should usually be formulas rather than hardcoded outputs. Headers, number/date formats, widths, row heights, and borders should be clean. For trackers or planning sheets, at least one visual summary such as a KPI block, dashboard area, or chart is often useful.

### Citation/source handling

When a spreadsheet uses researched or external data, sources should be included in the workbook itself. For row-wise data, this usually means a source URL column. For financial models, source URLs can be added in comments near model inputs.

---

## 3. Word / DOCX Skill

### What it is used for

Use the DOCX workflow when creating, editing, redlining, commenting on, reviewing, merging, cleaning, or formatting Microsoft Word documents.

Typical use cases include:

- Creating polished Word documents.
- Editing uploaded `.docx` files.
- Adding or removing comments.
- Adding tracked changes/redlines.
- Accepting or rejecting tracked changes.
- Cleaning metadata.
- Redacting sensitive content.
- Adding tables, captions, cross-references, hyperlinks, headers, footers, page numbers, watermarks, or navigation.
- Converting spreadsheet tables into Word tables.
- Performing accessibility checks.

### Quality rule

Every final DOCX must pass the core `docx_qa.py` structural check. For formal
or externally shared documents, render and visually inspect pages when the
optional rendering runtime is configured; structural inspection alone may miss
layout defects such as clipping, broken tables or header/footer problems.

### Main utilities

Common DOCX utilities include:

- `render_docx.py`: Render DOCX pages to PNGs for visual QA; can optionally emit a PDF.
- `comments_extract.py`: Extract comments to JSON.
- `comments_add.py`: Add comments by matching paragraph text.
- `comments_apply_patch.py`: Update or resolve comments.
- `comments_strip.py`: Remove comments for final delivery.
- `add_tracked_replacements.py`: Add tracked-change replacements.
- `accept_tracked_changes.py`: Accept tracked changes.
- `redact_docx.py`: Redact sensitive text while preserving layout.
- `privacy_scrub.py`: Remove personal metadata and revision IDs.
- `a11y_audit.py`: Audit accessibility and optionally fix simple issues such as table headers or image alt text.
- `style_lint.py`: Report formatting/style inconsistencies.
- `style_normalize.py`: Clean up inconsistent run or paragraph formatting.
- `captions_and_crossrefs.py`: Add captions and cross-reference support.
- `insert_ref_fields.py`: Insert Word reference fields.
- `flatten_ref_fields.py`: Flatten reference fields to stable visible text.
- `insert_toc.py`: Insert table-of-contents structures.
- `internal_nav.py`: Add internal navigation links.
- `xlsx_to_docx_table.py`: Convert spreadsheet ranges into Word tables.
- `docx_table_to_csv.py`: Extract Word tables to CSV.
- `watermark_add.py`: Add a watermark.
- `watermark_audit_remove.py`: Audit or remove watermarks.
- `set_protection.py`: Apply editing restrictions.
- `content_controls.py`: Work with Word content controls/forms.
- `render_and_diff.py`: Render and compare two DOCX files page by page.

### Default workflow

1. Create or edit the DOCX using `python-docx` for normal document structure.
2. Use OOXML-level patching only when needed for advanced features such as tracked changes, comments, fields, or relationships.
3. Save the final DOCX with its core QA report in the project workspace.
4. When the optional renderer is available or the document is high-stakes,
   render the DOCX to page PNGs and inspect every page.
5. Fix layout or content defects and rerun the checks.
6. Deliver only the final requested `.docx` file unless the user explicitly
   asks for intermediate PNGs or PDFs.

### Quality rules

The document must be visually clean before delivery. Tables should not break awkwardly, text should not clip, headers and footers should render correctly, fonts should display correctly, and any comments/tracked changes should be in the requested final state.

---

## 4. Practical Differences

| Artifact type | Best for | Main output | Key verification step |
|---|---|---|---|
| PPTX | Presentations, pitch decks, visual demos, posters | `.pptx` | Check slide layout and visual coherence |
| Excel | Trackers, models, dashboards, calculations, tables | `.xlsx` | Inspect formulas, scan errors, verify layout |
| DOCX | Written reports, memos, interview guides, redlines, comments | `.docx` | Core QA; rendered inspection when available/high-stakes |

---

## 5. Common Deliverable Pattern

For all real artifact tasks in this repository, the final file is written to
`projects/<project_name>/outputs/` and its QA evidence is written to `qa/`.
Intermediate scaffolds are not final deliverables.
