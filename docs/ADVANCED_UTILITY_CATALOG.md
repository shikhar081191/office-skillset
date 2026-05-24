# Advanced Utility Catalog

All executable Python utilities preserved in
`artifact_python_utilities_chatgpt/artifact_python_utilities/` are exposed
through `artifact_utilities.py`. The original sources remain intact.

List current capabilities and runtime availability:

```bash
python artifact_utilities.py list
python artifact_utilities.py list --json
```

Run any retained utility through the supported facade:

```bash
python artifact_utilities.py --project "Model Review" run docx.a11y_audit -- projects/model_review/outputs/note.docx
```

When a project is specified, execution logs are written into its `qa/` folder.
For presentation delivery, these utilities supplement the core workflow:
the AI first creates the refined patterned deck, interprets and repairs its
`pptx_qa.py` findings, and then always calls
`ArtifactUtilities(project).review_rendered_pptx(deck_path)`. When the
rendering runtime is available it produces previews and rendered overflow
testing; when it is not, it writes an explicit skipped-runtime report.

## DOCX Utilities

The DOCX utility family is available through IDs:

```text
docx.render_docx
docx.accept_tracked_changes
docx.add_tracked_replacements
docx.a11y_audit
docx.apply_template_styles
docx.captions_and_crossrefs
docx.comments_add
docx.comments_apply_patch
docx.comments_extract
docx.comments_strip
docx.content_controls
docx.docx_ooxml_patch
docx.docx_table_to_csv
docx.fields_materialize
docx.fields_report
docx.flatten_ref_fields
docx.footnotes_report
docx.heading_audit
docx.images_audit
docx.insert_note
docx.insert_ref_fields
docx.insert_toc
docx.internal_nav
docx.make_fixtures
docx.merge_docx_append
docx.privacy_scrub
docx.redact_docx
docx.render_and_diff
docx.section_audit
docx.set_protection
docx.style_lint
docx.style_normalize
docx.watermark_add
docx.watermark_audit_remove
docx.xlsx_to_docx_table
```

Core OOXML/editing utilities use the normal project environment plus `lxml`.
`docx.render_docx` and `docx.render_and_diff` additionally require the optional
rendering runtime described below.

Convenience Python API examples:

```python
from artifact_utilities import ArtifactUtilities
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Review")
tools = ArtifactUtilities(project)
csv_path = tools.export_docx_table(project.input_path("brief.docx"), "table_1.csv")
report_docx = tools.xlsx_table_to_docx(project.input_path("results.xlsx"))
```

## Slides Utilities

Supported retained IDs:

```text
slides.render_slides
slides.create_montage
slides.detect_font
slides.ensure_raster_image
slides.slides_test
```

The following artifact-tool examples are also cataloged and callable when that
optional presentation runtime is installed:

```text
slides.examples.auto_layout_bottom_right_alignment
slides.examples.auto_layout_center_card
slides.examples.auto_layout_header_footer_layout
slides.examples.auto_layout_horizontal_layout
slides.examples.auto_layout_static_engine_api
slides.examples.images_quick_start_local_path
slides.examples.integrated_example
slides.examples.presentation_edit
slides.examples.presentation_quick_start
slides.examples.slide_move_to
slides.examples.slide_quick_start
slides.examples.warn_about_overlaps
```

`slides.create_montage` works on existing PNG/JPG slide images in the core
environment. Final refined decks always record rendered-review status;
actually rendering slide PNGs and running rendered overflow checks requires the
optional rendering runtime.

## Spreadsheet Utilities

The retained artifact-tool workbook scaffold is exposed as:

```text
spreadsheets.spreadsheet_artifact_tool_starter
```

The core workbook workflow continues to use `XlsxBuilder` and `xlsx_qa.py`.
The retained scaffold becomes executable once an `artifact_tool` runtime is
standardized for the team.

## Optional Rendering Runtime

The richer visual QA paths need additional dependencies:

```bash
pip install -r requirements-optional-rendering.txt
```

They also require system binaries:

- LibreOffice / `soffice` for Office-to-PDF conversion.
- Poppler tools such as `pdfinfo` and `pdftoppm`.
- Fontconfig for the slide font-detection path.
- External image converters only when rasterizing vector or specialist formats.

Until these are installed, core PPTX, DOCX and XLSX structural QA remains
automatic and operational; final PPTX workflows record that rendered review
was skipped because advanced render tools are unavailable.
Structural QA does not make a Word-source scaffold a final presentation; the
refinement and warning-repair stages still apply.
