# DOCX Python utilities

This folder contains reusable Python utilities for DOCX creation/editing QA workflows.

Core entry points:

- `render_docx.py` — render a DOCX to page PNGs, optionally also PDF, for visual QA.
- `scripts/docx_ooxml_patch.py` — low-level OOXML helper used by several scripts.
- `scripts/render_and_diff.py` — render two DOCXs and produce page-level image diffs.
- `scripts/comments_*.py` — add, extract, patch, or strip comments.
- `scripts/accept_tracked_changes.py`, `scripts/add_tracked_replacements.py` — manage redlines/tracked changes.
- `scripts/a11y_audit.py` — accessibility checks and simple fixes.
- `scripts/privacy_scrub.py`, `scripts/redact_docx.py` — metadata/privacy and redaction helpers.
- `scripts/style_lint.py`, `scripts/style_normalize.py` — style diagnostics and conservative cleanup.
- `scripts/xlsx_to_docx_table.py`, `scripts/docx_table_to_csv.py` — table conversion helpers.

Typical visual QA command:

```bash
python docx/render_docx.py input.docx --output_dir out --emit_pdf
```
