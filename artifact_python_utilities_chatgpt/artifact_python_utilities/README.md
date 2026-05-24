# Artifact Python utilities bundle

This bundle contains the Python-side utilities corresponding to the DOCX, PPTX/slides, and spreadsheet artifact workflows.

## Contents

- `docx/` — DOCX render, OOXML patch, comments, tracked changes, redaction, style, accessibility, and table utilities.
- `slides/` — slide rendering/montage/image/font utilities plus Python examples for the presentation artifact tool.
- `spreadsheets/` — artifact_tool quick-start reference and a small Python starter scaffold.

## Notes

- These are utility scripts/scaffolds, not a full Python package with pinned dependencies.
- DOCX and slides workflows generally require rendering for visual QA before delivery.
- Spreadsheet work is expected to use `artifact_tool` APIs for workbook creation/editing/export.
- File paths in examples often assume `/mnt/data`; adjust paths for your environment.
