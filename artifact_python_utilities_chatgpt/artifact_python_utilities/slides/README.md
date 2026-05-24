# Slides Python utilities

This folder contains Python utilities and Python examples for presentation workflows.

Container tools:

- `container_tools/render_slides.py` — render slides to images for visual QA.
- `container_tools/create_montage.py` — build a montage/contact sheet from rendered slides.
- `container_tools/detect_font.py` — inspect font use/availability.
- `container_tools/ensure_raster_image.py` — normalize image assets before inserting into decks.
- `container_tools/slides_test.py` — smoke-test style helper.

Artifact-tool examples:

- `artifact_tool_examples/presentation_quick_start.py`
- `artifact_tool_examples/slide_quick_start.py`
- `artifact_tool_examples/integrated_example.py`
- plus focused examples for auto-layout, editing, images, and overlap warnings.

Typical QA commands depend on the chosen deck generation library, but the usual pattern is:

1. Create or edit the presentation.
2. Render slides to PNGs.
3. Inspect for clipping, overlaps, missing images, and layout drift.
4. Export the final `.pptx`.
