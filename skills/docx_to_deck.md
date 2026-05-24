# Workflow Skill: Word Source To Refined Deck

## When to use
Use when the user provides a Word document containing thinking, talking points,
tables, results, or a rough paper and asks for a PowerPoint deck. The Word file
is source material; the output is a concise presentation, not a document dump.

## Required workflow
1. Run the intake workflow; it copies the DOCX into the project, saves
   extracted context, creates a checked internal scaffold PPTX, and records the
   intake run in `project.json`:

```bash
python deck.py model_results.docx --project "Model Results" --no-open
```

The intake must retain all captured source paragraphs and table cells, along
with detected header/footer and text-box text. Long source material may span
continuation scaffold slides; never trim it merely to fit a single preview
slide. The original DOCX remains in `inputs/` as the authoritative source.

2. Before writing any refined code, read `skills/PATTERN_CATALOG.md` and
   produce a slide plan from the extracted context in `working/`. If
   `working/user_instructions.txt` exists, read it first and apply the editorial
   direction throughout. Record each slide in `working/slide_plan.md` with:
   `source evidence | assertion | selected pattern | why selected | inputs/assets | AI inference?`
   Sections marked `[DISTIL]` in the markdown must be rewritten — never copy
   source text verbatim onto a slide.
3. Use `DecisionLog.record_slide_plan(...)` so both `working/slide_plan.md`
   and the decision log contain the selection reasoning. Select patterns from
   the catalog, write the deck build script, and keep only content that
   advances the story.
4. Save the refined visual deck with final QA and register its output:

```python
deck_path = project.output_path("model_results_deck.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
project.register_output(
    deck_path, project.qa_path(deck_path.name), patterns=[...],
    qa_report_paths=[rendered_qa_path],
)
```

5. Read the refined QA report after every build. Correct errors and any
   actionable warnings about density, contrast, crowding, overlap or repeated
   panel symmetry, regenerate and rerun QA. Run rendered-review status on the
   final refined PPTX; inspect rendered slide PNGs if its optional runtime is
   present. Deliver only the corrected refined PPTX, never the scaffold.

For lower-level or custom scripted deck generation, run:

```bash
python build_deck_from_docx.py model_results.docx --project "Model Results"
```

This command also produces a scaffold only. Refinement is required before it
can become a user-facing deck.

## Content-shaping rules
1. Use Word tables as data sources; do not paste wide source tables blindly onto slides.
2. Turn quantitative comparison tables into charts or research patterns when the relationship matters more than exact transcription.
3. Preserve important qualifications, sample dates, units and source notes.
4. Treat working notes as raw input: rewrite them into short, audience-ready slide language.
5. Never use `content_slide` for narrative sections. Map every section to the best-fit pattern from `skills/`:
   - A key finding or recommendation → `assertion_evidence_slide`
   - A methodology or process → `process_slide`
   - A roadmap or schedule → `timeline_slide`
   - Headline statistics → `numbers_slide`
   - Workstream or programme status → `status_slide`
   - A numeric comparison matrix → `heat_map_slide` or `results_slide`
   - A paired metric comparison → `diverging_bar_slide`
   - A prioritisation or quadrant view → `two_by_two_slide`
   - Only fall back to `content_slide` if no pattern fits AND the section cannot reasonably be distilled into any of the above.
6. The catalog is exhaustive: consider patterns that are plausible for the
   evidence, then record why the selected one is best. You need not use every
   pattern, but you must not ignore suitable options.
