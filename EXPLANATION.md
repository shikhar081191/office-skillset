# How This Toolkit Works

## For The User

You do not need to design slides or write Python.

For ready-to-use prompt examples and practical guidance, see
[`BEST_PRACTICES.md`](BEST_PRACTICES.md).
For what is already built and what comes next, see
[`REMAINING_PLAN.md`](REMAINING_PLAN.md).

Give the AI assistant:

- a Word file with your thinking, results and tables; or
- text pasted into chat; or
- an Excel/data file plus a short explanation of the story you want to tell.

Then say what you need, for example:

```text
Turn this Word brief into a presentation for the Risk Committee.
The decision needed is approval of Model v3.
Use BlackRock styling, highlight the PMSE improvement, and give me the final deck.
```

The assistant should handle the rest.

## What Happens Behind The Scenes

For a PowerPoint built from a Word file:

```text
1. Your Word file is copied into a project folder.
2. The system extracts headings, narrative and tables.
3. It creates an internal scaffold to understand the content; long text and
   large tables are carried across continuation slides rather than discarded.
4. The AI reviews the full pattern catalog and writes a slide plan explaining
   the storyline and selected visual pattern for each slide.
5. It builds a refined presentation with charts, tables, timelines, scorecards,
   heat maps or other suitable slide layouts.
6. It runs quality checks for layout, text fit, overlap, contrast, density and repeated-panel alignment.
7. It records a rendered slide review step; if the optional renderer is installed, it creates images for visual inspection.
8. It fixes real issues found by those checks and rebuilds the deck.
9. You receive the refined PowerPoint, not the internal scaffold.
```

The scaffold is working material. If it looks like copied Word text, it is not
the final output and should not be delivered. Its extracted working files keep
the captured paragraphs, tables, headers, footers and detected text-box text
available for the AI refinement step, while the original Word file is retained
in `inputs/`.

## Where Files Go

Each piece of work gets its own folder:

```text
projects/<your_project>/
  inputs/    Your supplied Word, Excel or data files
  working/   Extracted content and the AI decision log
  assets/    Charts and images used in the output
  scripts/   Project-specific build logic, when needed
  outputs/   Generated files, including the final refined deliverable
  qa/        Quality-check reports
```

The useful files for you are normally:

- `outputs/<name>_refined.pptx` for the final presentation
- `working/slide_plan.md` to see which visual patterns were chosen and why
- `working/decision_log.md` if you want to see how the AI structured the story
- `qa/<name>_refined.qa.json` if you want evidence that checks were run
- `qa/<name>_refined.rendered_qa.json` to see whether rendered visual review ran or why it was unavailable

## How Feedback Works

You can ask for revisions in ordinary language:

```text
Make the recommendation more prominent.
Use fewer slides.
Replace the table with a chart.
Change to our client palette.
Show more detail on model validation.
```

The assistant should revise the deck, rerun its checks and return an updated
refined file.

## PowerPoint, Word And Excel

### PowerPoint

The most developed workflow. It converts raw thinking and data into a concise
story using reusable research-oriented slide patterns, followed by QA and
repair before delivery.

### Word Documents

The toolkit can create structured research notes, memos and template-driven
documents. Final Word files should pass document QA; rendered page inspection
can be added when the optional rendering setup is available.

### Excel Workbooks

The toolkit can create research workbooks and analysis tables with formatting
and QA checks. Over time, more reusable workbook patterns can be added in the
same way slide patterns are added.

## Styling And Brand Colours

BlackRock styling is the default. A different named palette or externally
specified brand palette can be used without rewriting the slide logic.

Example user instruction:

```text
Use our client palette: navy primary, teal secondary, and gold accent.
```

## What Good Delivery Means

A presentation is ready only when:

- it is the refined visual deck, not the text scaffold;
- the AI has recorded a pattern-based slide plan before building;
- the main message is clear on each slide;
- the QA report has been interpreted and real issues fixed;
- the final file is saved in the project `outputs/` folder.

Rendered slide-image inspection is an additional quality layer when its
optional runtime is installed.
