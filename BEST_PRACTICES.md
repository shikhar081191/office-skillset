# Best Practices For Using The Office Toolkit

## Who This Is For

This guide is for a user working through GitHub Copilot Chat or Windsurf who
wants a polished PowerPoint, Word document or Excel workbook without designing
the artifact manually or writing Python.

You provide the thinking, evidence and intended audience. The AI assistant
chooses the format and visual structure, builds the deliverable, checks it and
returns the finished file.

## Before You Ask The AI

Put together the material you already have. It does not need to be tidy.

Useful inputs include:

- A Word file containing talking points, rough narrative, tables or draft analysis.
- An Excel file or CSV with results that should become tables or charts.
- Text pasted into chat if the material is short.
- Existing charts or images that should appear in the output.
- A template or brand colour direction, if one must be followed.

The AI gives better results when you also state:

| Tell the AI | Example |
|---|---|
| What to create | `Build a PowerPoint deck` |
| Who will read it | `Credit Risk Committee` |
| What decision or outcome is needed | `Approve Model v3 for production` |
| What should stand out | `Lead with the PMSE improvement in treasury bills` |
| Desired style | `Use BlackRock styling` or `Use client navy and teal` |
| Any constraints | `Keep it to 10-12 slides` or `Include an appendix table` |

You do not need to specify individual slide designs. The assistant should
select suitable patterns such as charts, heat maps, timelines, scorecards or
recommendation slides based on your evidence.

## Starting From a Brief (No Word File)

If you don't have a Word document, use `BRIEF_TEMPLATE.md`. Copy the blank
brief, fill it in (takes ~2 minutes), then paste it into Copilot or Windsurf
chat along with your raw notes. The AI reads `AI_INSTRUCTIONS.md` and
`STORY_TEMPLATES.md`, picks the right story template, and builds the deck
directly — no intake step needed.

```text
[paste your filled BRIEF_TEMPLATE.md here]

NOTES:
Model v3 improved PMSE by 18.2% on holdout data. Bills drove 46pp of the gain.
Validated across Rates, Credit, Equities. FX excluded — separate Q4 workstream.
Legal sign-off on data lineage still in progress, target 30 Jun...
```

The AI will:
1. Detect the deck purpose and pick a story template (e.g. Approval Request)
2. Map your notes to the right slide slots
3. Build a 5–8 slide deck using the appropriate visual patterns
4. Run QA and fix any issues before delivery

This is the fastest path when your material is still rough thinking rather than
a structured document.

---

## A Good Prompt

### Recommended Word-To-PowerPoint Prompt

```text
Build a polished PowerPoint deck from this Word file:

projects/my_project/inputs/model_review_brief.docx

Project name: Model Review Committee Pack
Audience: Credit Risk Committee
Decision required: Approve Model v3 for production deployment
Style: Use the default BlackRock palette
Emphasis: Lead with the most important PMSE improvement, show model comparison
clearly, and finish with recommendation, implementation risks and next steps.

Please follow the repository workflow:
- Treat the Word file as the source, not as slide-ready text.
- Run the Word intake step first and keep the scaffold internal.
- Read the available slide-pattern catalog and decide the best visual structure.
- Create a slide plan before building the final deck.
- Build a refined visual presentation, not a copied-text deck.
- Run the required QA checks, fix real issues, and provide only the final
  refined deck with its QA evidence.
```

### Why This Prompt Works

It gives the assistant five things it cannot reliably guess:

1. The source file to use.
2. The audience.
3. The decision the presentation should support.
4. The visual/style direction.
5. The evidence or message that should dominate the story.

The rest, including pattern selection and slide layout, should be handled by
the assistant.

## Other Useful Prompt Examples

### From Notes Pasted In Chat

```text
Create an 8-slide PowerPoint for senior management from the notes below.
The decision required is whether to fund the next phase of the model build.
Use BlackRock styling. Use charts or visual patterns where the numbers support
them, keep detail concise, and run QA before giving me the final deck.

[Paste notes and figures here]
```

### From Free-Form Notes — Full Worked Example

Use this pattern when your source material is unstructured thinking you have
typed or dictated, with no Word file and no pre-formatted bullets. The AI
should select the right story template from `STORY_TEMPLATES.md`, simplify
language for the stated audience, design visual layouts around the structure,
and produce a clean deck — not slide-ified paragraphs.

The prompt below is a complete, ready-to-use example. Replace the bracketed
notes block with your own content before sending.

```text
Build a polished PowerPoint from the free-form notes below.

Project name: AI MDS — Automated Market Scenario Analysis
Audience: Junior analyst with no prior background in quantitative risk or
portfolio modelling. Assume no knowledge of internal systems, terminology
or team structure.
Purpose: Explain what the AI MDS project is, how the pipeline works, and
why it matters — both for speed and for governance.
Style: Use the prm palette (navy/teal). Plain language throughout — define
any technical concept the first time it appears, no unexplained acronyms.
Max slides: 8
Emphasis: The five-step pipeline is the core of the story. The human-in-
the-loop safeguard is the second priority. The governance and audit angle
is the closing message and should land as the deeper value, not an afterthought.

Please follow the AI_INSTRUCTIONS.md workflow:
- Read STORY_TEMPLATES.md and select the appropriate template for this content.
- Create a slide plan before building — record which pattern you chose for
  each slide and why.
- Build a refined visual presentation using process flows, numbered steps,
  cards and contrast panels where the content supports them. Do not copy
  paragraphs onto slides.
- Run QA, fix any real issues, and deliver only the final refined deck
  with its QA evidence.
- Save all outputs to a new project folder under projects/.

NOTES:
[Paste your free-form notes here — unedited, in whatever order they came to you.
The AI will organise, simplify and structure them. You do not need to clean them up.]
```

#### Why this prompt works

| Element | Why it matters |
|---|---|
| Audience line | Tells the AI what language level to target — this changes word choice, what gets defined, how much context each slide needs |
| Purpose line | Separates "explain what it is" from "approve this" — the AI picks a Project Explainer template rather than an Approval Request |
| Emphasis line | The AI has 8 slots and more than 8 topics — you decide what stays in and what lands at the end |
| Workflow instruction | Ensures a slide plan is written before the deck is built; without this Copilot may skip straight to output |
| Notes block | Paste raw text exactly as it is — the AI's job is to extract structure, not yours |

#### What the AI will do with this

Following `STORY_TEMPLATES.md`, the AI will identify this as a **Project Explainer**
and apply that template's opening principle: *lead with the problem, not the solution*.
A likely slide sequence for the notes above:

| # | Slide | Pattern |
|---|---|---|
| 1 | Title | `title_slide` |
| 2 | The problem: scenario analysis today is slow and manual | `two_column_contrast_slide` (navy bar — problem framing) |
| 3 | What we built: plain-English to portfolio P&L in minutes | `assertion_evidence_slide` |
| 4 | How it works: the five-step pipeline | `numbered_steps_slide` (navy bar) |
| 5 | Human in the loop: every step is reviewable and auditable | `three_column_card_slide` (teal bar) |
| 6 | Who is involved: three teams, one two-week cadence | `three_column_card_slide` (navy bar) |
| 7 | Where we are: prototype live, beta testing mid-year, production H2 2026 | `timeline_slide` |
| 8 | Why it matters: speed is the headline, governance is the real value | `callout_bar_slide` (teal bar) |

The AI may deviate from this sequence based on what fits the evidence. The
`working/slide_plan.md` it creates will explain every choice.

### From A Word File And Excel Results

```text
Create a presentation for the Portfolio Risk Committee using:
- Narrative brief: projects/liquidity_review/inputs/liquidity_brief.docx
- Results data: projects/liquidity_review/inputs/liquidity_results.xlsx

The main message is that liquidity risk remains controlled except in the
short-dated stress cohort. Use the Word file for context and the Excel data
for charts/tables. Build, check and deliver the refined deck.
```

### For A Word Document Rather Than Slides

```text
Create a short committee note in Word from this source:
projects/model_signoff/inputs/model_signoff_notes.docx

Use a formal whitepaper style with executive summary, recommendation,
supporting evidence, limitations and next steps. Preserve all important
numbers and sources. Run document QA before delivery.
```

### For An Excel Workbook

```text
Create an Excel workbook from the supplied PMSE results with:
- an inputs/results sheet,
- a model-comparison summary,
- a chart-ready dashboard tab,
- clear assumptions and formula conventions.

Use the repository workbook QA checks before delivering the final workbook.
```

## What You Need To Do

For a new piece of work:

1. Create or use a project folder under `projects/`.
2. Put your source files in its `inputs/` folder, or tell the assistant where
   your source material is located.
3. Open Copilot Chat or Windsurf and provide a prompt like the examples above.
4. Review the final artifact and ask for revisions in ordinary language.

Examples of useful follow-up feedback:

```text
Make the recommendation slide more decisive.
Replace the performance table with a heat map and keep exact values in appendix.
Reduce the deck to 9 slides without losing the evidence.
Use our client palette instead of BlackRock styling.
Make the risk slide more suitable for an executive audience.
```

The assistant should rebuild the artifact and rerun QA after each material
revision.

## What Happens Behind The Scenes

For a Word-to-PowerPoint request, the workflow is:

```text
+------------------------------------------------------+
| 1. You provide source files and desired outcome      |
+------------------------------------------------------+
                         |
                         v
+------------------------------------------------------+
| 2. Source content is copied and extracted            |
|    Text and tables are preserved for the AI          |
+------------------------------------------------------+
                         |
                         v
+------------------------------------------------------+
| 3. An internal scaffold is created                   |
|    This is working material, not the final deck      |
+------------------------------------------------------+
                         |
                         v
+------------------------------------------------------+
| 4. The AI chooses the story and visual patterns      |
|    A slide plan records what will be built and why   |
+------------------------------------------------------+
                         |
                         v
+------------------------------------------------------+
| 5. The AI builds the refined visual presentation     |
+------------------------------------------------------+
                         |
                         v
+------------------------------------------------------+
| 6. QA checks layout, text fit, overlap, contrast,    |
|    density and repeated-panel alignment              |
+------------------------------------------------------+
                         |
                         v
                  +---------------+
                  | Issues found? |
                  +---------------+
                    | Yes       | No
                    v           v
       +--------------------+   +--------------------------------------+
       | Fix and rebuild    |   | Record rendered-review status and     |
       | then return to QA  |   | inspect previews when available       |
       +--------------------+   +--------------------------------------+
                    |                           |
                    +-----> Step 6              v
                                    +----------------------------------+
                                    | Final refined deck and QA reports|
                                    +----------------------------------+
```

In plain language:

```text
your material
  -> captured source content
  -> internal working draft
  -> AI-designed slide plan
  -> polished deck
  -> quality checks and fixes
  -> final PowerPoint
```

The internal scaffold is not your finished presentation. It exists so the AI
can preserve and understand the source before turning it into a concise visual
story.

## Files You Should Expect

For a presentation project:

```text
projects/<project_name>/
  inputs/    Your source Word/Excel/data files
  working/   Extracted context, slide plan and decision log
  assets/    Charts or images used in the deck
  scripts/   Build script created for the project, when needed
  outputs/   The final refined presentation
  qa/        Quality reports
```

Most useful files:

| File | Why it matters |
|---|---|
| `outputs/<name>_refined.pptx` | Final deck to review or share |
| `working/slide_plan.md` | Which layouts the AI chose and why |
| `working/decision_log.md` | How the story was structured |
| `qa/<name>_refined.qa.json` | Structural QA results |
| `qa/<name>_refined.rendered_qa.json` | Whether rendered visual review ran or why it could not |

## Simple Rules For Better Results

- Give the audience and required decision, not just a topic.
- Include real data and qualifications where possible.
- Tell the AI which conclusion matters most.
- Provide source files instead of retyping large tables.
- Ask for the final refined deliverable, not merely a conversion.
- Treat the first output as revisable: ordinary-language feedback is expected.
- Keep sensitive or unverified information identified clearly.

## What A Good Final Delivery Looks Like

A finished presentation should have:

- A clear story suited to the audience and requested decision.
- Visual treatment of numbers rather than copied paragraphs.
- Appropriate charts, comparison tables, timelines or decision slides.
- A recorded slide plan and final QA evidence.
- No unresolved actionable layout or legibility issues.
- A final refined file in `outputs/`, not the internal scaffold.
