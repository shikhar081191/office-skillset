# Mandatory Slide Planning Catalog

## Purpose

This catalog is the decision surface for an AI assistant creating or refining
a presentation. Read it before writing a final deck script.

The assistant is the reasoning layer. Python builders do not decide the story:
they extract source content, draw selected patterns and run QA.

## Required Agent Decision Loop

For every user-facing deck:

1. Read the user's brief, audience, requested decision and palette direction.
2. For Word input, run `deck.py` first, then read the complete captured source
   under `working/` and treat the generated scaffold as internal only.
3. Consider the relevant alternatives in the full catalog below. Do not simply
   repeat a prior layout or convert each source heading into bullets.
4. Create `working/slide_plan.md` before authoring the refined deck. Every
   planned slide must record its source evidence, assertion, selected pattern,
   selection rationale, required inputs and whether AI inference was introduced.
5. Build the refined deck from that plan using the named patterns/utilities.
6. Save with final QA, read the QA report, fix errors and actionable warnings,
   rebuild, then run `ArtifactUtilities(project).review_rendered_pptx(deck_path)`.
   Inspect slide PNGs if available, or retain the explicit optional-runtime
   skip report, and deliver only the refined checked deck.

## Required Slide Plan Columns

```markdown
| # | Source evidence | Assertion | Selected pattern | Why selected | Inputs / assets | AI inference? |
|---|---|---|---|---|---|---|
| 1 | Executive Summary | Approve Model B | `recommendation_slide` | A decision is requested | Recommendation + evidence | No |
```

Use `DecisionLog.record_slide_plan(...)` in a refined build script; it records
the plan in both `working/slide_plan.md` and `working/decision_log.md`.

## Pattern Selection Catalog

| Pattern | Select when the slide's job is... | Required source shape | Avoid when... |
|---|---|---|---|
| `title_slide` | Introduce topic, audience and period | Title and subtitle | The slide needs evidence or detailed content |
| `section_divider` | Mark a major narrative transition | Short section name | A content slide is needed |
| `agenda_slide` | Explain the structure of the meeting/deck | 3-6 section labels | The content is a status list or timeline |
| `recommendation_slide` | Request a decision or land a proposed action | One decision, evidence, optional caveats | No clear recommendation exists |
| `assertion_evidence_slide` | Make one finding memorable and prove it | Short assertion and 3-4 supporting facts | Several options need comparison |
| `numbers_slide` | Surface 3-4 headline metrics | Large values and short qualifiers | Relationships/trends matter more than totals |
| `kpi_dashboard_slide` | Summarise operating/model health metrics | 3-6 metrics with optional deltas | The slide should argue one focused conclusion |
| `results_slide` | Compare exact values across alternatives/cohorts | Compact table with important exact values | A large matrix or trend is easier to read visually |
| `heat_map_slide` | Show intensity/pattern across a numeric matrix | Rows x columns of comparable numeric values | Exact reading of every number is critical |
| `diverging_bar_slide` | Compare two positive measures across categories | Paired values and category labels | Contributions add to a total |
| `waterfall_slide` | Explain contributors to a net change/total | Positive/negative component list | Values are independent comparisons |
| `scorecard_slide` | Support selection using criteria and weighted judgement | Options, criteria, comparable scores | The decision is already singular and proven |
| `bar_chart_slide` | Compare categories or series at one point/period | Categories and one/few numeric series | Values are temporal trend points |
| `line_chart_slide` | Show change over time | Ordered dates/periods and numeric series | Categories are unordered |
| `chart_context_slide` | Pair an existing chart with narrative implication | Chart image, headline, support, so-what | Native chart generation is preferable and available |
| `annotation_chart_slide` | Call out events or inflections on an existing chart | Chart image and positioned annotations | No chart or no specific events to explain |
| `process_slide` | Explain how a method/workflow works | 3-5 sequential steps | Dates/milestones drive the story |
| `timeline_slide` | Show dated delivery or event sequence | Milestones with state/date | The sequence is methodological rather than dated |
| `status_slide` | Show workstream health and ownership | Workstreams, RAG and owners | The objective is a scored selection |
| `two_by_two_slide` | Prioritise items along two judgement axes | Items with x/y placement | Exact numeric comparison is needed |
| `assumption_table_slide` | Document model/scenario assumptions and sensitivity | Assumptions, values, sources, sensitivity | The story is result performance rather than inputs |
| `content_slide` | Hold unavoidable residual prose only | Short material that fits no visual pattern | As a default transformation of Word narrative |

## Utility Selection Cues

Patterns are not the only available tools. While planning, consider:

| Need | Utility or workflow |
|---|---|
| Preserve text and tables from Word input | `deck.py`, `source_docx.py` and the copied DOCX in `inputs/` |
| Reuse an external chart image | `chart_context_slide` or `annotation_chart_slide` |
| Create/edit Word reports, templates or redlines | DOCX builders and `artifact_utilities.py` DOCX utilities |
| Create research workbooks or PMSE comparison files | `XlsxBuilder`, `xlsx_patterns.py`, `xlsx_qa.py` |
| Add captions, comments, redaction or accessibility review | `artifact_utilities.py` |
| Record rendered slide review and create previews when available | `ArtifactUtilities(project).review_rendered_pptx(deck_path)` |
| Check slide legibility, bounds and repeated panel alignment | Mandatory `pptx_qa.py` save/report loop |

## Final Planning Check

Before writing the refined deck script, confirm:

- Every user-facing slide has one assertion and an appropriate chosen pattern.
- All important source evidence has a destination or an intentional appendix decision.
- No slide uses `content_slide` merely because the source contained paragraphs.
- Required data, units, dates, qualifications and sources are preserved.
- Any inference not directly supported by the source is identified.
