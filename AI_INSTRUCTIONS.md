# AI Instructions: Building Decks with ms_office_utils

You are building `.pptx` presentation decks using Python.
Follow these instructions precisely — they tell you *how* to work, not just what to produce.

---

## Step 0: Always Read These Files First

Before writing any code, read:

1. `skills/PATTERN_CATALOG.md` — when to use each of the 19 patterns
2. `skills/` folder — individual pattern files for any pattern you plan to use

Do not skip this. The pattern catalog contains selection criteria that are not in
your training data. Reading it before planning prevents the wrong pattern choice.

---

## Step 1: Parse the Brief

If the user has filled in `BRIEF_TEMPLATE.md`, read those fields directly —
they map exactly to the fields below. Otherwise, extract from free-form text.

When given a deck brief, extract these before writing a single line of code:

**Required:**
- What is the deck's **purpose**? (approval request, exec update, project explainer, research findings, status update)
- Who is the **audience**? (junior analyst, senior internal, executive/sponsor, external client, technical peer)
- What **topics** must be covered?
- Is there a **key decision or message** the deck must land?

**Infer if not stated:**
- Audience expertise → determines how much to explain vs. assume
- Tone → formal for exec/external, direct for internal
- Slide count → use the story template default unless told otherwise

**If the brief is ambiguous on a critical point**, make a reasonable assumption,
state it in a comment at the top of your script, and proceed. Do not ask for
clarification before attempting — produce a draft and flag your assumptions.

---

## Step 2: Plan the Slide Sequence

Write the planned slide list as a comment before any code. Choose a story template
from the table below, then map the brief content to the slots.

For per-slot content rules and fill-in-the-blanks code scaffolds, read
`STORY_TEMPLATES.md` — it has a complete section for each template.

### Story Templates

**Approval Request** (5–6 slides)
```
1. Recommendation       → recommendation_slide
2. Supporting Evidence  → assertion_evidence_slide or numbers_slide
3. Data / Analysis      → heat_map_slide, results_slide, or bar_chart_slide
4. Risks & Caveats      → status_slide or content_slide
5. Timeline / Next Steps → timeline_slide
```

**Project Explainer** (8 slides — default for non-specialist audiences)
```
1. Title                → PptxBuilder.title_slide
2. The Problem          → two_column_contrast_slide (left: today's pain, right: cost/risk; navy bar)
3. The Solution         → assertion_evidence_slide (one headline claim + 3–4 proof points)
4. How It Works         → numbered_steps_slide (3–5 sequential steps; navy bar)
5. Governance / Controls → three_column_card_slide (3 control pillars; teal bar)  — OR status_slide
6. Team / Ownership     → three_column_card_slide (one card per team; navy bar)
7. Timeline / Where We Are → timeline_slide
8. Why It Matters       → callout_bar_slide (body bullets + single closing statement; teal bar)
```

**Exec Update** (5–6 slides)
```
1. Headline KPIs        → kpi_dashboard_slide
2. Status by Workstream → status_slide
3. Key Issue / Decision → recommendation_slide
4. Supporting Data      → results_slide or chart_context_slide
5. Timeline             → timeline_slide
6. Ask / Next Steps     → content_slide (dark_bg=True)
```

**Research Findings** (6–7 slides)
```
1. Core Finding         → assertion_evidence_slide
2. Method Summary       → process_slide
3. Key Result           → numbers_slide or kpi_dashboard_slide
4. Full Data            → heat_map_slide or results_slide
5. Sensitivity / Risks  → assumption_table_slide or scorecard_slide
6. Implications         → waterfall_slide or diverging_bar_slide
7. Recommendation       → recommendation_slide
```

**Status Update** (4–5 slides)
```
1. Summary KPIs         → kpi_dashboard_slide
2. Workstream Status    → status_slide
3. Key Risks            → two_by_two_slide
4. Timeline             → timeline_slide
5. Decision Needed      → recommendation_slide (optional)
```

### Pattern Diversity Rule

Apply these constraints before finalising the slide plan — they prevent the two
most common Copilot failure modes: card-heavy decks and slides with blank space.

**Variety:**
1. No single pattern may appear more than **twice** in one deck. If a pattern
   would appear three times, either combine two slides or substitute an
   alternative pattern with similar purpose.
2. No single **visual-bar** pattern (`three_column_card_slide`,
   `two_column_contrast_slide`, `numbered_steps_slide`, `callout_bar_slide`)
   may appear **more than once** per deck. Combined visual-bar patterns should
   not exceed half the content slides (title excluded). For an 8-slide deck
   (7 content slides) that means a maximum of 3–4 visual-bar slides.
3. A deck of 6 or more slides **must use at least 5 different pattern types**.
   (Title and section divider count as one type each.)

**Preference order for content with data:**
- Numbers → `numbers_slide`, `kpi_dashboard_slide`
- Comparisons → `results_slide`, `heat_map_slide`, `diverging_bar_slide`
- Trends → `bar_chart_slide`, `line_chart_slide`, `waterfall_slide`
- Sequential steps → `numbered_steps_slide` (visual-bar) or `process_slide`
- Cards only when content is genuinely 3 parallel equal-weight items with real
  body text — not when it is one argument with sub-points

**`callout_bar_slide` is a closing pattern.** Use it once per deck on the final
or penultimate slide. Do not use it as a generic bullet slide with a tagline.

---

## Step 3: Write the Build Script

### Imports

```python
from create_pptx import PptxBuilder
from patterns import (
    # Standard patterns (thin-rule header)
    status_slide, process_slide, timeline_slide, results_slide,
    chart_context_slide, numbers_slide, heat_map_slide, waterfall_slide,
    scorecard_slide, annotation_chart_slide, two_by_two_slide,
    assertion_evidence_slide, diverging_bar_slide, recommendation_slide,
    kpi_dashboard_slide, agenda_slide, assumption_table_slide,
    bar_chart_slide, line_chart_slide,
    # Visual-bar patterns (full-width header bar, prm-deck-kit style)
    three_column_card_slide, two_column_contrast_slide,
    numbered_steps_slide, callout_bar_slide,
)
```

### Builder Setup

```python
b = PptxBuilder(palette="blackrock")  # or: midnight_executive, teal_trust, charcoal_minimal
```

### Available Palettes

| Name                 | Feel                                                    |
|----------------------|---------------------------------------------------------|
| `prm`                | Navy + teal + amber — financial, polished (recommended) |
| `blackrock`          | Black primary, orange accent                            |
| `midnight_executive` | Navy + ice blue — premium/exec                          |
| `teal_trust`         | Teal — analytical/financial                             |
| `charcoal_minimal`   | Charcoal + white — clean/modern                         |
| `coral_energy`       | Coral + gold — energetic                                |
| `warm_terracotta`    | Terracotta + sage — approachable                        |

### Title Slide (always use PptxBuilder directly)

```python
b.title_slide(
    title="Deck Title Here",
    subtitle="Audience · Date · Author"
)
```

### Pattern Calls (chain onto the same builder `b`)

Every pattern function takes `builder=b` as its last argument and returns `b`.
This chains all slides into one file.

```python
from create_pptx import PptxBuilder, PALETTES
from patterns import recommendation_slide, numbers_slide
from project_workspace import ProjectWorkspace
from artifact_utilities import ArtifactUtilities

project = ProjectWorkspace("Model v3 Approval")
prm = PALETTES["prm"]  # use palette constants — never hardcode hex strings

b = PptxBuilder(palette="prm")
b.title_slide("Model v3 Approval", "Risk Committee · May 2026")

recommendation_slide(
    title="Recommendation",
    recommendation="Approve Model v3 for production deployment",
    rationale=[
        "PMSE improved 18% on holdout data",
        "Passes all regulatory stress tests",
        "Team validated on three asset cohorts",
    ],
    caveats=["Requires sign-off from Legal by 30 Jun"],
    builder=b,
)

numbers_slide(
    title="Performance at a Glance",
    stats=[
        {"number": "18%",   "label": "PMSE improvement",      "context": "vs. current production model"},
        {"number": "p<.01", "label": "Statistical significance", "context": "across all cohorts"},
        {"number": "3",     "label": "Asset cohorts validated", "context": "Rates, Credit, Equities"},
    ],
    builder=b,
)

# ── Save, QA, rendered review, output registration ────────────────────────────
deck_path = project.output_path("model_v3_approval.pptx")
qa_path   = project.qa_path(deck_path.name)
b.save(deck_path, final=True, report_path=qa_path)

rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)
project.register_output(
    deck_path,
    qa_report_path=qa_path,
    patterns=["recommendation_slide", "numbers_slide"],
    qa_report_paths=(rendered_qa_path,),
)
print(deck_path)
```

**Every brief-path build script must include all four of these closing calls:**
1. `b.save(deck_path, final=True, report_path=qa_path)` — runs QA
2. `ArtifactUtilities(project).review_rendered_pptx(deck_path)` — records rendered review
3. `project.register_output(...)` — logs the output in `project.json`
4. Read the QA report; fix errors and actionable warnings; rerun if needed

### Key Rules

1. **Always pass `builder=b`** to every pattern call — this puts all slides in one file
2. **Never pass more data than a pattern can render** — check the pattern file for limits
3. **Use `key_message` on every slide that has a clear single-sentence takeaway.** It renders as a bold banner below the title and fills what would otherwise be blank space at the top of the content area. If you can't write one sentence summarising the slide's point, the slide probably needs a different pattern.
4. **Bullet text must be short** — max 10 words per bullet; split long text across two slides
5. **Numeric data in tables/heat maps must be pre-processed** — pass floats, not formatted strings, to value arrays
6. **Do not create slides without data** — if you don't have content for a slot, drop that slide
7. **Fill all optional fields when content supports them.** In `three_column_card_slide`, write real body text in all cards and use the `tag` field. In `chart_context_slide`, populate `so_what`. In `numbered_steps_slide`, use `footnote` if there is a caveat. Blank optional fields are not neutral — they leave visible empty space.
8. **A slide with large empty areas means the wrong pattern was chosen.** If a slide looks sparse — few short bullets, mostly whitespace — switch to a richer data pattern or merge the content onto an adjacent slide. Do not pad with filler text.

### Text Quality Rules (apply to all text you write into slides)

- Short declarative sentences. One idea per sentence.
- Active voice: "The model reads your input" not "Your input is read by the model."
- No hedge-stacking: "This may potentially help to..." → "This helps..."
- No filler openings: not "In today's rapidly evolving landscape..." — start with the point.
- Bullet lists max 5 items — split into two groups or use a card pattern instead.
- Every number needs a unit and context: "18%" → "18% PMSE improvement vs. baseline"

---

## Step 4: QA and Fix

### Run the build script

```bash
python your_build_script.py
```

QA runs automatically on `b.save(..., final=True)`. The output is a JSON report
at the path you specify via `report_path`, and a summary is printed to the terminal.

### Interpret QA results

| Result | Meaning | Action |
|--------|---------|--------|
| `pass` | No issues | Deliver |
| `pass_with_warnings` | Minor issues flagged | Review warnings, fix user-visible ones |
| `fail` | Errors present | Must fix before delivering |

**Errors (must fix):**
- Off-canvas content — text or shapes extending past slide edges
- Placeholder text — any `[TODO]`, `[chart]`, `lorem ipsum` remaining in slides

**Warnings (fix if user-visible):**
- Text overflow in a text box → reduce font size or split into two slides
- Overlapping shapes → adjust layout coordinates
- Small fonts (< 11pt) → increase or drop the item
- Low contrast → change text colour

### Read the QA report

```python
import json
report = json.load(open("outputs/model_v3_approval.qa.json"))
for issue in report["warnings"] + report["errors"]:
    print(f"Slide {issue['slide']}: {issue['check']} — {issue['message']}")
```

Fix, re-run the build script, verify QA passes. One fix-and-verify cycle is usually enough.

---

## Step 5: Deliver

Tell the user:
- File path to the `.pptx`
- Slide list (what is on each slide)
- Any assumptions you made about the brief
- Any elements they should review (names, dates, data that was inferred)

Do not write a long post-amble. The user can open the file.

---

## Pattern Quick Reference

Patterns 1–19 use a thin-rule header. Patterns 20–23 use a **full-width colored
header bar** (prm-deck-kit style) — more prominent, better for financial audiences.
All patterns work with any palette; the `prm` palette looks best with 20–23.

### Standard patterns (thin-rule header)

| Pattern | Use when | Key inputs |
|---------|----------|------------|
| `status_slide` | Workstream RAG status | workstreams: [{name, status, rag, owner}] |
| `process_slide` | 3–5 step horizontal pipeline | steps: [{name, description}], highlight |
| `timeline_slide` | Roadmap or milestone schedule | milestones: [{date, label, state}] |
| `results_slide` | Model/scenario comparison table | columns, rows, win_col, worst_col |
| `chart_context_slide` | Chart image + findings + so-what | chart_path, headline, bullets, so_what |
| `numbers_slide` | 3–4 large stat callouts on dark bg | stats: [{number, label, context}] |
| `heat_map_slide` | Numeric matrix by row × column | row_labels, col_labels, values (2D float) |
| `waterfall_slide` | Contributions to a net result | items: [{name, value}], total_label |
| `scorecard_slide` | Criteria-based option comparison | criteria, options, scores, weights |
| `annotation_chart_slide` | Chart with event callouts | chart_path, annotations: [{label, x, y}] |
| `two_by_two_slide` | Prioritisation / impact vs effort | x_label, y_label, items: [{name, x, y}] |
| `assertion_evidence_slide` | One big claim + 3–4 proof points | assertion, evidence: [{stat, text}] |
| `diverging_bar_slide` | Side-by-side comparison by cohort | row_labels, left_values, right_values |
| `recommendation_slide` | Decision + rationale + caveats | recommendation, rationale, caveats |
| `kpi_dashboard_slide` | Up to 6 metric cards with deltas | kpis: [{name, value, delta, trend}] |
| `agenda_slide` | Numbered section list | items: [{label, description, active}] |
| `assumption_table_slide` | Model assumptions + sensitivity | assumptions: [{assumption, value, sensitivity}] |
| `bar_chart_slide` | Native column chart | categories, series: [{name, values}] |
| `line_chart_slide` | Native line chart for time series | categories, series: [{name, values}] |

### Visual-bar patterns (full-width header bar, prm-deck-kit style)

| Pattern | Use when | Key inputs |
|---------|----------|------------|
| `three_column_card_slide` | Team, 3-feature breakdown, or 3-property comparison | cards: [{heading, body, tag}], bar_color |
| `two_column_contrast_slide` | Problem vs solution, before vs after, context vs detail | left_panel: {heading, body}, right_panel: {heading, body}, bar_color |
| `numbered_steps_slide` | Vertical pipeline, onboarding flow, or governance process | steps: [{title, description}], bar_color |
| `callout_bar_slide` | Any content slide that ends with a strong governance or closing statement | body (str or list), callout_text, bar_color |

**`bar_color` for visual-bar patterns:**
- Teal (`p["secondary"]` or `"0A9396"`) — solution, process, positive
- Navy (`p["primary"]` or `"0D1B2A"`) — problem, team, status, governance

---

## Common Mistakes to Avoid

- **Never create a slide for every section in the brief** — pick the patterns that best represent the insight, not the structure
- **Never use the same pattern twice in a row** — vary the visual rhythm
- **Never use any pattern more than twice in one deck** — see Pattern Diversity Rule above
- **Never use visual-bar patterns (20–23) more than twice per deck** — they are accent patterns, not the default layout
- **Never use `three_column_card_slide` for narrative content** — it needs genuinely parallel items with real body text in all three cards; use `assertion_evidence_slide` or `callout_bar_slide` for single-argument slides
- **Never leave `key_message` blank** when the slide has a clear takeaway — its absence causes visible blank space below the title
- **Never use `callout_bar_slide` in the middle of a deck** — it is a closing pattern; one per deck, final or penultimate slide only
- **Never default to card patterns when quantitative data is present** — use `numbers_slide`, `kpi_dashboard_slide`, `results_slide`, `heat_map_slide` or a native chart instead
- **Never put more than 5 bullets on a slide** — cut or split
- **Never leave placeholder text** — if you don't have real data for a field, either drop the field or write "TBC" explicitly
- **Always chain patterns onto the same builder** — one `PptxBuilder`, one file
- **Always call `b.save(..., final=True)`** — omitting `final=True` skips QA
