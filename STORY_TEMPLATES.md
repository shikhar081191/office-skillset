# Story Templates — Detailed Reference

This file extends `AI_INSTRUCTIONS.md` Step 2. Use it when you need more than
the high-level slide list — it gives per-slot content rules and a code scaffold
you can fill in directly from the brief.

**How to pick a template:**

| If the brief says…                                     | Use template         |
|--------------------------------------------------------|----------------------|
| "approve", "sign off", "decision needed", "get buy-in" | Approval Request     |
| "explain", "introduce", "pitch", "what is", "overview" | Project Explainer    |
| "update", "progress", "where we are", "status"         | Exec Update          |
| "findings", "results", "analysis", "research", "model" | Research Findings    |
| "RAG", "workstream check-in", "sprint review"          | Status Update        |

If multiple match, pick the one whose **opening move** fits the audience best.

---

## Template 1: Approval Request

**When to use:** The audience must make a decision. The deck's job is to make
"yes" the obvious, low-risk answer.

**Opening principle:** Lead with the recommendation — not the journey that got
you there. Execs decide first, then read the evidence. Bury the ask and you lose
them before they see the data.

**Slide sequence:**

| # | Slide | Pattern | Content rule |
|---|-------|---------|--------------|
| 0 | Title | `title_slide` | Deck name + audience + date |
| 1 | Recommendation | `recommendation_slide` | One sentence ask. Rationale: 3 bullets max. Caveats: only real blockers. |
| 2 | Evidence | `numbers_slide` or `assertion_evidence_slide` | 3 headline numbers or 1 assertion + 3 proof points. No narrative — just the strongest facts. |
| 3 | Full Analysis | `results_slide` or `heat_map_slide` | The data behind slide 2. Tables or matrices only — no prose. |
| 4 | Risks & Mitigations | `two_column_contrast_slide` (navy bar) | Left: risks. Right: mitigations. 3–4 pairs max. |
| 5 | Timeline & Ask | `callout_bar_slide` (teal bar) | Body: key milestones as bullets. Callout: the single sentence ask repeated. |

**Code scaffold:**

```python
b = PptxBuilder(palette="prm")
b.title_slide(
    "[Deck title — state the decision, e.g. 'Model v3 Approval']",
    "[Audience] · [Date]"
)

recommendation_slide(
    title="Recommendation",
    recommendation="[One sentence: what you are asking them to approve or decide]",
    rationale=[
        "[Strongest reason — lead with the outcome, not the process]",
        "[Second reason — quantified if possible]",
        "[Third reason — risk of not acting, if relevant]",
    ],
    caveats=["[Only include if there is a real blocker or condition]"],
    key_message="[One-line summary of the recommendation context]",
    builder=b,
)

numbers_slide(
    title="[Headline metric slide title]",
    stats=[
        {"number": "[Key figure]",  "label": "[What it measures]",  "context": "[vs. baseline or target]"},
        {"number": "[Key figure]",  "label": "[What it measures]",  "context": "[vs. baseline or target]"},
        {"number": "[Key figure]",  "label": "[What it measures]",  "context": "[vs. baseline or target]"},
    ],
    builder=b,
)

# Slide 3: use results_slide for a comparison table, or heat_map_slide for a matrix
results_slide(
    title="[Analysis title]",
    columns=["Metric", "[Option A]", "[Option B]", "[Option C]"],
    rows=[
        ["[Metric 1]", "[val]", "[val]", "[val]"],
        ["[Metric 2]", "[val]", "[val]", "[val]"],
    ],
    win_col=2,   # 0-based index of the recommended option's column
    builder=b,
)

two_column_contrast_slide(
    title="Risks and Mitigations",
    left_panel={
        "heading": "Risks",
        "body": "[Risk 1]\n\n[Risk 2]\n\n[Risk 3]",
    },
    right_panel={
        "heading": "Mitigations",
        "body": "[Mitigation 1]\n\n[Mitigation 2]\n\n[Mitigation 3]",
    },
    bar_color="0D1B2A",  # navy
    builder=b,
)

callout_bar_slide(
    title="Timeline and Next Steps",
    body=[
        "[Milestone 1] — [date]",
        "[Milestone 2] — [date]",
        "[Milestone 3] — [date]",
    ],
    callout_text="[Restate the ask: 'We need approval by [date] to begin [phase].']",
    bar_color="0A9396",  # teal
    builder=b,
)

b.save("outputs/[deck_name].pptx", final=True)
```

---

## Template 2: Project Explainer

**When to use:** The audience is unfamiliar with the project. The deck must
build understanding from scratch and leave them with a clear picture of what,
why, and who.

**Opening principle:** Show the problem before the solution. The audience needs
to feel the pain before they value the relief. If you open with "our solution
is…", you've skipped the reason they should care.

**Slide sequence:**

| # | Slide | Pattern | Content rule |
|---|-------|---------|--------------|
| 0 | Title | `title_slide` | Project name + one-line premise |
| 1 | The Problem | `two_column_contrast_slide` (navy bar) | Left: current pain. Right: what this costs (time, risk, money). Keep to 2–3 points each. |
| 2 | The Solution | `assertion_evidence_slide` | Left panel assertion: the core idea in one sentence. Evidence: 3 properties of the solution. |
| 3 | How It Works | `numbered_steps_slide` (navy bar) | 3–5 steps. Each step: short title + one-line description. No jargon. |
| 4 | Governance & Controls | `three_column_card_slide` (teal bar) | 3 cards: human touchpoints, override mechanisms, audit trail. |
| 5 | The Team | `three_column_card_slide` (navy bar) | 3 cards: one per owner/role. Tag = role title. Body = what they own. |
| 6 | Timeline | `timeline_slide` | 4–6 milestones. Mark current one. Keep dates real. |
| 7 | Why It Matters | `callout_bar_slide` (teal bar) | Body: 3 bullets on the broader value. Callout: single sentence on long-term impact. |

**Code scaffold:**

```python
b = PptxBuilder(palette="prm")
b.title_slide(
    "[Project name]",
    "[One-line premise — what this is and for whom] · [Date]"
)

two_column_contrast_slide(
    title="The Problem",
    left_panel={
        "heading": "Today",
        "body": "[Pain point 1]\n\n[Pain point 2]\n\n[Pain point 3]",
        "accent_color": "EE9B00",  # amber left border = problem signal
    },
    right_panel={
        "heading": "The cost",
        "body": "[Quantified impact of pain 1]\n\n[Quantified impact of pain 2]",
    },
    bar_color="0D1B2A",  # navy
    builder=b,
)

assertion_evidence_slide(
    title="The Solution",
    assertion="[Core idea in one plain sentence — no jargon]",
    evidence=[
        {"stat": "[Key property]", "text": "[What this means in practice]"},
        {"stat": "[Key property]", "text": "[What this means in practice]"},
        {"stat": "[Key property]", "text": "[What this means in practice]"},
    ],
    builder=b,
)

numbered_steps_slide(
    title="How It Works",
    steps=[
        {"title": "[Step 1]", "description": "[One sentence — what happens and who does it]"},
        {"title": "[Step 2]", "description": "[One sentence]"},
        {"title": "[Step 3]", "description": "[One sentence]"},
    ],
    bar_color="0D1B2A",  # navy
    builder=b,
)

three_column_card_slide(
    title="Human Controls",
    cards=[
        {"tag": "[Control type]", "heading": "[Control name]", "body": "[What a human reviews or approves]"},
        {"tag": "[Control type]", "heading": "[Control name]", "body": "[What a human reviews or approves]"},
        {"tag": "[Control type]", "heading": "[Control name]", "body": "[What a human reviews or approves]"},
    ],
    bar_color="0A9396",  # teal
    builder=b,
)

three_column_card_slide(
    title="The Team",
    cards=[
        {"tag": "[Role]", "heading": "[Name or team]", "body": "[What they own and decide]"},
        {"tag": "[Role]", "heading": "[Name or team]", "body": "[What they own and decide]"},
        {"tag": "[Role]", "heading": "[Name or team]", "body": "[What they own and decide]"},
    ],
    bar_color="0D1B2A",  # navy
    builder=b,
)

timeline_slide(
    title="Where We Are",
    milestones=[
        {"date": "[Month Year]", "label": "[Milestone]", "state": "past"},
        {"date": "[Month Year]", "label": "[Milestone]", "state": "current"},
        {"date": "[Month Year]", "label": "[Milestone]", "state": "future"},
    ],
    builder=b,
)

callout_bar_slide(
    title="Why It Matters",
    body=[
        "[Benefit 1 — quantified]",
        "[Benefit 2 — quantified]",
        "[Benefit 3 — broader significance]",
    ],
    callout_text="[Long-term impact statement — one sentence, plain language]",
    bar_color="0A9396",  # teal
    builder=b,
)

b.save("outputs/[deck_name].pptx", final=True)
```

---

## Template 3: Exec Update

**When to use:** Recurring or one-off progress update for a senior audience.
They know the project; they want: are we on track, what's the issue, what do
they need to do.

**Opening principle:** Lead with the headline metric — give them the number
before context. Exec time is scarce. If slide 1 is a RAG table, they have to
read to find the signal. If slide 1 is a KPI card, they know in three seconds.

**Slide sequence:**

| # | Slide | Pattern | Content rule |
|---|-------|---------|--------------|
| 0 | Title | `title_slide` | Project + date + "Q[N] Update" or similar |
| 1 | Headline KPIs | `kpi_dashboard_slide` | 3–6 metrics. Include trend arrows. Only metrics they track. |
| 2 | Workstream Status | `status_slide` | RAG per workstream. 1-line status only — no paragraphs. |
| 3 | Key Issue / Decision | `recommendation_slide` | One issue or decision per slide. Don't aggregate issues — pick the most important. |
| 4 | Timeline | `timeline_slide` | Updated since last meeting. Mark current milestone clearly. |
| 5 | Next Steps / Ask | `callout_bar_slide` (navy bar) | Body: 3 bullets of concrete actions with owners and dates. Callout: what you need from this room. |

**Code scaffold:**

```python
b = PptxBuilder(palette="prm")
b.title_slide(
    "[Project name] — [Q/Period] Update",
    "[Audience] · [Date]"
)

kpi_dashboard_slide(
    title="[Period] at a Glance",
    kpis=[
        {"name": "[Metric]", "value": "[Value]", "delta": "[+/-X vs last]", "trend": "up"},
        {"name": "[Metric]", "value": "[Value]", "delta": "[+/-X vs last]", "trend": "down"},
        {"name": "[Metric]", "value": "[Value]", "delta": "[+/-X vs last]", "trend": "flat"},
    ],
    key_message="[One sentence: what these numbers mean together]",
    builder=b,
)

status_slide(
    title="Workstream Status",
    workstreams=[
        {"name": "[Workstream]", "status": "[One-line status]", "rag": "green", "owner": "[Initials]"},
        {"name": "[Workstream]", "status": "[One-line status — what's the issue]", "rag": "amber", "owner": "[Initials]"},
        {"name": "[Workstream]", "status": "[One-line status]", "rag": "red",   "owner": "[Initials]"},
    ],
    key_message="[Summary: N on track, N at risk, N blocked]",
    builder=b,
)

recommendation_slide(
    title="Key Issue",
    recommendation="[The decision or action needed — one sentence]",
    rationale=[
        "[Why this is the key issue]",
        "[What happens if not resolved]",
        "[What resolution looks like]",
    ],
    key_message="[Context: how this arose]",
    builder=b,
)

timeline_slide(
    title="Programme Timeline",
    milestones=[
        {"date": "[Date]", "label": "[Milestone]", "state": "past"},
        {"date": "[Date]", "label": "[Current]",   "state": "current"},
        {"date": "[Date]", "label": "[Next]",       "state": "future"},
    ],
    builder=b,
)

callout_bar_slide(
    title="Next Steps",
    body=[
        "[Action 1] — [Owner] by [Date]",
        "[Action 2] — [Owner] by [Date]",
        "[Action 3] — [Owner] by [Date]",
    ],
    callout_text="[What you need from this audience — decision, budget, resource, or sign-off]",
    bar_color="0D1B2A",  # navy
    builder=b,
)

b.save("outputs/[deck_name].pptx", final=True)
```

---

## Template 4: Research Findings

**When to use:** Presenting analysis, model results, or a study. The audience
needs to understand what you found, trust the method, and know what to do next.

**Opening principle:** Lead with the core finding — not the method. Finance
audiences are busy; if you bury the finding in slide 5 after three slides of
methodology, you've lost them. Show the answer first, then prove it.

**Slide sequence:**

| # | Slide | Pattern | Content rule |
|---|-------|---------|--------------|
| 0 | Title | `title_slide` | Research question as subtitle |
| 1 | Core Finding | `assertion_evidence_slide` | The answer in one sentence. 3 proof points with stats. |
| 2 | Method | `numbered_steps_slide` (navy bar) | 3–4 steps. Enough to understand, not enough to replicate. Skip implementation detail. |
| 3 | Key Result | `numbers_slide` or `kpi_dashboard_slide` | The 3 numbers that matter most. With units and context. |
| 4 | Full Data | `heat_map_slide` or `results_slide` | The complete matrix or table. Don't editorialize — let the data speak. |
| 5 | Sensitivity & Limits | `assumption_table_slide` or `two_column_contrast_slide` | What assumptions were made. What could change the result. |
| 6 | Recommendation | `recommendation_slide` | What to do with this finding. 3 actions, not 10. |

**Code scaffold:**

```python
b = PptxBuilder(palette="prm")
b.title_slide(
    "[Research title]",
    "Research question: [One sentence framing what you set out to answer] · [Date]"
)

assertion_evidence_slide(
    title="Core Finding",
    assertion="[The answer to the research question — one sentence, plain English]",
    evidence=[
        {"stat": "[Key stat]", "text": "[What this stat shows]"},
        {"stat": "[Key stat]", "text": "[What this stat shows]"},
        {"stat": "[Key stat]", "text": "[What this stat shows]"},
    ],
    key_message="[Implication: so what does this mean for the team/decision?]",
    builder=b,
)

numbered_steps_slide(
    title="Method",
    steps=[
        {"title": "[Step]", "description": "[What you did and why]"},
        {"title": "[Step]", "description": "[What you did and why]"},
        {"title": "[Step]", "description": "[What you did and why]"},
    ],
    bar_color="0D1B2A",  # navy
    builder=b,
)

numbers_slide(
    title="Key Results",
    stats=[
        {"number": "[Primary metric]", "label": "[Metric name]", "context": "[vs. baseline]"},
        {"number": "[Secondary]",      "label": "[Metric name]", "context": "[vs. baseline]"},
        {"number": "[Supporting]",     "label": "[Metric name]", "context": "[significance or threshold]"},
    ],
    builder=b,
)

# Slide 4: full data — use heat_map_slide for matrices, results_slide for tables
heat_map_slide(
    title="[Full analysis title]",
    row_labels=["[Row 1]", "[Row 2]", "[Row 3]"],
    col_labels=["[Col 1]", "[Col 2]", "[Col 3]"],
    values=[
        [0.0, 0.0, 0.0],  # replace with real floats
        [0.0, 0.0, 0.0],
        [0.0, 0.0, 0.0],
    ],
    key_message="[What pattern the matrix shows]",
    builder=b,
)

assumption_table_slide(
    title="Assumptions and Sensitivity",
    assumptions=[
        {"assumption": "[Assumption]", "value": "[Value used]", "sensitivity": "high"},
        {"assumption": "[Assumption]", "value": "[Value used]", "sensitivity": "medium"},
        {"assumption": "[Assumption]", "value": "[Value used]", "sensitivity": "low"},
    ],
    key_message="[Which assumption most changes the conclusion if wrong]",
    builder=b,
)

recommendation_slide(
    title="Recommendation",
    recommendation="[What to do, based on this finding — one sentence]",
    rationale=[
        "[Why this action follows from the finding]",
        "[What changes if this action is taken]",
        "[How to monitor whether it's working]",
    ],
    caveats=["[Condition under which this recommendation would change]"],
    builder=b,
)

b.save("outputs/[deck_name].pptx", final=True)
```

---

## Template 5: Status Update

**When to use:** Regular pulse check — sprint review, weekly/monthly RAG, or
standing committee update. Audience knows the project; they want fast signal.

**Opening principle:** KPIs first, always. This is a check-in, not a story.
Get to the red/amber/green in the first 60 seconds.

**Slide sequence:**

| # | Slide | Pattern | Content rule |
|---|-------|---------|--------------|
| 0 | Title | `title_slide` | Project + period (e.g. "Week 22") |
| 1 | KPI Summary | `kpi_dashboard_slide` | 3–6 metrics. Trend arrows mandatory. |
| 2 | Workstream Status | `status_slide` | RAG per stream. One-line status only. Owners visible. |
| 3 | Risks & Priorities | `two_by_two_slide` | Plot open risks by likelihood × impact. Label the top 3. |
| 4 | Timeline | `timeline_slide` | Highlight current milestone. Flag any slippage. |

**Code scaffold:**

```python
b = PptxBuilder(palette="prm")
b.title_slide(
    "[Project name] — Status Update",
    "[Period, e.g. Week 22 / May 2026] · [Audience]"
)

kpi_dashboard_slide(
    title="[Period] — Key Metrics",
    kpis=[
        {"name": "[Metric]", "value": "[Value]", "delta": "[vs. last]", "trend": "up"},
        {"name": "[Metric]", "value": "[Value]", "delta": "[vs. last]", "trend": "down"},
        {"name": "[Metric]", "value": "[Value]", "delta": "[vs. last]", "trend": "flat"},
    ],
    builder=b,
)

status_slide(
    title="Workstream Status",
    workstreams=[
        {"name": "[Workstream]", "status": "[Status — one sentence]", "rag": "green", "owner": "[Initials]"},
        {"name": "[Workstream]", "status": "[Status — one sentence]", "rag": "amber", "owner": "[Initials]"},
    ],
    builder=b,
)

two_by_two_slide(
    title="Risk Priorities",
    x_label="Likelihood", y_label="Impact",
    x_low="Low", x_high="High",
    y_low="Low", y_high="High",
    items=[
        {"name": "[Risk 1]", "x": 0.8, "y": 0.85, "descriptor": "Act now"},
        {"name": "[Risk 2]", "x": 0.4, "y": 0.6},
        {"name": "[Risk 3]", "x": 0.2, "y": 0.3},
    ],
    builder=b,
)

timeline_slide(
    title="Programme Timeline",
    milestones=[
        {"date": "[Date]", "label": "[Milestone]", "state": "past"},
        {"date": "[Date]", "label": "[Current]",   "state": "current"},
        {"date": "[Date]", "label": "[Next]",       "state": "future"},
    ],
    builder=b,
)

b.save("outputs/[deck_name].pptx", final=True)
```

---

## Cross-Template Rules

These apply regardless of which template you use:

- **One decision or finding per slide.** If a slide is trying to say two things, split it.
- **Never use the same pattern twice in a row.** Vary the visual rhythm.
- **Drop any slot you don't have content for.** A missing slide is better than a thin one.
- **The title slide subtitle is always: audience · date.** No exceptions.
- **Text in scaffolds marked `[like this]` must be replaced.** Never deliver a deck with bracket placeholders.
