# AI Decision Log

_Generated: 2026-05-25 01:07_

Records decisions taken during deck construction — source analysis, pattern selection, and QA results.

## Run context

- **Source**: `q2_model_review_brief.docx`
- **Audience**: credit risk committee
- **Decision sought**: approve Credit Model v3 for production

---

## Source analysis

- **Sections found**: 12
  - **Source notes** — 2 paragraph(s)
  - **Executive Summary** — 2 paragraph(s) `[DISTIL]`
  - **Agenda** — 6 paragraph(s)
  - **Q2 Performance Snapshot** — 7 paragraph(s)
  - **Programme Status** — 7 paragraph(s) `[DISTIL]`
  - **Model Performance Comparison** — 1 paragraph(s), 1 table(s)
  - **Implementation Plan** — 7 paragraph(s)
  - **Stress Test Assumptions** — 1 paragraph(s), 1 table(s) `[DISTIL]`
  - **Factor Return Attribution** — 1 paragraph(s), 1 table(s)
  - **Historical PMSE Trend** — 1 paragraph(s), 1 table(s)
  - **Core Finding** — 5 paragraph(s) `[DISTIL]`
  - **Implementation Risk Prioritisation** — 5 paragraph(s) `[DISTIL]`
- **Tables at document level**: 4
- **Approximate word count**: 701
- **⚠ Sections needing distillation**: 5 (marked `[DISTIL]` — rewrite into concise slide copy, do not copy verbatim)

---

## QA results

**Status**: **PASSED** (with warnings) — 25 slides, 0 error(s), 42 warning(s)

- WARN  [slide 2] Slide 2 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 2] Slide 2 has ~27 words with no visual - canvas will appear empty.
- WARN  [slide 3] TextBox 3 averages 16 words per bullet - aim for <=15 words each.
- WARN  [slide 3] Slide 3 has ~89 words - aim for <=80 words total per content slide.
- WARN  [slide 3] Slide 3 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 3] Slide 3 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 4] Slide 4 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 4] Slide 4 has ~32 words with no visual - canvas will appear empty.
- WARN  [slide 5] Slide 5 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 5] Slide 5 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 7] Slide 7 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 7] Slide 7 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 8] Slide 8 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 8] Slide 8 has ~18 words with no visual - canvas will appear empty.
- WARN  [slide 9] Slide 9 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 9] Slide 9 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 10] Slide 10 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 10] Slide 10 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 11] Slide 11 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 11] Slide 11 has ~21 words with no visual - canvas will appear empty.
- WARN  [slide 12] Slide 12 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 12] Slide 12 has ~18 words with no visual - canvas will appear empty.
- WARN  [slide 13] Slide 13 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 13] Slide 13 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 14] Slide 14 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 14] Slide 14 has ~30 words with no visual - canvas will appear empty.
- WARN  [slide 15] Slide 15 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 15] Slide 15 has ~21 words with no visual - canvas will appear empty.
- WARN  [slide 16] Slide 16 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 16] Slide 16 has ~21 words with no visual - canvas will appear empty.
- WARN  [slide 17] Slide 17 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 17] Slide 17 has ~22 words with no visual - canvas will appear empty.
- WARN  [slide 18] TextBox 3 averages 16 words per bullet - aim for <=15 words each.
- WARN  [slide 18] Slide 18 has ~88 words - aim for <=80 words total per content slide.
- WARN  [slide 18] Slide 18 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 18] Slide 18 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 19] Slide 19 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 19] Slide 19 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 20] Slide 20 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 20] Slide 20 is text-only with small type (16pt min) - increase body font to >=18pt.
- WARN  [slide 21] Slide 21 is text-only - add a table, chart image, or visual pattern.
- WARN  [slide 21] Slide 21 is text-only with small type (16pt min) - increase body font to >=18pt.

---

## Output

`projects\q2_model_review\outputs\q2_model_review_brief_deck.pptx`

---

# Slide Plan

Prepared by the AI assistant after considering the available pattern catalog.

| # | Source evidence | Assertion | Selected pattern | Why selected | Inputs / assets | AI inference? |
|---|---|---|---|---|---|---|
| 1 | Document title and committee context | Q2 2025 Credit Model Review for committee decision | `title_slide` | Establishes purpose and audience before evidence | Title; committee; date | No |
| 2 | Document section sequence | Five sections organise the decision discussion | `agenda_slide` | The committee needs the meeting route | Major source sections | No |
| 3 | Executive Summary | Approve Credit Model v3 for production deployment in Q3 | `recommendation_slide` | A formal decision is explicitly requested | Recommendation; evidence; residual risk | No |
| 4 | Q2 Performance Snapshot | Six headline metrics support model readiness | `kpi_dashboard_slide` | Multiple executive metrics need fast scanning | Six KPI values and deltas | No |
| 5 | Programme Status | Technology integration is the only at-risk workstream | `status_slide` | Workstreams have status and ownership | Workstream; RAG; owner | No |
| 6 | Model Performance Comparison table | v3 outperforms v2 across all maturity cohorts | `results_slide` | Committee should see exact comparable values | PMSE table | No |
| 7 | Factor Return Attribution table | v3 strengthens equity beta and credit attribution | `bar_chart_slide` | Category magnitudes are best compared visually | Attribution basis points by model | No |
| 8 | Historical PMSE Trend table | v3 improves consistently while v2 plateaus | `line_chart_slide` | Ordered quarterly data expresses a trend | Quarterly PMSE series | No |
| 9 | Stress Test Assumptions table | GDP and rate path are the highest-sensitivity inputs | `assumption_table_slide` | Inputs require values, sources and sensitivity | Assumptions table | No |
| 10 | Implementation Plan | Go-live remains scheduled for August 2025 | `timeline_slide` | Dated delivery milestones drive the read | Milestones; status; dates | No |
| 11 | Core Finding | Short-duration treatment drives 72% of PMSE gain | `assertion_evidence_slide` | One central conclusion should dominate | Finding and evidence statistics | No |
| 12 | Implementation Risk Prioritisation | API delay is the only high-urgency risk | `two_by_two_slide` | Risks are positioned by impact and urgency | Risk labels and axis positions | No |
| 13 | Supporting methodology | Supporting detail remains available | `section_divider` | Signals the shift from decision story to reference | Section label | No |
