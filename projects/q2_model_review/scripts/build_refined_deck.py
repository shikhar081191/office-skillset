"""Refined build script: Q2 Credit Model Review deck.

Replaces the first-pass bullet slides with proper patterns across all 12 content sections.
Covers: title, agenda, recommendation, kpi_dashboard, status, results, bar_chart,
        line_chart, assumption_table, timeline, assertion_evidence, two_by_two.
Plus: source_label, speaker_notes, ai_label where appropriate.

Usage:
    python -m projects.q2_model_review.scripts.build_refined_deck
    python -m projects.q2_model_review.scripts.build_refined_deck --project "Q2 Model Review Dry Run"
"""

import argparse

from artifact_utilities import ArtifactUtilities
from decision_log import DecisionLog
from patterns import (
    status_slide, results_slide, timeline_slide, assertion_evidence_slide,
    two_by_two_slide, agenda_slide, recommendation_slide, kpi_dashboard_slide,
    assumption_table_slide, bar_chart_slide, line_chart_slide,
)
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the refined Q2 Credit Model Review deck.")
    parser.add_argument(
        "--project",
        default="Q2 Model Review",
        help="Existing project workspace name to receive the refined output.",
    )
    parser.add_argument(
        "--output-name",
        default="q2_model_review_refined.pptx",
        help="Filename for the refined deck within the project's outputs folder.",
    )
    return parser.parse_args()


args = parse_args()

# ── project ───────────────────────────────────────────────────────────────────
project = ProjectWorkspace(args.project)

# ── log slide plan BEFORE writing any slides ──────────────────────────────────
log = DecisionLog(project.working_path("decision_log.md"))
log.record_slide_plan([
    {"section": "Cover", "source": "Document title and committee context", "pattern": "title_slide",
     "assertion": "Q2 2025 Credit Model Review for committee decision",
     "rationale": "Establishes purpose and audience before evidence", "inputs": "Title; committee; date", "inference": "No"},
    {"section": "Agenda", "source": "Document section sequence", "pattern": "agenda_slide",
     "assertion": "Five sections organise the decision discussion",
     "rationale": "The committee needs the meeting route", "inputs": "Major source sections", "inference": "No"},
    {"section": "Executive Summary", "source": "Executive Summary", "pattern": "recommendation_slide",
     "assertion": "Approve Credit Model v3 for production deployment in Q3",
     "rationale": "A formal decision is explicitly requested", "inputs": "Recommendation; evidence; residual risk", "inference": "No"},
    {"section": "Q2 Performance Snapshot", "source": "Q2 Performance Snapshot", "pattern": "kpi_dashboard_slide",
     "assertion": "Six headline metrics support model readiness",
     "rationale": "Multiple executive metrics need fast scanning", "inputs": "Six KPI values and deltas", "inference": "No"},
    {"section": "Programme Status", "source": "Programme Status", "pattern": "status_slide",
     "assertion": "Technology integration is the only at-risk workstream",
     "rationale": "Workstreams have status and ownership", "inputs": "Workstream; RAG; owner", "inference": "No"},
    {"section": "Model Performance", "source": "Model Performance Comparison table", "pattern": "results_slide",
     "assertion": "v3 outperforms v2 across all maturity cohorts",
     "rationale": "Committee should see exact comparable values", "inputs": "PMSE table", "inference": "No"},
    {"section": "Factor Attribution", "source": "Factor Return Attribution table", "pattern": "bar_chart_slide",
     "assertion": "v3 strengthens equity beta and credit attribution",
     "rationale": "Category magnitudes are best compared visually", "inputs": "Attribution basis points by model", "inference": "No"},
    {"section": "Historical PMSE Trend", "source": "Historical PMSE Trend table", "pattern": "line_chart_slide",
     "assertion": "v3 improves consistently while v2 plateaus",
     "rationale": "Ordered quarterly data expresses a trend", "inputs": "Quarterly PMSE series", "inference": "No"},
    {"section": "Stress Test Assumptions", "source": "Stress Test Assumptions table", "pattern": "assumption_table_slide",
     "assertion": "GDP and rate path are the highest-sensitivity inputs",
     "rationale": "Inputs require values, sources and sensitivity", "inputs": "Assumptions table", "inference": "No"},
    {"section": "Implementation Plan", "source": "Implementation Plan", "pattern": "timeline_slide",
     "assertion": "Go-live remains scheduled for August 2025",
     "rationale": "Dated delivery milestones drive the read", "inputs": "Milestones; status; dates", "inference": "No"},
    {"section": "Core Finding", "source": "Core Finding", "pattern": "assertion_evidence_slide",
     "assertion": "Short-duration treatment drives 72% of PMSE gain",
     "rationale": "One central conclusion should dominate", "inputs": "Finding and evidence statistics", "inference": "No"},
    {"section": "Risk Prioritisation", "source": "Implementation Risk Prioritisation", "pattern": "two_by_two_slide",
     "assertion": "API delay is the only high-urgency risk",
     "rationale": "Risks are positioned by impact and urgency", "inputs": "Risk labels and axis positions", "inference": "No"},
    {"section": "Appendix", "source": "Supporting methodology", "pattern": "section_divider",
     "assertion": "Supporting detail remains available",
     "rationale": "Signals the shift from decision story to reference", "inputs": "Section label", "inference": "No"},
])

# ── build ─────────────────────────────────────────────────────────────────────
b = PptxBuilder(palette="blackrock")

# 1. Title
b.title_slide(
    "Q2 2025 Credit Model Review",
    subtitle="Credit Risk Committee  |  24 May 2025",
)

# 2. Agenda
agenda_slide(
    "Agenda",
    items=[
        {"label": "Executive Summary",         "description": "Recommendation and key findings"},
        {"label": "Model Performance",         "description": "Quantitative results and holdout validation"},
        {"label": "Stress Testing",            "description": "Macro scenario assumptions and sensitivities"},
        {"label": "Implementation Plan",       "description": "Go-live timeline and workstream status"},
        {"label": "Appendix",                  "description": "Methodology and supporting data"},
    ],
    key_message="Five sections: summary decision, model evidence, stress scenarios, implementation timeline, and supporting data.",
    builder=b,
)

# 3. Recommendation
recommendation_slide(
    title="Credit Model v3 - Decision",
    recommendation="Approve Credit Model v3 for full production deployment in Q3 2025.",
    rationale=[
        "PMSE improved 18pp on holdout versus Model v2",
        "Independent validation confirmed on 14 May 2025",
        "90-day parallel run shows immaterial capital impact",
        "All six regulatory stress scenarios passed",
    ],
    caveats=[
        "Technology integration is delayed; August go-live remains achievable",
    ],
    key_message="Evidence supports committee approval today.",
    builder=b,
)
b.speaker_notes(
    "The technology delay does not affect go-live date: the parallel run buffer absorbs it. "
    "Anticipated question: 'Is the capital impact really immaterial?' — yes, PRA threshold is 5% and we are at 0.3%."
)

# 4. KPI Dashboard
kpi_dashboard_slide(
    "Q2 2025 - Portfolio at a Glance",
    kpis=[
        {"name": "PMSE (Holdout)",    "value": "0.046",  "delta": "-23% vs v2",  "trend": "up"},
        {"name": "Sharpe Ratio",      "value": "1.18",   "delta": "+0.08 vs Q1", "trend": "up"},
        {"name": "AUM Covered",       "value": "$4.7B",  "delta": "",            "trend": "flat"},
        {"name": "Model Coverage",    "value": "98.3%",  "delta": "",            "trend": "flat"},
        {"name": "Info Ratio",        "value": "0.71",   "delta": "",            "trend": "flat"},
        {"name": "Stress Pass Rate",  "value": "100%",   "delta": "All 6 scenarios", "trend": "up"},
    ],
    key_message="All six headline metrics are positive vs prior quarter; PMSE improved 23% vs the v2 baseline.",
    builder=b,
)
b.source_label("Credit Risk Modelling Team, Q2 2025 Run")

# 5. Programme Status
status_slide(
    "Programme Status - Path to Production",
    workstreams=[
        {"name": "Model Development",      "status": "Complete; v3 signed off 28 Apr", "rag": "green", "owner": "CRM"},
        {"name": "Independent Validation", "status": "Complete; Accenture sign-off 14 May", "rag": "green", "owner": "EV"},
        {"name": "Regulatory Submission",  "status": "In progress; legal review by 28 May", "rag": "amber", "owner": "Legal"},
        {"name": "Technology Integration", "status": "At risk; API patch in vendor testing", "rag": "red",   "owner": "IT"},
        {"name": "Parallel Run",           "status": "Complete; all cohorts within tolerance", "rag": "green", "owner": "CRM"},
        {"name": "Documentation",          "status": "80% complete; sign-off target 7 Jun", "rag": "amber", "owner": "CRM"},
    ],
    key_message="4 of 6 workstreams complete; technology integration at risk but go-live date remains August 2025.",
    builder=b,
)
b.speaker_notes(
    "Technology integration (red): vendor has a patch in testing. Current ETA is +2 weeks but "
    "the parallel run buffer means go-live is still August. IT lead will present an update if asked."
)

# 6. Model Performance Comparison
results_slide(
    "PMSE by Model Version and Cohort",
    columns=["Cohort", "Model v1", "Model v2", "Model v3"],
    rows=[
        ["Short (<1Y)",   "0.078", "0.048", "0.031"],
        ["Medium (1-5Y)", "0.085", "0.061", "0.049"],
        ["Long (>5Y)",    "0.091", "0.072", "0.058"],
        ["All Cohorts",   "0.084", "0.060", "0.046"],
    ],
    win_col=3,
    worst_col=1,
    key_message="Model v3 outperforms v2 by 23% PMSE across all maturity cohorts; sharpest gain in the short-duration segment.",
    builder=b,
)
b.source_label("Out-of-time holdout set, Jan-Mar 2025. Lower PMSE = better.")

# 7. Factor Return Attribution
bar_chart_slide(
    "Factor Return Attribution - Q2 2025 (bps)",
    categories=["Equity Beta", "Duration", "Credit", "FX Carry", "Residual"],
    series=[
        {"name": "Model v2", "values": [28, -12, 15, 6, -4]},
        {"name": "Model v3", "values": [32, -8,  18, 6, -2]},
    ],
    footnote="Source: MSCI Barra. Gross of fees. Figures in basis points.",
    key_message="v3 improves on equity beta and credit; duration drag is smaller and residual noise is reduced.",
    builder=b,
)

# 8. Historical PMSE Trend
line_chart_slide(
    "PMSE Trend - Parallel Run Period",
    categories=["Q3 23", "Q4 23", "Q1 24", "Q2 24", "Q3 24", "Q4 24", "Q1 25", "Q2 25"],
    series=[
        {"name": "Model v2", "values": [0.065, 0.063, 0.061, 0.060, 0.062, 0.061, 0.060, 0.060]},
        {"name": "Model v3", "values": [0.058, 0.055, 0.052, 0.049, 0.048, 0.047, 0.047, 0.046]},
    ],
    footnote="Lower PMSE = better. 90-day parallel run, Jan 2023 - May 2025. Source: CRM Team.",
    key_message="v3 PMSE has declined every quarter; v2 has plateaued since Q3 2023 — the gap is structural, not cyclical.",
    builder=b,
)

# 9. Stress Test Assumptions
assumption_table_slide(
    "Stress Scenario Assumptions",
    assumptions=[
        {"assumption": "GDP growth (2025)",   "value": "-1.5%",   "source": "IMF WEO",        "sensitivity": "high"},
        {"assumption": "Rate path (10Y UST)", "value": "+75bps",  "source": "Fed dots",        "sensitivity": "high"},
        {"assumption": "IG spread widening",  "value": "+120bps", "source": "Internal",        "sensitivity": "medium"},
        {"assumption": "Equity drawdown",     "value": "-20%",    "source": "Internal",        "sensitivity": "high"},
        {"assumption": "FX (EURUSD)",         "value": "0.95",    "source": "Bloomberg",       "sensitivity": "low"},
        {"assumption": "Recovery rate",       "value": "40%",     "source": "S&P LossStats",   "sensitivity": "medium"},
    ],
    footnote="1-in-20 adverse scenario. Not a point forecast. Coherent stress, not individual extremes.",
    key_message="GDP and rate path are the high-sensitivity inputs; the model passed all six macro scenarios.",
    builder=b,
)

# 10. Implementation Timeline
timeline_slide(
    "Implementation Plan - Credit Model v3",
    milestones=[
        {"date": "Feb 25", "label": "Model Development",     "state": "past"},
        {"date": "Mar 25", "label": "Internal Validation",   "state": "past"},
        {"date": "May 25", "label": "External Validation",   "state": "current"},
        {"date": "Jun 25", "label": "Regulatory Submission", "state": "future"},
        {"date": "Jul 25", "label": "Tech Integration",      "state": "future"},
        {"date": "Aug 25", "label": "Production Go-Live",    "state": "future"},
    ],
    key_message="On current milestone (May); three steps remain and go-live is confirmed for August 2025.",
    builder=b,
)

# 11. Core Finding
assertion_evidence_slide(
    title="Core Finding",
    assertion="Bills drive 72% of the PMSE improvement",
    evidence=[
        {"stat": "-35%",      "text": "PMSE in short-duration cohort: 0.048 to 0.031"},
        {"stat": "72%",       "text": "Of total gain from vintage and prepayment factor treatment"},
        {"stat": "4pp",       "text": "Additional gain from credit spread factor, medium and long cohorts"},
        {"stat": "Confirmed", "text": "External panel (Jan 2020-Dec 2024) validates finding out-of-sample"},
    ],
    key_message="Short-duration vintage treatment explains the majority of the gain; external validation confirms it is not period-specific.",
    builder=b,
)
b.source_label("Accenture independent validation report, 14 May 2025", x=6.45, width=6.1)

# 12. Risk Prioritisation
two_by_two_slide(
    title="Implementation Risk Prioritisation",
    x_label="Urgency",
    y_label="Impact",
    x_low="Low urgency",
    x_high="High urgency",
    y_low="Low impact",
    y_high="High impact",
    items=[
        {"name": "API Delay",     "descriptor": "Mitigated by parallel run", "x": 0.82, "y": 0.82},
        {"name": "Vintage Drift", "descriptor": "Monthly monitoring agreed",  "x": 0.18, "y": 0.80},
        {"name": "Training Gap",  "descriptor": "3 workshops in June",        "x": 0.52, "y": 0.48},
        {"name": "Doc Lag",       "descriptor": "Post-launch completion ok",  "x": 0.18, "y": 0.18},
    ],
    key_message="API delay is the only high-urgency risk; it is mitigated and the remaining three risks are manageable.",
    builder=b,
)

# 13. Appendix divider
b.section_divider("Appendix")

# ── save with QA ──────────────────────────────────────────────────────────────
deck_path = project.output_path(args.output_name)
qa_path = project.qa_path(args.output_name)
b.save(deck_path, final=True, report_path=qa_path)
rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(deck_path)

project.register_output(
    output_path=deck_path,
    qa_report_path=qa_path,
    patterns=[
        "title_slide", "agenda_slide", "recommendation_slide", "kpi_dashboard_slide",
        "status_slide", "results_slide", "bar_chart_slide", "line_chart_slide",
        "assumption_table_slide", "timeline_slide", "assertion_evidence_slide",
        "two_by_two_slide", "section_divider",
    ],
    notes="Refined deck: all 12 content patterns, key messages, visual improvements applied.",
    working_files=[project.working_path("slide_plan.md")],
    qa_report_paths=[rendered_qa_path],
)

print(f"\nDeck: {deck_path}")
print(f"Log:  {project.working_path('decision_log.md')}")
