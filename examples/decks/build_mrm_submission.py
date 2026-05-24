#!/usr/bin/env python
"""
MRM Submission Deck: Full Revaluation Framework
Policy-driven validation approach with minimum viable metrics
"""

from patterns import (
    assertion_evidence_slide,
    numbers_slide,
    process_slide,
    status_slide,
    timeline_slide,
)
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("MRM Submission - Full Reval Framework")

# Title slide
b = PptxBuilder(palette="midnight_executive")
b.title_slide(
    "Full Revaluation Framework",
    subtitle="MRM Submission — Governance & Validation Plan",
)

# Slide 2: Opening assertion — governance angle
b = assertion_evidence_slide(
    title="Governance Approach",
    assertion="Policy-driven validation framework with clear accountability and flexibility",
    evidence=[
        {"stat": "Proactive", "text": "Advancing MRM governance before future requirements"},
        {"stat": "Flexible", "text": "Minimum viable proposal — designed for MRM-driven enhancements"},
        {"stat": "Transparent", "text": "Measurable metrics with 18-month delivery roadmap"},
    ],
    builder=b,
)

# Slide 3: Value proposition — full reval concept
b.content_slide(
    title="Full Revaluation vs. Parametric",
    body=[
        "Full revaluation reprices instruments directly under shocked market states",
        "Parametric approach linearly extrapolates from base-case sensitivities",
        "Question: How well does full reval actually perform, and can we prove it?",
    ],
    dark_bg=False,
)

# Slide 4: Key metrics
b.content_slide(
    title="Validation Metrics",
    body=[
        "PMSE (Parametric Mean Squared Error): Squared difference between full reval P&L and benchmark, normalised",
        "Parametric Improvement Ratio: Parametric PMSE ÷ Full Reval PMSE",
        "  • Ratio > 1.0 = Full reval outperforms linear",
        "  • Ratio ≤ 1.0 = Triggers investigation",
        "Thresholds intentionally TBD — allows iterative refinement with MRM feedback",
    ],
    dark_bg=False,
)

# Slide 5: Key numbers
b = numbers_slide(
    title="Scope & Roadmap at a Glance",
    stats=[
        {"number": "26", "label": "Asset Types in Scope", "context": "Equities, FX, commodities, rates, derivatives"},
        {"number": "2", "label": "Metrics Submitted", "context": "PMSE, Parametric Improvement Ratio"},
        {"number": "7", "label": "Pipeline Steps", "context": "Pull, consolidate, tag, filter, enrich, calculate, output"},
        {"number": "Sep 30", "label": "MRM Sign-off Target", "context": "Full validation and RAG scoring complete"},
    ],
    builder=b,
)

# Slide 6: Process — 7-step pipeline
b = process_slide(
    title="Data Pipeline — 7-Step Workflow",
    steps=[
        {"name": "Pull Returns", "description": "Retrieve parametric and full reval returns from RAS via BRPC API"},
        {"name": "Consolidate", "description": "Merge data feeds and resolve inconsistencies across sources"},
        {"name": "Tag & Filter", "description": "Classify instruments by cohort; apply initial quality filters"},
        {"name": "Enrich", "description": "Add asset attributes and regime classification"},
        {"name": "Realized Returns", "description": "Calculate actual P&L outcomes for validation"},
        {"name": "Metrics (L1/L2/L3)", "description": "Compute PMSE and improvement ratios at three cohort levels"},
        {"name": "Output & Report", "description": "Generate structured deliverables for MRM review"},
    ],
    builder=b,
)

# Slide 7: Status — workstreams
b = status_slide(
    title="Team Status & Dependencies",
    workstreams=[
        {
            "name": "Data Pipeline",
            "status": "BRPC API integration 60% complete; full scope by end May",
            "rag": "green",
            "owner": "PR",
        },
        {
            "name": "Metrics Prototype",
            "status": "Running on 4 tickers (USD/EUR govts & corps); logic validation complete",
            "rag": "green",
            "owner": "HE",
        },
        {
            "name": "PnL Data Blocker",
            "status": "David's portfolios can't pull data; needs resolution by May 15 or escalate",
            "rag": "red",
            "owner": "DM",
        },
        {
            "name": "QA Briefing",
            "status": "Patrick not yet briefed; scheduling required",
            "rag": "amber",
            "owner": "PQ",
        },
    ],
    builder=b,
)

# Slide 8: Timeline
b = timeline_slide(
    title="Delivery Roadmap",
    milestones=[
        {"date": "Apr 20", "label": "Design Doc to Claire", "state": "past"},
        {"date": "Apr 30", "label": "MRM Submission", "state": "past"},
        {"date": "May 15", "label": "Engineering Blocker Resolved", "state": "current"},
        {"date": "May 31", "label": "Full Pipeline & 6 Metrics", "state": "future"},
        {"date": "Jun 30", "label": "RAG Scoring & Dry Runs", "state": "future"},
        {"date": "Sep 30", "label": "MRM Sign-off", "state": "future"},
    ],
    builder=b,
)

# Save with QA
deck_path = project.output_path("mrm_submission_framework.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(f"✓ Deck complete: {deck_path}")
