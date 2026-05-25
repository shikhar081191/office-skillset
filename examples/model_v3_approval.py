"""
End-to-end test: Model v3 Approval Request
Built from the example brief in BRIEF_TEMPLATE.md following AI_INSTRUCTIONS.md.

Brief summary:
  DECK PURPOSE:   approval request
  AUDIENCE:       Risk Committee — senior internal
  KEY MESSAGE:    Model v3 ready for production; approve by end of June
  PALETTE:        prm
  MAX SLIDES:     6

Template chosen: Approval Request (STORY_TEMPLATES.md)
Slide plan:
  0. Title
  1. Recommendation          → recommendation_slide
  2. Evidence                → numbers_slide
  3. Full analysis           → results_slide
  4. Risks & Mitigations     → two_column_contrast_slide  (navy bar)
  5. Timeline & Ask          → callout_bar_slide           (teal bar)
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from create_pptx import PptxBuilder, PALETTES
from patterns import (
    recommendation_slide,
    numbers_slide,
    results_slide,
    two_column_contrast_slide,
    callout_bar_slide,
)

prm = PALETTES["prm"]

Path("output").mkdir(exist_ok=True)

b = PptxBuilder(palette="prm")

# ── 0. Title ──────────────────────────────────────────────────────────────────
b.title_slide(
    "Model v3 — Production Approval",
    "Risk Committee · May 2026",
)

# ── 1. Recommendation ─────────────────────────────────────────────────────────
recommendation_slide(
    title="Recommendation",
    recommendation="Approve Model v3 for production deployment in Q3 2026",
    rationale=[
        "PMSE improved 18.2% on holdout data — largest gain since 2021 model refresh",
        "Validated across Rates, Credit, and Equities cohorts; p < 0.01 throughout",
        "End-to-end runtime of 6.4 min is within the 10-minute SLA",
    ],
    caveats=[
        "FX cohort excluded from v3 scope — separate workstream planned for Q4",
        "Legal sign-off on data lineage required before go-live",
    ],
    key_message="Validation is complete. The committee decision is the critical path item.",
    builder=b,
)

# ── 2. Evidence ───────────────────────────────────────────────────────────────
numbers_slide(
    title="Model v3 — Performance at a Glance",
    stats=[
        {"number": "18.2%",  "label": "PMSE improvement",      "context": "vs current production model on holdout data"},
        {"number": "p<0.01", "label": "Statistical significance", "context": "across all three validated asset cohorts"},
        {"number": "6.4 min", "label": "Pipeline runtime",       "context": "within 10-minute SLA — no infrastructure change needed"},
    ],
    builder=b,
)

# ── 3. Full Analysis ──────────────────────────────────────────────────────────
results_slide(
    title="PMSE by Asset Cohort — v3 vs Current Model",
    columns=["Cohort", "Current model", "Model v3", "Improvement"],
    rows=[
        ["Rates",    "0.74", "0.52", "−29.7%"],
        ["Credit",   "0.68", "0.53", "−22.1%"],
        ["Equities", "0.77", "0.71", "−7.8%"],
        ["Bills*",   "0.81", "0.44", "−45.7%"],
    ],
    win_col=3,
    key_message="Bills drive 46pp of the aggregate gain — the factor decomposition change is working.",
    builder=b,
)

# ── 4. Risks & Mitigations ───────────────────────────────────────────────────
two_column_contrast_slide(
    title="Risks and Mitigations",
    left_panel={
        "heading": "Risks",
        "body": (
            "FX cohort not validated\n"
            "Model behaviour untested under rapid FX dislocation scenarios.\n\n"
            "Legal data lineage sign-off outstanding\n"
            "Required before production go-live per governance policy.\n\n"
            "Shadow running period\n"
            "Two-month parallel run adds latency to full deployment."
        ),
        "accent_color": prm["accent"],
    },
    right_panel={
        "heading": "Mitigations",
        "body": (
            "FX cohort scoped to v4\n"
            "Dedicated workstream starts Q4; existing FX model remains in production.\n\n"
            "Legal review in progress\n"
            "Target sign-off by 30 Jun — not on critical path for committee approval.\n\n"
            "Shadow run from Jul — Aug\n"
            "Sufficient runway before Q3 deployment date."
        ),
        "accent_color": prm["secondary"],
    },
    bar_color=prm["primary"],
    key_message="No blocker prevents approval today. Conditions are manageable pre-deployment.",
    builder=b,
)

# ── 5. Timeline & Ask ─────────────────────────────────────────────────────────
callout_bar_slide(
    title="Deployment Timeline",
    body=[
        "Jan – Mar 2026   Development and backtesting complete",
        "Apr – May 2026   Validation across three cohorts — done",
        "May 2026          Risk Committee approval ← we are here",
        "Jun 2026          Legal sign-off and shadow run begins",
        "Aug 2026          Full production deployment",
    ],
    callout_text=(
        "We need Risk Committee approval today to begin the shadow run in June "
        "and meet the August deployment date."
    ),
    bar_color=prm["secondary"],
    builder=b,
)

# ── Save ──────────────────────────────────────────────────────────────────────
out = b.save("output/model_v3_approval.pptx", final=True)
print(f"\nSaved: {out}")
print(f"Slides: {len(b.prs.slides)}")
