"""
Private credit signal adjustment research deck.
"""

from patterns import (
    process_slide,
    results_slide,
    numbers_slide,
    assertion_evidence_slide,
)
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Private Credit Signal Adjustment")
b = PptxBuilder(palette="midnight_executive")

b.title_slide(
    "Private Credit Signal Adjustment",
    subtitle="Unsmoothing quarterly returns, checking factor exposures, and surfacing validation gaps"
)

b.content_slide(
    "The Problem",
    [
        "Private credit valuations are quarterly at best; daily returns are not observed.",
        "Our standard factor risk framework assumes daily market-price data.",
        "Reported returns are therefore smoothed, which distorts risk signals:",
        "  • Volatility is understated",
        "  • Public-market correlations appear artificially low",
        "  • Rates and credit betas move toward zero",
        "This means raw private credit input can give a misleading risk picture.",
    ]
)

b.content_slide(
    "Approach",
    [
        "Treat reported quarterly returns as a smoothed version of true economic returns.",
        "Apply a Geltner-style unsmoothing adjustment to recover a more plausible daily volatility profile.",
        "Evaluate the adjusted series through internal consistency checks, not ground truth.",
        "Pilot coverage is roughly 40% of the private credit book.",
    ]
)

results_slide(
    title="Before vs After: Risk Picture Changes Dramatically",
    columns=["Metric", "Raw reported series", "Adjusted series", "Why it matters"],
    rows=[
        ["Volatility", "Understated", "~2.3x higher", "Risk is materially larger"],
        ["Correlation to public IG", "0.21", "0.58", "Public market sensitivity is stronger"],
        ["Beta to rates", "~0.00", "0.31", "Duration/rate risk is meaningful"],
        ["Pilot coverage", "N/A", "~40% of book", "Proof of concept stage"],
        ["Validation", "Missing ground truth", "Plausible internal consistency", "Still uncertain"],
    ],
    win_col=2,
    builder=b,
)

numbers_slide(
    title="Early Quantitative Findings",
    stats=[
        {"number": "2.3x", "label": "Volatility uplift", "context": "Adjusted vs reported"},
        {"number": "0.58", "label": "IG correlation", "context": "Up from 0.21"},
        {"number": "0.31", "label": "Rates beta", "context": "From near zero"},
        {"number": "40%", "label": "Book covered", "context": "Pilot sample"},
    ],
    builder=b,
)

assertion_evidence_slide(
    title="What We Don\'t Know Yet",
    assertion="The adjustment changes the risk picture, but it is not fully validated.",
    evidence=[
        {"stat": "Ground truth", "text": "True daily returns are unobservable for private credit."},
        {"stat": "Assumptions", "text": "The unsmoothing relies on a specific autocorrelation model."},
        {"stat": "Governance", "text": "MRM review and formal policy are not yet in place."},
        {"stat": "Model risk", "text": "A wrong adjustment could systematically misstate risk."},
    ],
    builder=b,
)

process_slide(
    title="Next Steps",
    highlight=0,
    steps=[
        {
            "name": "Scale the model",
            "description": "Extend unsmoothing to the full private credit book and test across vintages."
        },
        {
            "name": "Embed in reporting",
            "description": "Integrate adjusted exposures into standard risk reports for PM visibility."
        },
        {
            "name": "Qualitative validation",
            "description": "Work with investment teams to confirm factor exposure intuition."
        },
    ],
    builder=b,
)

b.content_slide(
    "Governance and Controls",
    [
        "This is a material measurement change; we need formal governance before adoption.",
        "MRM sign-off, usage limits, and escalation rules should be defined up front.",
        "The pilot is promising, but it should not feed client or external reporting yet.",
        "The first practical gate is a governance review with risk and investment stakeholders.",
    ]
)

output_path = project.output_path("private_credit_signal_adjustment.pptx")
b.save(output_path, final=True, report_path=project.qa_path(output_path.name))
print(f"Deck generated: {output_path}")
