"""
Private Credit Factor Model Adjustment — Research Presentation
Geltner unsmoothing and validation of factor exposures
"""

from patterns import (
    results_slide,
    assertion_evidence_slide,
    numbers_slide,
)
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Private Credit Factor Models")

# Start the deck
b = PptxBuilder(palette="midnight_executive")

# Title slide
b.title_slide(
    "Private Credit Factor Models",
    subtitle="Return smoothing adjustment and factor exposure validation"
)

# Slide 1: Problem statement (custom content slide)
b.content_slide(
    title="The Problem: Stale Valuations Break Daily Factor Frameworks",
    body=[
        "Private credit instruments mark quarterly at best; daily returns don't exist.",
        "Our standard factor risk framework is built on daily market observations.",
        "Result: Reported returns are smoothed. Raw data gives:",
        "  • Volatilities understated (appear less risky than true underlying)",
        "  • Correlations with public markets artificially compressed",
        "  • Betas near zero (look uncorrelated; actually not)",
        "Everyone knows this. Nobody had a scalable fix until recently.",
    ]
)

# Slide 2: Methodology
b.content_slide(
    title="Approach: Reverse the Smoothing with Geltner Adjustment",
    body=[
        "Insight: Quarterly reported returns are a smoothed approximation of true daily returns.",
        "Geltner method: Reverse-engineer the unsmoothed series from observed returns.",
        "Assume first-order autocorrelation in the smoothing process — standard in appraisals.",
        "Solve for the unsmoothed volatility and correlation structure.",
        "",
        "Pilot: Applied to subset of private credit exposures (40% of book).",
        "Internal consistency checks: dispersions, portfolio manager intuition, cross-sectional patterns.",
    ]
)

# Slide 3: Before and after results (results_slide pattern)
b = results_slide(
    title="Before and After: Raw Returns vs Adjusted Returns",
    columns=[
        "Metric",
        "Raw (Reported)",
        "After Adjustment",
        "Interpretation",
    ],
    rows=[
        ["Volatility",                "~0.8%",  "~1.8%",  "2.3x uplift"],
        ["Correlation, Public IG",    "0.21",   "0.58",   "+0.37 (170% higher)"],
        ["Beta to Rates",             "~0.00",  "0.31",   "Meaningful duration risk"],
        ["Coverage (Book %)",         "100%",   "40%",    "Pilot phase"],
        ["Validation",                "n/a",    "Internal consistency passes", "No ground truth"],
    ],
    win_col=2,  # Adjusted column
    builder=b,
)

# Slide 4: Key findings (numbers_slide pattern)
b = numbers_slide(
    title="Early Results: Material Repricing of Risk",
    stats=[
        {"number": "2.3x",  "label": "Volatility Uplift", "context": "Raw to adjusted"},
        {"number": "0.58",  "label": "IG Correlation",     "context": "Up from 0.21"},
        {"number": "0.31",  "label": "Rates Beta",          "context": "Up from ~0.0"},
        {"number": "40%",   "label": "Book Covered",        "context": "Pilot phase"},
    ],
    builder=b,
)

# Slide 5: What we don't know (assertion_evidence pattern with caveats)
b = assertion_evidence_slide(
    title="Known Unknowns: Validation Gaps and Risks",
    assertion="Results are internally consistent but lack ground truth validation.",
    evidence=[
        {
            "stat": "No",
            "text": "direct observation of true daily returns — can't verify unsmoothing is correct",
        },
        {
            "stat": "Governance",
            "text": "framework not yet approved; MRM conversation hasn't started",
        },
        {
            "stat": "Assumptions",
            "text": "depend critically on first-order autocorrelation model; may not hold across vintage/strategy",
        },
        {
            "stat": "Risk",
            "text": "if adjustment is wrong, we systematically misrepresent risk to PMs and clients",
        },
    ],
    builder=b,
)

# Slide 6: Next steps (custom content slide)
b.content_slide(
    title="What's Next: Three Workstreams",
    body=[
        "1. Scale the model: Extend adjustment to full private credit book (currently 40%).",
        "   • Check robustness across vintage, geography, credit quality.",
        "   • Refine autocorrelation assumptions based on broader dataset.",
        "",
        "2. Integrate into risk reporting: Feed adjusted exposures into standard reports.",
        "   • Portfolio managers see adjusted numbers, not raw ones.",
        "   • QA the reporting pipeline before any external use.",
        "",
        "3. Validate with investment team: Does this match your intuition?",
        "   • Qualitative alignment between adjusted factor exposures and stated strategy.",
        "   • Refine model if there are material mismatches.",
    ]
)

# Slide 7: Governance and next conversation (custom content slide with emphasis)
b.content_slide(
    title="Critical Path: MRM Review and Governance",
    body=[
        "This is a material change to how we measure private credit risk.",
        "Before any internal adoption (let alone client communication), we need:",
        "  • MRM sign-off on methodology and assumptions.",
        "  • Defined limits on how the adjustment can be used.",
        "  • Clear policy on when we revert to unadjusted if assumptions break.",
        "",
        "That conversation hasn't happened yet. It's the gate for moving from pilot to production.",
        "Technical team is ready; governance team needs to be engaged.",
    ]
)

# Save with QA
deck_path = project.output_path("private_credit_factors.pptx")
b.save(
    deck_path,
    final=True,
    report_path=project.qa_path(deck_path.name),
)

print(f"✓ Deck complete: {deck_path}")
