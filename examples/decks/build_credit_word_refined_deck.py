from project_workspace import ProjectWorkspace
from patterns import (
    assertion_evidence_slide,
    numbers_slide,
    process_slide,
    recommendation_slide,
    timeline_slide,
)

project = ProjectWorkspace("Credit Word Refined")
from create_pptx import PptxBuilder
b = PptxBuilder(palette="blackrock")

b.title_slide(
    "Credit Factor Decay",
    subtitle="Stability analysis and refresh strategy for credit factors",
)

assertion_evidence_slide(
    title="Problem diagnosis",
    assertion="Credit factor decay is faster than annual refresh.",
    evidence=[
        {"stat": "Post-2022", "text": "Rates, correlations and market structure have shifted."},
        {"stat": "", "text": "Liquidity and momentum factors are decaying fastest, especially in structured credit and EM."},
        {"stat": "3× / 1.3pp", "text": "Stale factors understate volatility by ~3× and VaR by ~1.3pp."},
    ],
    builder=b,
)

numbers_slide(
    title="Risk impact summary",
    stats=[
        {"number": "3×", "label": "Volatility understatement", "context": "Stale factors understate risk by roughly three times."},
        {"number": "1.3pp", "label": "VaR gap", "context": "Average underestimation from stale credit factors."},
        {"number": "Structured / EM", "label": "Worst affected books", "context": "Largest stale-model divergence."},
    ],
    builder=b,
)

process_slide(
    title="Refresh options evaluated",
    steps=[
        {"name": "Annual batch", "description": "Current cadence; simple but too slow for regime shifts."},
        {"name": "Quarterly rolling", "description": "More regular updates, but still calendar-driven."},
        {"name": "Trigger-based refresh", "description": "Update only when stability tests detect change."},
        {"name": "Continuous online", "description": "Most adaptive but operationally and MRM complex."},
    ],
    highlight=2,
    builder=b,
)

recommendation_slide(
    title="Recommended approach",
    recommendation="Adopt trigger-based refresh for credit factors.",
    rationale=[
        "Targets updates only when data signals a regime shift.",
        "Balances responsiveness with MRM and production feasibility.",
        "Reduces stale-factor risk without unnecessary model churn.",
    ],
    caveats=[
        "Trigger design and validation must be complete by June.",
        "PM alignment is needed before MRM submission.",
    ],
    builder=b,
)

timeline_slide(
    title="Next milestones",
    milestones=[
        {"date": "June", "label": "Validate trigger framework", "state": "current"},
        {"date": "Late June", "label": "PM alignment sessions", "state": "future"},
        {"date": "July", "label": "Begin MRM conversation", "state": "future"},
        {"date": "Q3", "label": "Production integration pending sign-off", "state": "future"},
    ],
    builder=b,
)

deck_path = project.output_path("credit_word_refined_deck.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(deck_path)
