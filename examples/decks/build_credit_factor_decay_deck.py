from project_workspace import ProjectWorkspace
from create_pptx import PptxBuilder
import patterns


def build():
    project = ProjectWorkspace("Credit Factor Decay")

    b = PptxBuilder()

    # Title
    b.title_slide(
        title="Credit Factor Decay — Analytics Update",
        subtitle="Internal analytics | Summary & recommended approach"
    )

    # Problem statement
    b.content_slide(
        title="Problem: Factors going stale faster than refresh cycle",
        body=[
            "Post-2022 regime shift: rates, correlations and market structure changed",
            "Existing factors (estimated 2020–21) no longer reflect reality",
            "Current annual refresh is too slow given observed decay",
        ]
    )

    # High-level finding (chart + context placeholder)
    b = patterns.chart_context_slide(
        title="Cross-Asset Analysis",
        chart_path=None,
        headline="Liquidity & momentum factors decay fastest; spread duration is stable",
        bullets=[
            "Five asset classes analysed: IG, HY, EM, Structured, CDS index",
            "Pattern consistent across assets; worst in Structured & EM",
            "Liquidity and momentum show the largest stability loss post-2022",
        ],
        so_what="Trigger-based refresh recommended: targeted, efficient, auditable"
    )

    # Heat map of decay severity (higher = more decay)
    row_labels = ["IG Corporates", "HY Corporates", "EM Credit", "Structured Credit", "CDS Index"]
    col_labels = ["Liquidity", "Momentum", "Spread Duration", "Convexity"]
    values = [
        [0.70, 0.65, 0.30, 0.20],
        [0.75, 0.70, 0.35, 0.25],
        [0.90, 0.85, 0.40, 0.30],
        [0.95, 0.90, 0.45, 0.35],
        [0.60, 0.55, 0.25, 0.15],
    ]
    b = patterns.heat_map_slide(
        title="Factor Stability (higher = more decay)",
        row_labels=row_labels,
        col_labels=col_labels,
        values=values,
        show_values=True,
    )

    # Key numbers
    stats = [
        {"number": "~3x", "label": "Understated volatility", "context": "Stale factors understate realised vol by ~3x"},
        {"number": "+1.3pp", "label": "VaR understatement", "context": "Average VaR understated by ~1.3 percentage points"},
        {"number": "Structured & EM", "label": "Worst affected", "context": "Biggest gaps observed in structured products and EM blend"},
    ]
    b = patterns.numbers_slide(title="Risk Impact (summary)", stats=stats)

    # Results comparison stale vs refreshed
    columns = ["Metric", "Stale model", "Refreshed model"]
    rows = [
        ["Volatility (avg)", "5%", "15%"],
        ["VaR (99%, 1-day)", "0.8%", "2.1%"],
        ["Worst-case gap (Structured/EM)", "0.5%", "2.0%"],
    ]
    b = patterns.results_slide(title="Stale vs Refreshed — Representative metrics", columns=columns, rows=rows, win_col=2, worst_col=1)

    # Prioritisation: which factors to fix
    items = [
        {"name": "Liquidity", "x": 0.2, "y": 0.95, "descriptor": "High impact, simple"},
        {"name": "Quality", "x": 0.25, "y": 0.9, "descriptor": "High impact, simple"},
        {"name": "Spread Duration", "x": 0.85, "y": 0.8, "descriptor": "High impact, complex (Q3)"},
        {"name": "Convexity", "x": 0.6, "y": 0.35, "descriptor": "Low impact, can wait"},
    ]
    b = patterns.two_by_two_slide(
        title="Prioritisation: Impact vs Complexity",
        x_label="Complexity →",
        y_label="Impact →",
        x_low="Low", x_high="High", y_low="Low", y_high="High",
        items=items
    )

    # Options for refresh strategy
    steps = [
        {"name": "Annual batch", "description": "Status quo — cheap but too slow"},
        {"name": "Quarterly rolling", "description": "Faster, more operational effort"},
        {"name": "Trigger-based refresh", "description": "Re-estimate when a stability test fires — targeted and auditable"},
        {"name": "Continuous online", "description": "Always-on updating — operationally intensive, MRM concerns"},
    ]
    b = patterns.process_slide(title="Refresh Options — trade-offs", steps=steps, highlight=2,
                               footnote="Recommendation: trigger-based refresh (targeted, audit-friendly)")

    # Next steps timeline
    milestones = [
        {"date": "June", "label": "Benedek validates trigger framework", "state": "current"},
        {"date": "Late June", "label": "PM alignment sessions", "state": "future"},
        {"date": "July", "label": "MRM conversation (earliest)", "state": "future"},
        {"date": "Q3", "label": "Production integration (pending sign-off)", "state": "future"},
    ]
    b = patterns.timeline_slide(title="Next Steps & Timing", milestones=milestones)

    # Clarifying message
    b = b.content_slide(
        title="Note: not a model failure",
        body=[
            "Original factors were correct for the regime they were estimated in.",
            "Regime shifted post-2022 — factors decayed as expected.",
            "We detected the issue and propose a targeted trigger-based refresh.",
        ]
    )

    # Appendix / methodology
    b = b.section_divider("Appendix: Methodology & tests")
    b = b.content_slide(
        title="Methodology (appendix)",
        body=[
            "Stability tested across daily exposures and re-estimated factor sets.",
            "Comparisons made across five asset classes with representative books.",
            "Full technical notes available on request.",
        ]
    )

    deck_path = project.output_path("credit_factor_decay.pptx")
    out = b.save(str(deck_path), final=True, report_path=project.qa_path(deck_path.name))
    project.record_workflow("credit_factor_decay_build", outputs=[out], patterns=["chart_context", "heat_map", "numbers", "results", "two_by_two", "process", "timeline"]) 
    print(f"Created deck: {out}")


if __name__ == "__main__":
    build()
