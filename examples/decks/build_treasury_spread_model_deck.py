"""Build the Treasury Spread Model results and recommendation deck."""

from PIL import Image, ImageDraw, ImageFont

from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace
from patterns import (
    BLACKROCK,
    annotation_chart_slide,
    assertion_evidence_slide,
    diverging_bar_slide,
    heat_map_slide,
    numbers_slide,
    process_slide,
    results_slide,
    scorecard_slide,
    status_slide,
    timeline_slide,
    two_by_two_slide,
    waterfall_slide,
)


PROJECT = ProjectWorkspace("Treasury Spread Model", create=False)
CHART_PATH = PROJECT.asset_path("treasury_bill_bond_correlation_placeholder.png")
DECK_PATH = PROJECT.output_path("treasury_spread_model_results_recommendations.pptx")


def _font(size):
    try:
        return ImageFont.truetype("arial.ttf", size)
    except OSError:
        return ImageFont.load_default()


def build_correlation_chart():
    """Create an illustrative regime-break image for the annotation slide."""
    PROJECT.ensure()
    p = BLACKROCK
    image = Image.new("RGB", (1400, 760), "#" + p["background"])
    draw = ImageDraw.Draw(image)
    title_font = _font(28)
    label_font = _font(18)
    small_font = _font(15)

    left, top, right, bottom = 105, 90, 1325, 650
    zero_y = 365
    break_x = 770

    draw.text((left, 25), "Illustrative rolling bill-bond correlation", fill="#" + p["primary"],
              font=title_font)
    draw.line((left, top, left, bottom), fill="#" + p["border"], width=3)
    draw.line((left, zero_y, right, zero_y), fill="#" + p["muted_text"], width=2)
    draw.line((left, bottom, right, bottom), fill="#" + p["border"], width=3)

    draw.text((30, top - 8), "+0.6", fill="#" + p["muted_text"], font=label_font)
    draw.text((46, zero_y - 10), "0.0", fill="#" + p["muted_text"], font=label_font)
    draw.text((30, bottom - 12), "-0.6", fill="#" + p["muted_text"], font=label_font)
    draw.text((left, bottom + 24), "2020", fill="#" + p["muted_text"], font=label_font)
    draw.text((530, bottom + 24), "Mar 2022", fill="#" + p["muted_text"], font=label_font)
    draw.text((break_x - 34, bottom + 24), "Sep 2022", fill="#" + p["secondary"],
              font=label_font)
    draw.text((1190, bottom + 24), "2025", fill="#" + p["muted_text"], font=label_font)

    shade = _mix_hex(p["background"], p["accent"], 0.10)
    draw.rectangle((break_x, top, right, bottom), fill="#" + shade)
    draw.line((break_x, top, break_x, bottom), fill="#" + p["secondary"], width=4)
    draw.text((break_x + 14, top + 10), "Post-break regime", fill="#" + p["secondary"],
              font=small_font)

    points = [
        (left, 255), (245, 270), (385, 220), (530, 285), (650, 320),
        (745, 342), (805, 395), (910, 470), (1020, 448), (1150, 505),
        (1260, 480), (right, 530),
    ]
    draw.line(points, fill="#" + p["primary"], width=7, joint="curve")
    for x, y in points:
        draw.ellipse((x - 7, y - 7, x + 7, y + 7), fill="#" + p["accent"],
                     outline="#" + p["primary"], width=2)

    image.save(CHART_PATH)


def _mix_hex(start, end, ratio):
    a = start.lstrip("#")
    b = end.lstrip("#")
    rgb = [
        round(int(a[i:i + 2], 16) * (1 - ratio) + int(b[i:i + 2], 16) * ratio)
        for i in (0, 2, 4)
    ]
    return "".join(f"{value:02X}" for value in rgb)


def build_deck():
    PROJECT.ensure()
    build_correlation_chart()
    b = PptxBuilder(palette=BLACKROCK)

    b.title_slide(
        title="Treasury Spread Model:\nResults & Recommendations",
        subtitle="AFE Portfolio Risk Modelling | May 2026",
    )

    numbers_slide(
        title="Four Signals Support a Model Change",
        stats=[
            {"number": "46pp", "label": "Bills PMSE improvement", "context": "60.2% to 14.1%"},
            {"number": "10", "label": "Active factors in Model B", "context": "Significant on 54%+ of days"},
            {"number": "5-fold", "label": "Cross-validation passes", "context": "No OOS overfitting detected"},
            {"number": "Sep 2022", "label": "Regime break", "context": "Bill-bond correlation flips"},
        ],
        builder=b,
    )

    process_slide(
        title="How the Model Works",
        steps=[
            {"name": "Curve KRD Return", "description": "Remove curve return using key rate durations."},
            {"name": "Excess Spread", "description": "Normalise residual return by spread duration."},
            {"name": "Daily Regression", "description": "Run OLS across USD Treasury instruments."},
            {"name": "Factor Extraction", "description": "Extract maturity factors and type dummies."},
            {"name": "PMSE Evaluation", "description": "Evaluate error in total return space."},
        ],
        highlight=3,
        footnote="Estimated daily. Sample Jan 2020-Jun 2025. Five-fold out-of-sample validation confirms stability.",
        builder=b,
    )

    results_slide(
        title="Model Performance vs Current: Model B Wins Across Cohorts",
        columns=["Cohort", "Current Model", "Model A", "Model B"],
        rows=[
            ["Govt Bonds", "3.3%", "1.1%", "0.4%"],
            ["Treasury Bills (largest gain)", "60.2%", "22.4%", "14.1%"],
            ["P Strips", "1.2%", "0.8%", "0.7%"],
            ["C Strips", "2.1%", "1.5%", "1.3%"],
            ["OTR All", "3.9%", "1.4%", "0.5%"],
            ["OTR Bills", "40.0%", "18.2%", "11.6%"],
            ["Short End <1Y", "45.4%", "28.1%", "21.6%"],
        ],
        win_col=3,
        worst_col=1,
        builder=b,
    )

    waterfall_slide(
        title="What Drives the Bills PMSE Improvement",
        items=[
            {"name": "Maturity buckets added", "value": 28.4},
            {"name": "Bill type factor added", "value": 14.2},
            {"name": "Strip type factor added", "value": 1.8},
            {"name": "Overfitting correction", "value": -0.5},
            {"name": "OOS adjustment", "value": -0.2},
        ],
        total_label="Net PMSE reduction (pp)",
        builder=b,
    )

    heat_map_slide(
        title="Factor Significance Rate by Cohort (% of Trading Days)",
        row_labels=["Govt Bonds", "Treasury Bills", "P Strips", "C Strips", "Short End <1Y"],
        col_labels=["<2m", "2-3m", "1-3y", "3-7y", "7-15y", ">15y", "Bill Type", "Strip Type"],
        values=[
            [0.54, 0.58, 0.76, 0.71, 0.68, 0.62, 0.00, 0.00],
            [0.81, 0.74, 0.00, 0.00, 0.00, 0.00, 0.60, 0.00],
            [0.00, 0.00, 0.55, 0.61, 0.64, 0.58, 0.00, 0.70],
            [0.00, 0.00, 0.57, 0.63, 0.66, 0.59, 0.00, 0.70],
            [0.82, 0.76, 0.54, 0.00, 0.00, 0.00, 0.60, 0.00],
        ],
        builder=b,
    )

    annotation_chart_slide(
        title="Bill-Bond Correlation: Structural Break at September 2022",
        chart_path=CHART_PATH,
        annotations=[
            {"label": "Pre-2022: move together", "x": 0.20, "y": 0.25},
            {"label": "Fed hiking cycle begins", "x": 0.45, "y": 0.55},
            {"label": "Correlation flips negative", "x": 0.55, "y": 0.75},
            {"label": "RRP amplifies demand", "x": 0.72, "y": 0.65},
        ],
        headline="Bill type factor captures behaviour maturity factors cannot",
        builder=b,
    )

    assertion_evidence_slide(
        title="Key Finding",
        assertion="Bills drive nearly all the PMSE improvement - for structural reasons",
        evidence=[
            {"text": "Bills PMSE falls from 60.2% to 14.1% under Model B", "stat": "46pp"},
            {"text": "Bill type factor is significant on 60% of trading days", "stat": "60%"},
            {"text": "Near-zero correlation with IR curve factors: genuinely orthogonal", "stat": "r ~ 0"},
            {"text": "September 2022 break confirms a structural driver", "stat": "Sep 22"},
        ],
        builder=b,
    )

    two_by_two_slide(
        title="Model Decisions: Effort vs Payoff",
        x_label="Implementation Complexity", y_label="PMSE Improvement",
        x_low="Simple", x_high="Complex", y_low="Marginal", y_high="Large",
        items=[
            {"name": "Maturity buckets", "x": 0.30, "y": 0.85, "descriptor": "High payoff"},
            {"name": "Bill type factor", "x": 0.35, "y": 0.90, "descriptor": "Largest gain"},
            {"name": "Strip dummies", "x": 0.40, "y": 0.25, "descriptor": "Marginal"},
            {"name": "Unit exposure", "x": 0.50, "y": 0.70, "descriptor": "Next test"},
            {"name": "GARCH-DCC", "x": 0.85, "y": 0.65, "descriptor": "Future"},
            {"name": "Sub-period split", "x": 0.45, "y": 0.40, "descriptor": "Diagnostic"},
        ],
        builder=b,
    )

    scorecard_slide(
        title="Model Selection: Current vs A vs B vs C",
        criteria=["PMSE improvement", "Factor parsimony", "OOS stability",
                  "Interpretability", "Implementation cost"],
        options=["Current", "Model A", "Model B", "Model C"],
        scores=[[1, 3, 5, 5], [5, 4, 4, 2], [3, 4, 5, 4],
                [5, 4, 4, 3], [5, 4, 4, 2]],
        weights=[0.35, 0.15, 0.25, 0.10, 0.15],
        builder=b,
    )

    diverging_bar_slide(
        title="PMSE Before vs After: Model B by Cohort",
        row_labels=["Govt Bonds", "Treasury Bills", "P Strips", "C Strips",
                    "OTR All", "OTR Bills", "Short End <1Y"],
        left_label="Current Model", right_label="Model B",
        left_values=[3.3, 60.2, 1.2, 2.1, 3.9, 40.0, 45.4],
        right_values=[0.4, 14.1, 0.7, 1.3, 0.5, 11.6, 21.6],
        builder=b,
    )

    status_slide(
        title="Next Steps & Ownership",
        workstreams=[
            {"name": "Unit Exposure Experiment", "status": "Compare raw vs normalised returns for sub-3m bills", "rag": "green", "owner": "A"},
            {"name": "Sub-period PMSE Split", "status": "Decompose improvement pre/post June 2022", "rag": "green", "owner": "A"},
            {"name": "Data Refresh", "status": "Extend sample through the current period", "rag": "amber", "owner": "A"},
            {"name": "BIS Self-Index", "status": "Obtain and integrate benchmark index", "rag": "amber", "owner": "S"},
            {"name": "Final Model Proposal", "status": "Complete cohort and index-level backtesting", "rag": "amber", "owner": "S"},
            {"name": "RQA Sign-off", "status": "Present findings and obtain ratification", "rag": "red", "owner": "S"},
        ],
        builder=b,
    )

    timeline_slide(
        title="Delivery Schedule",
        milestones=[
            {"date": "May 26", "label": "Unit exposure", "state": "past"},
            {"date": "May 26", "label": "Sub-period split", "state": "current"},
            {"date": "Jun 26", "label": "Data refresh", "state": "future"},
            {"date": "Jun 26", "label": "BIS index", "state": "future"},
            {"date": "Jul 26", "label": "Model proposal", "state": "future"},
            {"date": "Aug 26", "label": "RQA ratification", "state": "future"},
            {"date": "Sep 26", "label": "Production approval", "state": "future"},
        ],
        builder=b,
    )

    qa_path = PROJECT.qa_path(DECK_PATH.name)
    b.save(DECK_PATH, final=True, report_path=qa_path)
    print(f"Saved: {DECK_PATH} ({len(b.prs.slides)} slides)")
    print(f"Quality report: {qa_path}")


if __name__ == "__main__":
    build_deck()
