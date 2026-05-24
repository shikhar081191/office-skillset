"""Build the EM fixed income tail risk review deck."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from create_pptx import PptxBuilder
from patterns import BLACKROCK, chart_context_slide, numbers_slide, process_slide, results_slide
from project_workspace import ProjectWorkspace

PROJECT = ProjectWorkspace("EM FI Tail Risk Review")
CHART_PATH = PROJECT.asset_path("em_fi_divergence_chart.png")
DECK_PATH = PROJECT.output_path("em_fi_tail_risk_review.pptx")


def _font(size: int) -> ImageFont.ImageFont:
    try:
        return ImageFont.truetype("arial.ttf", size)
    except OSError:
        return ImageFont.load_default()


def build_divergence_chart() -> None:
    PROJECT.ensure()
    p = BLACKROCK
    width, height = 1400, 860
    image = Image.new("RGB", (width, height), "#" + p["background"])
    draw = ImageDraw.Draw(image)
    title_font = _font(28)
    label_font = _font(18)
    value_font = _font(16)
    legend_font = _font(16)

    data = [
        ("Brazil", -4.2, -6.8),
        ("Mexico", -2.8, -3.7),
        ("India", -2.1, -2.2),
        ("Indonesia", -3.0, -3.9),
        ("South Africa", -3.9, -5.9),
        ("Turkey", -5.1, -7.4),
        ("Poland", -1.8, -1.9),
        ("Chile", -2.4, -2.9),
    ]
    max_length = max(abs(old) for _, old, _ in data) * 1.15

    chart_left = 280
    chart_top = 120
    chart_width = 1040
    chart_height = 620
    row_height = chart_height / len(data)
    bar_base_x = chart_left + 220
    bar_max_width = chart_width - 280

    draw.text((40, 25), "Pilot divergence: old vs updated correlation stress outputs",
              fill="#" + p["primary"], font=title_font)
    draw.text((40, 70), "Values are pilot book P&L for a taper-style shock, post-2022 correlations.",
              fill="#" + p["muted_text"], font=label_font)

    draw.rectangle((chart_left, chart_top, chart_left + chart_width, chart_top + chart_height),
                   outline="#" + p["border"], width=2)

    for idx, (country, old_value, new_value) in enumerate(data):
        y = chart_top + idx * row_height + 20
        draw.text((40, y + 6), country, fill="#" + p["text_dark"], font=label_font)

        old_width = int((abs(old_value) / max_length) * bar_max_width)
        new_width = int((abs(new_value) / max_length) * bar_max_width)
        row_y = y + 10
        bar_height = 32

        draw.rectangle(
            (bar_base_x, row_y, bar_base_x + old_width, row_y + bar_height),
            fill="#" + p["secondary"], outline=None)
        draw.rectangle(
            (bar_base_x, row_y + bar_height + 10, bar_base_x + new_width, row_y + bar_height * 2 + 10),
            fill="#" + p["accent"], outline=None)

        draw.text((bar_base_x + bar_max_width + 18, row_y - 2), f"{old_value:.1f}%",
                  fill="#" + p["text_dark"], font=value_font)
        draw.text((bar_base_x + bar_max_width + 18, row_y + bar_height + 8), f"{new_value:.1f}%",
                  fill="#" + p["text_dark"], font=value_font)

    legend_y = chart_top + chart_height + 10
    draw.rectangle((bar_base_x, legend_y, bar_base_x + 32, legend_y + 20), fill="#" + p["secondary"])
    draw.text((bar_base_x + 40, legend_y), "Old correlation propagation", fill="#" + p["text_dark"], font=legend_font)
    draw.rectangle((bar_base_x + 420, legend_y, bar_base_x + 452, legend_y + 20), fill="#" + p["accent"])
    draw.text((bar_base_x + 472, legend_y), "Updated post-2022 correlation propagation", fill="#" + p["text_dark"], font=legend_font)

    image.save(CHART_PATH)


def build_deck() -> None:
    PROJECT.ensure()
    build_divergence_chart()
    b = PptxBuilder(palette="blackrock")

    b.title_slide(
        title="EM Fixed Income Tail Risk Review",
        subtitle="Stress testing post-2022 regime transmission | May 2026",
    )

    numbers_slide(
        title="Why this matters",
        stats=[
            {
                "number": "62%",
                "label": "Understated dollar-sensitive EM tail risk",
                "context": "Taper-style shock on updated correlations"
            },
            {
                "number": "-4.2% vs -6.8%",
                "label": "Old vs updated EM book output",
                "context": "Same scenario, different transmission"
            },
            {
                "number": "8",
                "label": "Countries in pilot",
                "context": "Brazil, Mexico, India, Indonesia, South Africa, Turkey, Poland, Chile"
            },
            {
                "number": "Post-2022",
                "label": "Regime break",
                "context": "Current strong dollar, high rates environment"
            },
        ],
        builder=b,
    )

    chart_context_slide(
        title="Divergence by country",
        chart_path=CHART_PATH,
        headline="Dollar-sensitive names show the largest additional risk once correlations update",
        bullets=[
            "Brazil, South Africa and Turkey move from materially negative to much more stressed.",
            "India and Poland remain largely stable, revealing genuine defensive exposures.",
            "A current-regime propagation highlights where the book is truly resilient.",
        ],
        so_what="The current historical shock framework can still be used—its transmission needs to reflect today’s EM FX-USD regime.",
        builder=b,
    )

    process_slide(
        title="Methodology in plain language",
        steps=[
            {
                "name": "Keep the story",
                "description": "Preserve the scenario narrative: taper shock, USD strength, EM outflows."
            },
            {
                "name": "Update the transmission",
                "description": "Recalibrate factor correlations using post-2022 data, not the original vintage regime."
            },
            {
                "name": "Run the same shock",
                "description": "Propagate the historical scenario through current correlation structure."
            },
            {
                "name": "Compare and allocate",
                "description": "Use the result to separate genuinely defensive exposures from stale apparent resilience."
            },
        ],
        highlight=1,
        footnote="This keeps the scenario content intact while ensuring the model reflects today’s market relationships.",
        builder=b,
    )

    results_slide(
        title="What we recommend",
        columns=["Action", "Why it matters"],
        rows=[
            [
                "Keep the scenario set",
                "Historical shocks remain valid; the transmission mechanism is the thing to refresh."
            ],
            [
                "Update correlation propagation quarterly",
                "Ensure the same shock reflects today's USD/EM regime, not the regime in which the scenario occurred."
            ],
            [
                "Prioritise dollar-sensitive EM",
                "Brazil, South Africa and Turkey show the biggest additional tail-risk once correlations update."
            ],
            [
                "Use India and Poland to identify genuine defence",
                "These countries are the most resilient under the current regime, so allocation can be more selective."
            ],
        ],
        win_col=None,
        worst_col=None,
        builder=b,
    )

    b.save(DECK_PATH, final=True, report_path=PROJECT.qa_path(DECK_PATH.name))
    print(f"Saved deck to {DECK_PATH}")


if __name__ == "__main__":
    build_deck()
