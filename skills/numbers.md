# Pattern: Numbers

## When to use
Use when 3–4 headline statistics need to land with maximum visual impact. The dark
background and oversized numbers make this pattern ideal for opening slides,
executive summaries, or "portfolio at a glance" pages. Each stat has a number
(very large, accent colour), a label, and a one-line context note.

## Function signature
```python
from patterns import numbers_slide

numbers_slide(
    title: str,             # slide title — shown in text_light on the dark background
    stats: list[dict],      # 3–4 stat dicts (see below)
    builder=None,           # pass existing PptxBuilder to append; else creates new
    palette=None,           # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

Each stat dict:

| Key     | Type | Rules                                                  |
|---------|------|--------------------------------------------------------|
| number  | str  | The big number — short: "$2.4B", "0.91", "-11%", "94bps" |
| label   | str  | What the number measures — 2–4 words                  |
| context | str  | One-line qualifier — benchmark, trend, or limit, ≤ 40 chars |

## Example call
```python
from patterns import numbers_slide

b = numbers_slide(
    title="Portfolio at a Glance — 31 May 2025",
    stats=[
        {"number": "$2.4B", "label": "AUM",           "context": "+12% vs prior year"},
        {"number": "0.91",  "label": "Sharpe Ratio",  "context": "Above 0.80 target"},
        {"number": "-11%",  "label": "Max Drawdown",  "context": "Within -15% limit"},
        {"number": "94bps", "label": "Active Return", "context": "vs benchmark Q2 YTD"},
    ]
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Portfolio Glance")
deck_path = project.output_path("portfolio_glance.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. **Count**: 3 or 4 stats. 3 gives more breathing room; 4 fills the slide.
   Never use 2 (too sparse) or 5+ (numbers shrink, impact is lost).
2. **Number format**: keep it short so it renders large.
   - Currency: "$2.4B" not "$2,400,000,000"
   - Percentages: "-11%" not "-11.32%"
   - Ratios: "0.91" not "0.9143"
   - Basis points: "94bps" not "0.94%"
3. **Label**: what the number measures. Noun phrase, not a sentence.
   "Sharpe Ratio" not "The portfolio Sharpe Ratio for the quarter".
4. **Context**: a single short qualifier that tells the audience whether the number
   is good, bad, or on target. Reference a benchmark, limit, or prior period.
   "Above 0.80 target", "Within -15% limit", "+12% vs prior year".
5. **Title**: include the portfolio or entity name and date.
   "Q2 2025 at a Glance", "Fund X — 31 May 2025".
6. **Palette note**: the primary colour is used as the background. For BLACKROCK this
   is black, which gives the most visual impact. For light palettes, consider using
   a `content_slide` with `dark_bg=True` instead.
