# Pattern: Results Comparison

## When to use
Use when you need to compare 2–4 models, scenarios, or options across a set of
metrics. The first column is row labels; subsequent columns are the things being
compared. Winning and losing columns get a subtle colour highlight. Ideal for:
model backtesting, scenario analysis, vendor evaluation, A/B test results.

## Function signature
```python
from patterns import results_slide

results_slide(
    title: str,               # slide title (shown in primary-colour header bar)
    columns: list[str],       # column headers including label col, e.g. ["Metric", "A", "B"]
    rows: list[list],         # data rows, each matching len(columns)
    win_col: int = None,      # 0-based column index to highlight green (optional)
    worst_col: int = None,    # 0-based column index to highlight amber (optional)
    builder=None,             # pass existing PptxBuilder to append; else creates new
    palette=None,             # palette dict (default: BLACKROCK)
) -> PptxBuilder
```

## Example call
```python
from patterns import results_slide

b = results_slide(
    title="Factor Model Comparison — Backtesting Results",
    columns=["Metric", "Baseline", "Factor A", "Factor B", "Delta vs Base"],
    rows=[
        ["Annualised Return", "8.2%",  "9.4%",  "11.1%", "+2.9%"],
        ["Volatility",        "14.1%", "13.8%", "12.6%", "-1.5%"],
        ["Sharpe Ratio",      "0.58",  "0.68",  "0.88",  "+0.30"],
        ["Max Drawdown",      "-18%",  "-15%",  "-11%",  "+7%"],
        ["Tracking Error",    "n/a",   "2.1%",  "2.4%",  "—"],
    ],
    win_col=3,     # "Factor B" column gets green highlight
    worst_col=1,   # "Baseline" column gets amber highlight
)
from project_workspace import ProjectWorkspace
project = ProjectWorkspace("Model Comparison")
deck_path = project.output_path("model_comparison.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```

## Content-shaping rules
1. **Column 0 is always labels**: metric names, plain text, left-aligned.
2. **Data columns**: values should be consistently formatted across a row (all % or all $,
   not mixed). Use "n/a" or "—" for missing data — never leave cells blank.
3. **win_col / worst_col**: these are 0-based indices into the `columns` list.
   - `win_col=2` highlights the third column (index 2) in green.
   - If there is no clear winner/loser, omit both arguments.
4. **Key metric rows**: numeric cells are automatically rendered larger and bold.
   Keep values as strings (e.g. "11.1%") — the pattern detects and sizes them.
5. **Delta column**: if the user wants a comparison-to-baseline column, include it as
   the last data column. Show deltas as "+X%" or "-X%" with sign.
6. **Row count**: 4–8 rows is ideal. More than 10 rows compress too much.
7. **Title**: include what is being compared and the context
   (e.g. "Q2 Scenario Analysis — Equity Portfolio").
