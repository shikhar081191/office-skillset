"""Reusable Excel workbook patterns for research and model validation."""

from create_xlsx import FMT_2DP, FMT_PCT, XlsxBuilder


def factor_summary_workbook(title, factor_rows, assumptions=None):
    """Create a factor-return workbook with inputs, formulas and a summary chart."""
    b = XlsxBuilder()
    summary = b.add_sheet("Factor Summary")
    summary.label_cell(1, 1, title, bold=True)
    summary.data_table(
        start_row=3,
        start_col=1,
        headers=["Factor", "Return", "Volatility", "Sharpe"],
        rows=[
            [row["factor"], row["return"], row["volatility"], f"=B{index}/C{index}"]
            for index, row in enumerate(factor_rows, start=4)
        ],
        col_widths=[24, 14, 14, 14],
    )
    for row in range(4, 4 + len(factor_rows)):
        summary.ws.cell(row=row, column=2).number_format = FMT_PCT
        summary.ws.cell(row=row, column=3).number_format = FMT_PCT
        summary.ws.cell(row=row, column=4).number_format = FMT_2DP
    summary.add_bar_chart(
        title="Factor Returns",
        data_min_row=3,
        data_max_row=3 + len(factor_rows),
        data_min_col=2,
        data_max_col=2,
        categories_col=1,
        anchor_cell="F3",
    )
    summary.freeze_panes("A4")

    if assumptions:
        sheet = b.add_sheet("Assumptions")
        sheet.headers(1, ["Assumption", "Value"], col_width=28)
        for row_no, item in enumerate(assumptions, start=2):
            sheet.label_cell(row_no, 1, item["name"])
            sheet.assumption_cell(row_no, 2, item["value"], item.get("number_format"))
        sheet.col_width(2, 18)
    return b


def pmse_comparison_workbook(title, cohorts, models, values):
    """Create a PMSE comparison sheet and calculated best-model summary."""
    if len(values) != len(cohorts) or any(len(row) != len(models) for row in values):
        raise ValueError("values must match cohorts x models")
    b = XlsxBuilder()
    comparison = b.add_sheet("PMSE Comparison")
    comparison.label_cell(1, 1, title, bold=True)
    rows = []
    for cohort, row_values in zip(cohorts, values):
        best_index = min(range(len(row_values)), key=row_values.__getitem__)
        rows.append([cohort, *row_values, models[best_index]])
    comparison.data_table(
        start_row=3,
        start_col=1,
        headers=["Cohort", *models, "Lowest PMSE"],
        rows=rows,
        col_widths=[24] + [15] * len(models) + [18],
    )
    for row in range(4, 4 + len(cohorts)):
        for col in range(2, 2 + len(models)):
            comparison.ws.cell(row=row, column=col).number_format = FMT_PCT
    comparison.add_bar_chart(
        title="PMSE by Model",
        data_min_row=3,
        data_max_row=3 + len(cohorts),
        data_min_col=2,
        data_max_col=1 + len(models),
        categories_col=1,
        anchor_cell=f"{chr(67 + len(models))}3",
    )
    comparison.freeze_panes("B4")
    return b

