"""Reusable Word document patterns for research-team deliverables."""

from create_docx import DocxBuilder


def research_note_document(
    title,
    subtitle="",
    executive_summary=None,
    findings=None,
    data_tables=None,
    methodology=None,
    next_steps=None,
    template=None,
):
    """Build a concise research note suitable for circulation or conversion to slides."""
    b = DocxBuilder(template=template)
    b.heading(title, level=1)
    if subtitle:
        b.paragraph(subtitle, italic=True, color="666666")
    b.horizontal_rule()

    if executive_summary:
        b.heading("Executive Summary", level=2)
        b.paragraph(executive_summary)

    if findings:
        b.heading("Key Findings", level=2)
        b.bullet_list(findings)

    for table in data_tables or []:
        b.heading(table.get("title", "Supporting Analysis"), level=2)
        if table.get("context"):
            b.paragraph(table["context"])
        b.table(
            table.get("data", []),
            header_row=True,
            col_widths=table.get("col_widths"),
        )

    if methodology:
        b.heading("Methodology", level=2)
        b.paragraph(methodology)

    if next_steps:
        b.heading("Next Steps", level=2)
        b.bullet_list(next_steps)
    return b


def model_validation_memo(
    title,
    model_name,
    recommendation,
    evidence,
    risks,
    actions,
    metadata=None,
    template=None,
):
    """Build a short governance-oriented memo for model review or sign-off."""
    b = DocxBuilder(template=template)
    b.heading(title, level=1)
    b.paragraph(f"Model: {model_name}", bold=True)
    for item in metadata or []:
        b.paragraph(str(item), italic=True, color="666666")
    b.horizontal_rule()

    b.heading("Recommendation", level=2)
    b.paragraph(recommendation, bold=True)
    b.heading("Supporting Evidence", level=2)
    b.bullet_list(evidence)
    b.heading("Risks And Limitations", level=2)
    b.bullet_list(risks)
    b.heading("Required Actions", level=2)
    b.numbered_list(actions)
    return b

