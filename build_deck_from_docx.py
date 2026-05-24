"""Build a checked starter presentation from a Word source brief.

The command provides a dependable first pass for chat assistants. For important
decks, the assistant should still refine wording, ordering and visual choice.
"""

from __future__ import annotations

import argparse
import json
import textwrap
from pathlib import Path

from create_pptx import PptxBuilder
from patterns import heat_map_slide, results_slide
from project_workspace import ProjectWorkspace
from source_docx import read_source_docx, source_to_markdown


_NARRATIVE_CHARS = 110
_NARRATIVE_ITEMS_PER_SLIDE = 5
_TABLE_ROWS_PER_SLIDE = 8
_TABLE_COLUMNS_PER_SLIDE = 8


def _split_text(text: str, max_chars: int = _NARRATIVE_CHARS) -> list[str]:
    """Wrap narrative into slide-safe chunks without dropping source words."""
    return textwrap.wrap(
        text,
        width=max_chars,
        break_long_words=False,
        break_on_hyphens=False,
    ) or [text]


def _chunks(items: list, size: int):
    for start in range(0, len(items), size):
        yield items[start:start + size]


def build_deck_from_docx(
    source_path,
    output_path=None,
    palette="blackrock",
    project_name=None,
    projects_root="projects",
    workspace=None,
    return_details=False,
) -> Path | dict:
    source_file = Path(source_path)
    if workspace is not None and project_name:
        raise ValueError("Supply either workspace or project_name, not both.")
    if workspace is None and project_name:
        workspace = ProjectWorkspace(project_name, root=projects_root)
    if workspace is not None:
        source_file = workspace.ingest_input(source_file)
        output = Path(output_path) if output_path else workspace.output_path(f"{source_file.stem}_deck.pptx")
        report = workspace.qa_path(output.name)
    else:
        output = Path(output_path) if output_path else Path("output") / f"{source_file.stem}_deck.pptx"
        report = output.with_suffix(".qa.json")
    source = read_source_docx(source_file)
    working_files = []
    if workspace is not None:
        structured_path = workspace.working_path(f"{source_file.stem}_source.json")
        markdown_path = workspace.working_path(f"{source_file.stem}_source.md")
        structured_path.write_text(json.dumps(source, indent=2), encoding="utf-8")
        markdown_path.write_text(source_to_markdown(source), encoding="utf-8")
        working_files.extend([structured_path, markdown_path])

    b = PptxBuilder(palette=palette)
    b.title_slide(source["title"], subtitle=f"Prepared from {source_file.name}")
    patterns_used = ["title_slide"]

    content_slides = 0
    for section in source["sections"]:
        paragraphs = section["paragraphs"]
        if not paragraphs:
            continue
        bullets = [chunk for paragraph in paragraphs for chunk in _split_text(paragraph)]
        pages = list(_chunks(bullets, _NARRATIVE_ITEMS_PER_SLIDE))
        for page_no, page in enumerate(pages, start=1):
            suffix = f" ({page_no}/{len(pages)})" if len(pages) > 1 else ""
            b.content_slide(section["heading"] + suffix, page)
            content_slides += 1
            patterns_used.append("content_slide")

    for label, texts in source.get("supplemental_text", {}).items():
        if not texts:
            continue
        bullets = [chunk for text in texts for chunk in _split_text(text)]
        pages = list(_chunks(bullets, _NARRATIVE_ITEMS_PER_SLIDE))
        heading = f"Captured Source {label.replace('_', ' ').title()}"
        for page_no, page in enumerate(pages, start=1):
            suffix = f" ({page_no}/{len(pages)})" if len(pages) > 1 else ""
            b.content_slide(heading + suffix, page)
            content_slides += 1
            patterns_used.append("content_slide")

    for table_no, table in enumerate(source["tables"], start=1):
        title = f"Source Data: Table {table_no}"
        numeric_headers = table["numeric_columns"]
        row_label_index = _first_label_column(table["headers"], numeric_headers)
        full_numeric_matrix = (
            numeric_headers
            and row_label_index == 0
            and len(numeric_headers) == len(table["headers"]) - 1
            and len(table["headers"]) <= _TABLE_COLUMNS_PER_SLIDE
            and len(table["rows"]) <= _TABLE_ROWS_PER_SLIDE
        )
        if full_numeric_matrix and table["rows"]:
            numeric_indices = [table["headers"].index(header) for header in numeric_headers]
            values = _numeric_matrix(table["rows"], numeric_indices)
            if values is not None:
                row_labels = [row[row_label_index] for row in table["rows"]]
                heat_map_slide(title, row_labels, numeric_headers, values, builder=b)
                patterns_used.append("heat_map_slide")
                continue
        if table["rows"]:
            for page_title, headers, rows in _table_pages(title, table["headers"], table["rows"]):
                results_slide(page_title, headers, rows, builder=b)
                patterns_used.append("results_slide")

    if content_slides == 0 and not source["tables"]:
        b.content_slide("Source Material", ["No usable narrative or table content was detected."])
        patterns_used.append("content_slide")

    b.save(output, final=True, report_path=report)
    details = {
        "source": source_file,
        "output": output,
        "report": report,
        "working_files": working_files,
        "patterns": patterns_used,
        "palette": palette,
    }
    return details if return_details else output


def _first_label_column(headers, numeric_headers):
    for index, header in enumerate(headers):
        if header not in numeric_headers:
            return index
    return None


def _parse_numeric(value):
    cleaned = (
        str(value).replace(",", "")
        .replace("%", "")
        .replace("$", "")
        .replace("(", "-")
        .replace(")", "")
        .strip()
    )
    number = float(cleaned)
    return number / 100 if "%" in str(value) else number


def _numeric_matrix(rows, numeric_indices):
    try:
        return [[_parse_numeric(row[index]) for index in numeric_indices] for row in rows]
    except (IndexError, TypeError, ValueError):
        return None


def _table_pages(title, headers, rows):
    """Yield readable table page slices that collectively preserve every cell."""
    if not headers:
        return
    label_index = 0
    other_indices = list(range(1, len(headers)))
    column_groups = list(_chunks(other_indices, _TABLE_COLUMNS_PER_SLIDE - 1)) or [[]]
    row_groups = list(_chunks(rows, _TABLE_ROWS_PER_SLIDE)) or [[]]
    total_pages = len(column_groups) * len(row_groups)
    page_no = 0
    for column_group in column_groups:
        indices = [label_index] + column_group
        for row_group in row_groups:
            page_no += 1
            suffix = f" ({page_no}/{total_pages})" if total_pages > 1 else ""
            yield (
                title + suffix,
                [headers[index] for index in indices],
                [[row[index] if index < len(row) else "" for index in indices] for row in row_group],
            )


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a checked PPTX starter deck from a DOCX brief.")
    parser.add_argument("source_docx")
    parser.add_argument("--output", dest="output_path")
    parser.add_argument("--palette", default="blackrock")
    parser.add_argument("--project", dest="project_name", help="Store inputs, outputs and QA under projects/<name>/")
    parser.add_argument("--projects-root", default="projects")
    args = parser.parse_args()
    output = build_deck_from_docx(
        args.source_docx,
        args.output_path,
        palette=args.palette,
        project_name=args.project_name,
        projects_root=args.projects_root,
    )
    print(f"Saved: {output}")
    if args.project_name:
        report = ProjectWorkspace(args.project_name, root=args.projects_root, create=False).qa_path(output.name)
    else:
        report = output.with_suffix(".qa.json")
    print(f"Quality report: {report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
