"""Append-friendly Markdown decision log for Office toolkit workflows.

Created automatically by artifact_workflow.py and deck.py.
AI build scripts extend it with slide plan decisions before writing any code.

Example (in an AI build script):
    from decision_log import DecisionLog
    from project_workspace import ProjectWorkspace

    project = ProjectWorkspace("Model Review", create=False)
    log = DecisionLog(project.working_path("decision_log.md"))
    log.record_slide_plan([
        {"section": "Introduction",       "pattern": "recommendation_slide", "assertion": "Approve model v3 for Q3"},
        {"section": "Model Performance",  "pattern": "kpi_dashboard_slide",  "assertion": "PMSE improved 18% on holdout"},
        {"section": "Stress Tests",       "pattern": "assumption_table_slide","assertion": "All parameters within regulatory bounds"},
        {"section": "Implementation",     "pattern": "timeline_slide",        "assertion": "Go-live Q3; two milestones at risk"},
    ])
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


class DecisionLog:
    """Append-friendly Markdown log. Call create() to start a fresh log;
    instantiate directly to append to an existing one."""

    def __init__(self, path: Path | str):
        self.path = Path(path)

    # ── creation ──────────────────────────────────────────────────────────────

    @classmethod
    def create(
        cls,
        path: Path | str,
        source_path: Path | str | None = None,
        audience: str | None = None,
        decision: str | None = None,
    ) -> "DecisionLog":
        """Write a fresh log with a header. Overwrites any existing file."""
        log = cls(path)
        log.path.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            "# AI Decision Log",
            "",
            f"_Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}_",
            "",
            "Records decisions taken during deck construction — "
            "source analysis, pattern selection, and QA results.",
            "",
        ]
        if source_path or audience or decision:
            lines += ["## Run context", ""]
            if source_path:
                lines.append(f"- **Source**: `{Path(source_path).name}`")
            if audience:
                lines.append(f"- **Audience**: {audience}")
            if decision:
                lines.append(f"- **Decision sought**: {decision}")
            lines.append("")
        log.path.write_text("\n".join(lines), encoding="utf-8")
        return log

    # ── append helpers ────────────────────────────────────────────────────────

    def _append(self, text: str) -> None:
        with self.path.open("a", encoding="utf-8") as f:
            f.write(text)

    def record_source_analysis(self, source: dict) -> "DecisionLog":
        """Log what was found in the source document.

        source: the dict returned by read_source_docx()
        """
        sections = source.get("sections", [])
        top_tables = source.get("tables", [])
        distil = source.get("sections_needing_distillation", 0)
        word_count = sum(
            len(" ".join(s.get("paragraphs", [])).split())
            for s in sections
        )
        lines = ["\n---\n", "## Source analysis\n"]
        lines.append(f"- **Sections found**: {len(sections)}")
        for s in sections:
            heading = s.get("heading") or "(untitled)"
            p_count = len(s.get("paragraphs", []))
            t_count = len(s.get("table_ids", s.get("tables", [])))
            tag = f" `[DISTIL]`" if s.get("distillation_needed") else ""
            detail = f"{p_count} paragraph(s)"
            if t_count:
                detail += f", {t_count} table(s)"
            lines.append(f"  - **{heading}** — {detail}{tag}")
        lines += [
            f"- **Tables at document level**: {len(top_tables)}",
            f"- **Approximate word count**: {word_count}",
        ]
        if distil:
            lines.append(
                f"- **⚠ Sections needing distillation**: {distil} "
                "(marked `[DISTIL]` — rewrite into concise slide copy, do not copy verbatim)"
            )
        lines.append("")
        self._append("\n".join(lines))
        return self

    def record_slide_plan(self, plan: list[dict], overwrite: bool = False) -> "DecisionLog":
        """Record the AI's pattern selection as a standalone plan and log section.

        plan: list of dicts — section, pattern and assertion are required;
            source, rationale, inputs and inference make the decision auditable.
        overwrite: if False (default), preserve an existing standalone plan and
            avoid appending a duplicate plan section to the decision log.

        Call this BEFORE writing the build script so the user can see the
        reasoning without reading code.
        """
        plan_text = self._slide_plan_text(plan)
        plan_path = self.path.with_name("slide_plan.md")
        if overwrite or not plan_path.exists():
            plan_path.write_text(plan_text, encoding="utf-8")
        if not overwrite:
            try:
                existing = self.path.read_text(encoding="utf-8")
                if "## Slide plan" in existing or "# Slide Plan" in existing:
                    return self
            except OSError:
                pass
        self._append("\n---\n\n" + plan_text)
        return self

    @staticmethod
    def _slide_plan_text(plan: list[dict]) -> str:
        lines = [
            "# Slide Plan",
            "",
            "Prepared by the AI assistant after considering the available pattern catalog.",
            "",
            "| # | Source evidence | Assertion | Selected pattern | Why selected | Inputs / assets | AI inference? |",
            "|---|---|---|---|---|---|---|",
        ]
        for index, row in enumerate(plan, start=1):
            source = _table_value(row.get("source", row.get("section", "")))
            assertion = _table_value(row.get("assertion", ""))
            pattern = _table_value(row.get("pattern", ""))
            rationale = _table_value(row.get("rationale", ""))
            inputs = _table_value(row.get("inputs", ""))
            inference = _table_value(row.get("inference", "No"))
            lines.append(
                f"| {index} | {source} | {assertion} | `{pattern}` | "
                f"{rationale} | {inputs} | {inference} |"
            )
        lines.append("")
        return "\n".join(lines)

    def record_qa(self, qa_report: dict | Path | str) -> "DecisionLog":
        """Append QA gate results.

        qa_report: either the report dict or a path to the JSON report file.
        """
        if isinstance(qa_report, (str, Path)):
            path = Path(qa_report)
            if path.exists():
                qa_report = json.loads(path.read_text(encoding="utf-8"))
            else:
                self._append(f"\n> QA report not found at `{path}`\n")
                return self

        errors = qa_report.get("errors", [])
        warnings = qa_report.get("warnings", [])
        passed = qa_report.get("passed", False)

        if not passed:
            status = "**FAILED**"
        elif warnings:
            status = "**PASSED** (with warnings)"
        else:
            status = "**PASSED**"

        lines = ["\n---\n", "## QA results\n"]
        lines.append(
            f"**Status**: {status} — "
            f"{qa_report.get('slide_count', 0)} slides, "
            f"{len(errors)} error(s), {len(warnings)} warning(s)\n"
        )
        for issue in errors:
            ref = f"[slide {issue['slide']}] " if "slide" in issue else ""
            lines.append(f"- ERROR {ref}{issue['message']}")
        for issue in warnings:
            ref = f"[slide {issue['slide']}] " if "slide" in issue else ""
            lines.append(f"- WARN  {ref}{issue['message']}")
        if not errors and not warnings:
            lines.append("- All checks clean.")
        lines.append("")
        self._append("\n".join(lines))
        return self

    def record_output(self, output_path: Path | str) -> "DecisionLog":
        """Log the final output file path."""
        lines = ["\n---\n", "## Output\n", f"`{output_path}`\n"]
        self._append("\n".join(lines))
        return self

    def note(self, heading: str, text: str) -> "DecisionLog":
        """Append a free-form note under a sub-heading."""
        self._append(f"\n### {heading}\n\n{text}\n")
        return self


def _table_value(value) -> str:
    return str(value or "").replace("|", r"\|").replace("\n", " ")
