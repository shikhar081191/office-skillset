"""Supported access layer for the retained artifact Python utility bundle.

The supplemental utility source is preserved unchanged. This facade discovers
every executable utility, reports optional runtime requirements, and executes
tools in a project-aware way without forcing optional dependencies into the
lightweight Office builder installation.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

from project_workspace import ProjectWorkspace


BUNDLE_ROOT = Path("artifact_python_utilities_chatgpt") / "artifact_python_utilities"


@dataclass(frozen=True)
class UtilityCapability:
    tool_id: str
    category: str
    source_path: str
    purpose: str
    optional_requirements: tuple[str, ...] = ()
    kind: str = "command"


PURPOSES = {
    "docx.render_docx": "Render a DOCX to page images for visual inspection.",
    "docx.accept_tracked_changes": "Report, accept or reject tracked revisions.",
    "docx.add_tracked_replacements": "Insert tracked replacement edits.",
    "docx.a11y_audit": "Audit and optionally repair accessibility issues.",
    "docx.apply_template_styles": "Apply template styles and themes to a DOCX.",
    "docx.captions_and_crossrefs": "Add figure/table captions and cross-reference support.",
    "docx.comments_add": "Add anchored Word comments.",
    "docx.comments_apply_patch": "Edit or resolve existing Word comments.",
    "docx.comments_extract": "Extract comments to structured output.",
    "docx.comments_strip": "Remove comments for final delivery.",
    "docx.content_controls": "List, wrap or fill Word content controls.",
    "docx.docx_ooxml_patch": "Apply advanced OOXML edits.",
    "docx.docx_table_to_csv": "Export a Word table to CSV.",
    "docx.fields_materialize": "Materialize field display values for deterministic QA.",
    "docx.fields_report": "Report Word fields.",
    "docx.flatten_ref_fields": "Flatten reference fields to visible text.",
    "docx.footnotes_report": "Report footnotes and endnotes.",
    "docx.heading_audit": "Audit heading hierarchy.",
    "docx.images_audit": "Audit embedded images.",
    "docx.insert_note": "Insert document notes.",
    "docx.insert_ref_fields": "Insert REF fields for cross-references.",
    "docx.insert_toc": "Insert a Word table of contents.",
    "docx.internal_nav": "Add document navigation links.",
    "docx.make_fixtures": "Generate DOCX edge-case fixtures.",
    "docx.merge_docx_append": "Append one DOCX to another.",
    "docx.privacy_scrub": "Remove personal metadata and revision identifiers.",
    "docx.redact_docx": "Redact sensitive document text.",
    "docx.render_and_diff": "Render two DOCX files and compare pages.",
    "docx.section_audit": "Audit sections and page settings.",
    "docx.set_protection": "Set document editing protection.",
    "docx.style_lint": "Find inconsistent document styles.",
    "docx.style_normalize": "Apply conservative style cleanup.",
    "docx.watermark_add": "Add a detectable Word watermark.",
    "docx.watermark_audit_remove": "Find or remove Word watermarks.",
    "docx.xlsx_to_docx_table": "Convert an Excel worksheet into a Word table.",
    "slides.create_montage": "Create a slide/contact-sheet montage from images.",
    "slides.detect_font": "Detect missing or substituted presentation fonts.",
    "slides.ensure_raster_image": "Normalize supported visual assets to raster images.",
    "slides.render_slides": "Render PPTX slides to PNG images.",
    "slides.slides_test": "Run rendered slide overflow checks.",
    "spreadsheets.spreadsheet_artifact_tool_starter": "Artifact-tool spreadsheet starter and render workflow.",
}


def _tool_id(path: Path) -> tuple[str, str]:
    relative = path.relative_to(BUNDLE_ROOT)
    parts = relative.parts
    stem = path.stem
    if parts[0] == "docx":
        return "docx", f"docx.{stem}"
    if parts[0] == "slides" and "artifact_tool_examples" in parts:
        return "slides", f"slides.examples.{stem}"
    if parts[0] == "slides":
        return "slides", f"slides.{stem}"
    return "spreadsheets", f"spreadsheets.{stem}"


def _optional_requirements(tool_id: str, source: str) -> tuple[str, ...]:
    requirements = []
    if "artifact_tool" in source or "presentation_artifact_tool" in source:
        requirements.append("artifact_tool")
    if tool_id in {"docx.render_docx", "docx.render_and_diff", "slides.render_slides", "slides.slides_test"}:
        requirements.extend(["pdf2image", "soffice", "poppler"])
    if tool_id == "slides.detect_font":
        requirements.extend(["fontconfig", "soffice"])
    if tool_id == "slides.ensure_raster_image":
        requirements.append("external image converters for non-raster inputs")
    return tuple(dict.fromkeys(requirements))


def discover_capabilities(bundle_root: str | Path = BUNDLE_ROOT) -> dict[str, UtilityCapability]:
    """Discover every retained executable Python utility and example."""
    root = Path(bundle_root)
    capabilities = {}
    for path in sorted(root.rglob("*.py")):
        relative_path = path.relative_to(root)
        if path.name == "__init__.py":
            continue
        category, tool_id = _tool_id(path)
        source = path.read_text(encoding="utf-8", errors="replace")
        kind = "example" if ".examples." in tool_id else "command"
        capabilities[tool_id] = UtilityCapability(
            tool_id=tool_id,
            category=category,
            source_path=str(relative_path),
            purpose=PURPOSES.get(tool_id, f"Retained {category} utility: {path.stem}."),
            optional_requirements=_optional_requirements(tool_id, source),
            kind=kind,
        )
    return capabilities


def runtime_status(capability: UtilityCapability) -> dict[str, bool]:
    """Report whether optional Python/system dependencies are discoverable."""
    checks = {}
    for requirement in capability.optional_requirements:
        if requirement == "artifact_tool":
            checks[requirement] = bool(
                importlib.util.find_spec("artifact_tool")
                or importlib.util.find_spec("presentation_artifact_tool")
            )
        elif requirement == "pdf2image":
            checks[requirement] = importlib.util.find_spec("pdf2image") is not None
        elif requirement == "soffice":
            checks[requirement] = shutil.which("soffice") is not None
        elif requirement == "poppler":
            checks[requirement] = shutil.which("pdftoppm") is not None or shutil.which("pdfinfo") is not None
        elif requirement == "fontconfig":
            checks[requirement] = shutil.which("fc-list") is not None
        else:
            checks[requirement] = False
    return checks


class ArtifactUtilities:
    """Execute retained utilities while placing logs and outputs in a project."""

    def __init__(
        self,
        project: ProjectWorkspace | None = None,
        bundle_root: str | Path = BUNDLE_ROOT,
    ):
        self.project = project
        self.bundle_root = Path(bundle_root)
        self.capabilities = discover_capabilities(self.bundle_root)

    def list(self, category: str | None = None) -> list[dict]:
        records = []
        for capability in self.capabilities.values():
            if category and capability.category != category:
                continue
            record = asdict(capability)
            record["runtime_status"] = runtime_status(capability)
            records.append(record)
        return records

    def run(
        self,
        tool_id: str,
        args: list[str] | tuple[str, ...],
        log_name: str | None = None,
        check: bool = True,
    ) -> subprocess.CompletedProcess:
        if tool_id not in self.capabilities:
            raise KeyError(f"Unknown retained utility: {tool_id}")
        capability = self.capabilities[tool_id]
        script = (self.bundle_root / capability.source_path).resolve()
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            [sys.executable, str(script), *[str(arg) for arg in args]],
            cwd=str(script.parent),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        if self.project:
            target = self.project.qa_dir / (log_name or f"{tool_id.replace('.', '_')}.run.json")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(
                json.dumps(
                    {
                        "tool_id": tool_id,
                        "command": [str(script), *[str(arg) for arg in args]],
                        "returncode": result.returncode,
                        "stdout": result.stdout,
                        "stderr": result.stderr,
                        "runtime_status": runtime_status(capability),
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
        if check and result.returncode != 0:
            missing = [key for key, available in runtime_status(capability).items() if not available]
            hint = f" Missing optional runtime: {', '.join(missing)}." if missing else ""
            raise RuntimeError(
                f"Retained utility {tool_id} failed with exit code {result.returncode}.{hint}\n"
                f"{result.stderr or result.stdout}"
            )
        return result

    def export_docx_table(self, docx_path, output_name="table.csv", table_index=0) -> Path:
        output = self._working_output(output_name)
        self.run(
            "docx.docx_table_to_csv",
            [Path(docx_path).resolve(), "--table_index", table_index, "--out", output.resolve()],
        )
        return output

    def xlsx_table_to_docx(self, xlsx_path, output_name="table_report.docx", sheet=None, title=None) -> Path:
        output = self._output(output_name)
        args = [Path(xlsx_path).resolve(), "--out", output.resolve()]
        if sheet:
            args.extend(["--sheet", sheet])
        if title:
            args.extend(["--title", title])
        self.run("docx.xlsx_to_docx_table", args)
        return output

    def render_pptx(self, pptx_path, preview_folder="slides_rendered") -> Path:
        target = self._qa_folder(preview_folder)
        self.run("slides.render_slides", [Path(pptx_path).resolve(), "--output_dir", target.resolve()])
        return target

    def render_docx(self, docx_path, preview_folder="docx_rendered") -> Path:
        target = self._qa_folder(preview_folder)
        self.run("docx.render_docx", [Path(docx_path).resolve(), "--output_dir", target.resolve()])
        return target

    def create_montage(self, image_dir, output_name="slides_montage.png") -> Path:
        output = self._qa_output(output_name)
        self.run(
            "slides.create_montage",
            ["--input_dir", Path(image_dir).resolve(), "--output_file", output.resolve()],
        )
        return output

    def review_rendered_pptx(self, pptx_path) -> Path:
        """Render and smoke-test a final PPTX when the optional visual runtime is available."""
        source = Path(pptx_path)
        report_path = self._qa_output(f"{source.stem}.rendered_qa.json")
        needed_tools = ("slides.render_slides", "slides.slides_test")
        status = {
            tool_id: runtime_status(self.capabilities[tool_id])
            for tool_id in needed_tools
        }
        missing = sorted({
            requirement
            for values in status.values()
            for requirement, available in values.items()
            if not available
        })
        report = {
            "path": str(source),
            "status": "skipped" if missing else "pending",
            "runtime_status": status,
            "missing_requirements": missing,
        }
        if missing:
            report["message"] = (
                "Rendered visual QA was not run because the optional runtime is unavailable: "
                + ", ".join(missing)
                + ". Structural PPTX QA remains mandatory and has run separately."
            )
            report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
            return report_path

        preview_folder = f"{source.stem}_rendered"
        rendered_dir = self.render_pptx(source, preview_folder=preview_folder)
        overflow_result = self.run(
            "slides.slides_test",
            [source.resolve()],
            log_name=f"{source.stem}_slides_test.run.json",
            check=False,
        )
        montage_path = self.create_montage(
            rendered_dir,
            output_name=f"{source.stem}_montage.png",
        )
        report.update({
            "status": "review_required" if overflow_result.returncode == 0 else "failed",
            "rendered_slides": str(rendered_dir),
            "montage": str(montage_path),
            "rendered_overflow_returncode": overflow_result.returncode,
            "rendered_overflow_stdout": overflow_result.stdout,
            "rendered_overflow_stderr": overflow_result.stderr,
            "message": (
                "Rendered previews were created. The assistant must inspect individual slide PNGs "
                "for clipping, alignment, contrast and visual quality before delivery."
            ),
        })
        report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
        return report_path

    def _working_output(self, name) -> Path:
        if not self.project:
            return Path(name)
        return self.project.working_path(name)

    def _output(self, name) -> Path:
        if not self.project:
            return Path(name)
        return self.project.output_path(name)

    def _qa_output(self, name) -> Path:
        if not self.project:
            return Path(name)
        return self.project.qa_dir / name

    def _qa_folder(self, name) -> Path:
        folder = self._qa_output(name)
        folder.mkdir(parents=True, exist_ok=True)
        return folder


def main() -> int:
    parser = argparse.ArgumentParser(description="Access all retained artifact utility capabilities.")
    parser.add_argument("--project", help="Project workspace name used for run logs and output conventions")
    parser.add_argument("--projects-root", default="projects")
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser("list", help="List retained utilities and optional requirements")
    list_parser.add_argument("--category", choices=["docx", "slides", "spreadsheets"])
    list_parser.add_argument("--json", action="store_true")

    run_parser = subparsers.add_parser("run", help="Execute any retained Python utility by ID")
    run_parser.add_argument("tool_id")
    run_parser.add_argument("tool_args", nargs=argparse.REMAINDER)
    args = parser.parse_args()

    project = ProjectWorkspace(args.project, root=args.projects_root) if args.project else None
    tools = ArtifactUtilities(project=project)
    if args.command == "list":
        records = tools.list(args.category)
        if args.json:
            print(json.dumps(records, indent=2))
        else:
            for item in records:
                requirements = ", ".join(item["optional_requirements"]) or "core"
                print(f"{item['tool_id']}: {item['purpose']} [{requirements}]")
        return 0

    result = tools.run(args.tool_id, args.tool_args, check=False)
    if result.stdout:
        print(result.stdout, end="")
    if result.stderr:
        print(result.stderr, file=sys.stderr, end="")
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
