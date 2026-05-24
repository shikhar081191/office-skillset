"""Single entry point for project-scoped Office artifact generation.

The first supported automated route turns a source DOCX brief into a checked
PowerPoint deck. Other output routes can be added behind the same contract
without changing how a non-technical user or an AI assistant invokes it.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
from pathlib import Path

from build_deck_from_docx import build_deck_from_docx
from decision_log import DecisionLog
from project_workspace import ProjectWorkspace
from source_docx import read_source_docx


SUPPORTED_ROUTES = {("docx", "pptx")}


def run_workflow(
    source_path: str | Path,
    *,
    output_type: str = "pptx",
    project_name: str | None = None,
    projects_root: str | Path = "projects",
    output_name: str | None = None,
    palette="blackrock",
    instructions: str | None = None,
    audience: str | None = None,
    decision: str | None = None,
) -> dict:
    """Generate a checked artifact and record the complete project run."""
    source = Path(source_path)
    if not source.exists():
        raise FileNotFoundError(f"Workflow input not found: {source}")
    input_type = source.suffix.lstrip(".").lower()
    output_type = output_type.lower().lstrip(".")
    if (input_type, output_type) not in SUPPORTED_ROUTES:
        raise ValueError(
            f"Unsupported workflow route: {input_type or 'unknown'} -> {output_type}. "
            "The automated route currently available is DOCX -> PPTX."
        )

    project = ProjectWorkspace(
        project_name or source.stem.replace("_", " ").title(),
        root=projects_root,
    )
    project_input = project.ingest_input(source)
    working_files = []

    # Prepend audience and decision context to instructions so the AI assistant
    # applies them throughout the editorial pass.
    context_lines = []
    if audience:
        context_lines.append(f"Audience: {audience}")
    if decision:
        context_lines.append(f"Decision sought: {decision}")
    full_instructions = "\n".join(context_lines + ([instructions] if instructions else []))

    if full_instructions.strip():
        instruction_path = project.working_path("user_instructions.txt")
        instruction_path.write_text(full_instructions.strip() + "\n", encoding="utf-8")
        working_files.append(instruction_path)

    # Create decision log before any build work so the AI can append to it
    log_path = project.working_path("decision_log.md")
    log = DecisionLog.create(
        log_path,
        source_path=source,
        audience=audience,
        decision=decision,
    )
    source_info = read_source_docx(project_input)
    log.record_source_analysis(source_info)
    working_files.append(log_path)

    output_path = project.output_path(output_name or f"{project_input.stem}_deck.pptx")
    result = build_deck_from_docx(
        project_input,
        output_path=output_path,
        palette=palette,
        workspace=project,
        return_details=True,
    )
    working_files.extend(result["working_files"])

    # Append QA results and output path to the decision log
    log.record_qa(result["report"])
    log.record_output(result["output"])

    run = project.record_workflow(
        "docx_to_pptx",
        inputs=[project_input],
        outputs=[result["output"]],
        working_files=working_files,
        qa_reports=[result["report"]],
        palette=palette,
        patterns=result["patterns"],
        settings={
            "input_type": input_type,
            "output_type": output_type,
            "instructions_supplied": bool(instructions),
            "visual_qa": "structural",
        },
    )
    return {
        "project": project.base_dir,
        "input": project_input,
        "output": result["output"],
        "qa_report": result["report"],
        "decision_log": log_path,
        "working_files": working_files,
        "patterns": result["patterns"],
        "manifest": project.base_dir / "project.json",
        "run": run,
    }


def _json_safe(result: dict) -> dict:
    return {key: _json_value(value) for key, value in result.items()}


def _json_value(value):
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {key: _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a checked project-scoped Office artifact from supplied source material."
    )
    parser.add_argument("--source", required=True, help="Source document; DOCX -> PPTX is currently supported.")
    parser.add_argument("--output-type", default="pptx", choices=["pptx"])
    parser.add_argument("--project", dest="project_name", help="Project folder name; defaults to source filename.")
    parser.add_argument("--projects-root", default="projects")
    parser.add_argument("--output-name", help="Optional final artifact filename inside outputs/.")
    parser.add_argument("--palette", default="blackrock")
    parser.add_argument("--instructions", help="Optional user direction recorded with the workflow run.")
    parser.add_argument("--audience", help="Who will see this deck (e.g. 'risk committee', 'board').")
    parser.add_argument("--decision", help="The decision or action this deck should drive.")
    parser.add_argument("--json", action="store_true", dest="as_json", help="Print machine-readable results.")
    args = parser.parse_args()
    if args.as_json:
        with contextlib.redirect_stdout(io.StringIO()):
            result = run_workflow(
                args.source,
                output_type=args.output_type,
                project_name=args.project_name,
                projects_root=args.projects_root,
                output_name=args.output_name,
                palette=args.palette,
                instructions=args.instructions,
                audience=args.audience,
                decision=args.decision,
            )
        print(json.dumps(_json_safe(result), indent=2))
    else:
        result = run_workflow(
            args.source,
            output_type=args.output_type,
            project_name=args.project_name,
            projects_root=args.projects_root,
            output_name=args.output_name,
            palette=args.palette,
            instructions=args.instructions,
            audience=args.audience,
            decision=args.decision,
        )
        print(f"Saved:         {result['output']}")
        print(f"Decision log:  {result['decision_log']}")
        print(f"QA report:     {result['qa_report']}")
        print(f"Manifest:      {result['manifest']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
