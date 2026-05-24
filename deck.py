"""One-shot CLI entry point for checked deck scaffolding.

Usage:
    python deck.py brief.docx
    python deck.py brief.docx --project "Model Review" --palette midnight_executive
    python deck.py brief.docx --audience "risk committee" --decision "approve model v2"
    python deck.py brief.docx --instructions "Lead with the recommendation slide."

After structural QA passes the scaffold path is printed and the file is opened
automatically. An AI assistant then uses the extracted context and skills to
create the refined presentation.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from artifact_workflow import run_workflow


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a checked PowerPoint scaffold from a Word source document.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("source", help="Source .docx file")
    parser.add_argument("--project", dest="project_name",
                        help="Project folder name (defaults to source filename)")
    parser.add_argument("--palette", default="blackrock",
                        choices=["blackrock", "midnight_executive", "coral_energy",
                                 "teal_trust", "charcoal_minimal", "warm_terracotta"],
                        help="Colour palette (default: blackrock)")
    parser.add_argument("--audience",
                        help="Who will see this deck — shapes tone and emphasis")
    parser.add_argument("--decision",
                        help="The decision or action this deck should drive")
    parser.add_argument("--instructions",
                        help="Free-form editorial direction for the AI pass")
    parser.add_argument("--output-name", dest="output_name",
                        help="Override the output filename inside outputs/")
    parser.add_argument("--no-open", action="store_true",
                        help="Do not open the file after generation")
    args = parser.parse_args()

    source = Path(args.source)
    if not source.exists():
        print(f"Error: source file not found: {source}", file=sys.stderr)
        return 1

    print(f"Building checked scaffold from {source.name} …")
    try:
        result = run_workflow(
            source,
            project_name=args.project_name,
            palette=args.palette,
            audience=args.audience,
            decision=args.decision,
            instructions=args.instructions,
            output_name=args.output_name,
        )
    except RuntimeError as exc:
        print(f"\nQA gate failed: {exc}", file=sys.stderr)
        return 1

    output = Path(result["output"])
    log = result.get("decision_log")
    print(f"\nScaffold:     {output}")
    if log:
        print(f"Decision log: {log}")

    if not args.no_open:
        try:
            if sys.platform == "win32":
                subprocess.Popen(["start", "", str(output)], shell=True)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(output)])
            else:
                subprocess.Popen(["xdg-open", str(output)])
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
