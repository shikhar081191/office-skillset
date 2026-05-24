"""Project-oriented paths for Office artifact generation workflows."""

from __future__ import annotations

import argparse
import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path


def project_slug(name: str) -> str:
    """Create a stable folder-safe slug from a project name."""
    slug = re.sub(r"[^a-z0-9]+", "_", name.strip().lower()).strip("_")
    if not slug:
        raise ValueError("Project name must contain letters or numbers.")
    return slug


class ProjectWorkspace:
    """Manage a predictable folder structure for one piece of team work."""

    def __init__(self, name: str, root: str | Path = "projects", create: bool = True):
        self.name = name
        self.slug = project_slug(name)
        self.base_dir = Path(root) / self.slug
        self.inputs_dir = self.base_dir / "inputs"
        self.assets_dir = self.base_dir / "assets"
        self.working_dir = self.base_dir / "working"
        self.outputs_dir = self.base_dir / "outputs"
        self.qa_dir = self.base_dir / "qa"
        if create:
            self.ensure()

    def ensure(self) -> "ProjectWorkspace":
        for directory in [
            self.inputs_dir,
            self.assets_dir,
            self.working_dir,
            self.outputs_dir,
            self.qa_dir,
        ]:
            directory.mkdir(parents=True, exist_ok=True)
        self._write_manifest()
        return self

    def input_path(self, filename: str) -> Path:
        return self.inputs_dir / filename

    def asset_path(self, filename: str) -> Path:
        return self.assets_dir / filename

    def working_path(self, filename: str) -> Path:
        return self.working_dir / filename

    def output_path(self, filename: str) -> Path:
        return self.outputs_dir / filename

    def qa_path(self, artifact_filename: str) -> Path:
        return self.qa_dir / f"{Path(artifact_filename).stem}.qa.json"

    def ingest_input(self, source_path: str | Path) -> Path:
        """Copy an external source into the project input folder if needed."""
        source = Path(source_path)
        if not source.exists():
            raise FileNotFoundError(f"Project input not found: {source}")
        destination = self.input_path(source.name)
        if source.resolve() != destination.resolve():
            shutil.copy2(source, destination)
        return destination

    def record_workflow(
        self,
        workflow: str,
        *,
        inputs=(),
        outputs=(),
        working_files=(),
        qa_reports=(),
        palette=None,
        patterns=(),
        settings=None,
    ) -> dict:
        """Append one reproducible generation run to the project manifest."""
        manifest = self.base_dir / "project.json"
        data = self._read_manifest()
        run = {
            "workflow": workflow,
            "run_utc": datetime.now(timezone.utc).isoformat(),
            "inputs": [self._manifest_path(path) for path in inputs],
            "outputs": [self._manifest_path(path) for path in outputs],
            "working_files": [self._manifest_path(path) for path in working_files],
            "qa_reports": [self._manifest_path(path) for path in qa_reports],
            "patterns": list(patterns),
        }
        if palette is not None:
            run["palette"] = palette if isinstance(palette, str) else "custom"
        if settings:
            run["settings"] = settings
        data.setdefault("workflow_runs", []).append(run)
        manifest.write_text(json.dumps(data, indent=2), encoding="utf-8")
        return run

    def register_output(
        self,
        output_path: str | Path,
        qa_report_path: str | Path | None = None,
        patterns: tuple | list = (),
        notes: str | None = None,
        working_files: tuple | list = (),
        qa_report_paths: tuple | list = (),
    ) -> dict:
        """Record a manually-built refined deck in the project manifest.

        Use this at the end of hand-authored build scripts (e.g. build_q2_model_review.py)
        so the output appears alongside workflow-generated decks in project.json.
        Regenerating the same named refined output replaces its earlier registration
        because QA repair iterations overwrite the same deliverable.
        """
        manifest_output = self._manifest_path(output_path)
        qa_reports = [self._manifest_path(qa_report_path)] if qa_report_path else []
        qa_reports.extend(self._manifest_path(path) for path in qa_report_paths)
        run = {
            "workflow": "manual_refinement",
            "run_utc": datetime.now(timezone.utc).isoformat(),
            "inputs": [],
            "outputs": [manifest_output],
            "working_files": [self._manifest_path(path) for path in working_files],
            "qa_reports": qa_reports,
            "patterns": list(patterns),
        }
        if notes:
            run["notes"] = notes
        manifest = self.base_dir / "project.json"
        data = self._read_manifest()
        workflow_runs = data.setdefault("workflow_runs", [])
        data["workflow_runs"] = [
            existing for existing in workflow_runs
            if not (
                existing.get("workflow") == "manual_refinement"
                and manifest_output in existing.get("outputs", [])
            )
        ]
        data["workflow_runs"].append(run)
        manifest.write_text(json.dumps(data, indent=2), encoding="utf-8")
        return run

    def _write_manifest(self) -> None:
        manifest = self.base_dir / "project.json"
        existing = self._read_manifest()
        data = {
            "name": existing.get("name", self.name),
            "slug": self.slug,
            "created_utc": existing.get("created_utc", datetime.now(timezone.utc).isoformat()),
            "folders": {
                "inputs": "inputs",
                "assets": "assets",
                "working": "working",
                "outputs": "outputs",
                "qa": "qa",
            },
            "workflow_runs": existing.get("workflow_runs", []),
        }
        manifest.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def _read_manifest(self) -> dict:
        manifest = self.base_dir / "project.json"
        if manifest.exists():
            try:
                return json.loads(manifest.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                pass
        return {
            "name": self.name,
            "slug": self.slug,
            "created_utc": datetime.now(timezone.utc).isoformat(),
            "folders": {
                "inputs": "inputs",
                "assets": "assets",
                "working": "working",
                "outputs": "outputs",
                "qa": "qa",
            },
            "workflow_runs": [],
        }

    def _manifest_path(self, path: str | Path) -> str:
        item = Path(path)
        try:
            return str(item.resolve().relative_to(self.base_dir.resolve()))
        except ValueError:
            return str(item)


def main() -> int:
    parser = argparse.ArgumentParser(description="Initialize a project workspace.")
    parser.add_argument("project_name")
    parser.add_argument("--root", default="projects")
    args = parser.parse_args()
    workspace = ProjectWorkspace(args.project_name, root=args.root)
    print(f"Created project workspace: {workspace.base_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
