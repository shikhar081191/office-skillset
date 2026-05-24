# Workflow Skill: Unified Artifact Workflow

## When to use
Use when a user supplies source material and expects the system to manage the
project folders and ultimately deliver a refined presentation without exposing
Python mechanics. The automated code route currently handles Word intake to an
internal PPTX scaffold; the AI assistant performs refinement and QA repair.

## Supported route

```text
DOCX source -> extracted context -> internal scaffold PPTX -> read PATTERN_CATALOG -> working/slide_plan.md -> refined deck -> QA repair loop -> rendered-review status -> refined PPTX delivered
```

Run:

```bash
python artifact_workflow.py --source model_brief.docx --output-type pptx --project "Model Brief"
```

Use `--palette` when the user provides another brand palette and
`--instructions` to store a short editorial direction with the generation run.

## Output contract
The workflow creates or updates:

```text
projects/<project>/
  inputs/       copied source DOCX
  working/      extracted context, instructions, decision log and slide_plan.md
  outputs/      internal scaffold PPTX and delivered refined PPTX
  qa/           scaffold, refined PPTX and rendered-review status reports
  project.json  workflow run with palette and patterns used
```

The assistant must not deliver the scaffold as the finished editorial deck.
It must read `skills/PATTERN_CATALOG.md`, document pattern choices and reasons
in `working/slide_plan.md`, build the patterned refined deck, inspect the
resulting QA report, repair every actionable warning, rerun QA and register
the refined output. It must also call
`ArtifactUtilities(project).review_rendered_pptx(deck_path)`. That helper
generates previews and a rendered overflow test when the optional runtime is
configured, or records an explicit skipped status while structural QA remains
mandatory.
