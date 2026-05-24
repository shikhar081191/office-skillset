# Remaining Plan

## Where The Toolkit Is Today

The core PowerPoint workflow is usable through an AI chat assistant:

- Users can supply notes in chat or provide a Word brief containing narrative
  and tables.
- Word intake preserves source context, creates an internal scaffold and keeps
  the original input inside a project workspace.
- The AI is instructed to read the full pattern catalog, create a slide plan
  and build a refined visual deck rather than deliver copied source text.
- The library includes research-oriented slide patterns for comparisons,
  findings, heat maps, waterfalls, timelines, scorecards, prioritisation and
  decision slides.
- BlackRock is the default palette and alternate/custom palettes can be used.
- Final refined PPTX builds run structural QA, including overlap, crowding,
  contrast, density and repeated-panel alignment checks.
- Final refined PPTX builds record rendered-review status; actual preview
  rendering runs when its optional runtime is installed.
- A reusable Q2 model-review project demonstrates the intended workflow.

Word and Excel builders, QA modules and retained advanced utilities are present,
but their reusable pattern coverage is less developed than PowerPoint.

## What Is Not The Goal Right Now

The current workflow uses GitHub Copilot or Windsurf as the reasoning layer.
The Python package does not independently call an embedded language model to
turn arbitrary documents into complete editorial deliverables. That can be
considered later, but it is not required for the team to use the current
chat-driven workflow.

## Recommended Next Priorities

### 1. Make Repeated Use Simpler

Objective: make each new real project feel predictable for a non-technical
user.

- Add a clean reusable project starter/example that does not contain completed
  outputs.
- Define one standard prompt for PowerPoint, one for Word documents and one
  for Excel workbooks.
- Add a small revision workflow guide for requests such as "make the
  recommendation stronger" or "turn this table into a chart".
- Decide whether project outputs and QA evidence should be committed for
  examples only, or retained locally for real confidential work.

### 2. Strengthen PowerPoint Delivery

Objective: move from strong scripted decks to repeatable team delivery.

- Install and document the optional rendering runtime on team machines so
  rendered slide previews and overflow checks execute rather than record a
  skipped status.
- Add further real-world sample decks across research use cases, such as
  factor attribution, stress testing and portfolio exposure review.
- Continue adding layout QA only where it can be detected reliably without
  generating noisy false warnings.
- Review whether a generic refined-deck build template can reduce duplicated
  project scripts while preserving AI pattern choice.

### 3. Expand Word Document Workflows

Objective: support committee notes, whitepapers and template-driven reports
with the same ease as slides.

- Add reusable DOCX patterns for short committee notes, validation reports and
  whitepapers.
- Provide a template-filling workflow for a user-supplied Word template.
- Integrate available advanced utilities for comments, tracked changes,
  accessibility, privacy scrubbing, captions and cross-references where they
  fit common team use cases.
- Define a final Word delivery loop that includes rendered page review when
  the optional runtime is configured.

### 4. Expand Excel Workflows

Objective: turn supplied analysis data into clean, editable research
workbooks and chart sources.

- Add reusable workbook patterns for scenario shocks, exposure heat maps,
  validation trackers and model-comparison scorecards.
- Improve the workflow for reading user-supplied Excel inputs and selecting
  appropriate summary tabs and charts.
- Keep financial formatting conventions and formula QA mandatory in final
  workbooks.

### 5. Evolve The Code Structure Carefully

Objective: keep the repository easy to maintain as utilities increase.

- Preserve current top-level imports and scripts so existing prompts and
  examples continue to work.
- When growth warrants it, move internal implementation into an organized
  package structure such as `office_utils/pptx/`, `office_utils/docx/`,
  `office_utils/xlsx/` and `office_utils/shared/`.
- Keep retained supplemental utilities available through
  `artifact_utilities.py`; do not remove capabilities during restructuring.

## Suggested Sequence

| Phase | Focus | Outcome |
|---|---|---|
| Current release | Publish the working PPTX-first toolkit and documentation | Team can begin controlled Copilot/Windsurf trials |
| Next iteration | Install rendered QA runtime and create reusable project starters | More reliable visual delivery and easier onboarding |
| Following iteration | Add Word report/template patterns | Committee notes and whitepapers become repeatable |
| Later iteration | Add richer Excel model/reporting patterns | Complete PPTX/DOCX/XLSX research workflow |
| As needed | Internal package restructuring | Cleaner maintainability without breaking user prompts |

## Success Criteria

The toolkit is mature for routine team use when:

- A non-technical user can provide source material and a plain-language prompt
  and receive a refined deliverable without touching Python.
- The AI consistently creates and follows an explicit plan before building.
- Presentation, document and workbook outputs each have appropriate QA and
  reusable design patterns.
- Optional rendered inspections run on standard team setups.
- Project folders remain clean, reproducible and safe to share.

