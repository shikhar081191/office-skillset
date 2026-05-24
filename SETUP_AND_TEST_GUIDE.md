# Setup & Testing Guide — office_tools_mini (Python-only)

No npm, no Node.js, no LibreOffice needed for the core builders, pattern demos,
or mandatory PowerPoint quality gate, including repeated-panel alignment
warnings. Final refined decks record rendered-review status; generating visual
previews and rendered overflow tests requires optional LibreOffice/Poppler
runtimes.

For a short explanation written for non-technical users, start with
[`EXPLANATION.md`](EXPLANATION.md).
For recommended prompts and the expected user journey, read
[`BEST_PRACTICES.md`](BEST_PRACTICES.md).

---

## Step 1: Unzip and open in VS Code

Unzip `office_scripts_mini.zip` anywhere, e.g.:
```
C:\Users\<you>\projects\office_scripts_mini\
```

Open the folder in VS Code:
```
File → Open Folder → select office_scripts_mini
```

---

## Step 2: Install Python dependencies

Open a terminal (`Ctrl + ~`) and run:

```bash
pip install -r requirements.txt
```

This installs: `python-docx`, `python-pptx`, `openpyxl`, `pandas`, `Pillow`, `defusedxml`.
It also installs `lxml` for the integrated advanced DOCX utility layer.

If pip is not on PATH:
```bash
python -m pip install -r requirements.txt
```

That's it. No other installs needed.

---

## Step 3: Enable GitHub Copilot workspace instructions

1. Open VS Code Settings (`Ctrl + ,`)
2. Search: `github.copilot.chat.codeGeneration.useInstructionFiles`
3. Set to: `true`

Copilot will now follow the rules in `.github/copilot-instructions.md` for all
code it generates in this workspace.

---

## Step 4: Run the test suite

```bash
python test_office_tools_mini.py -v
```

**Expected output:**

```
============================================================
office_tools_mini test suite (Python-only)
============================================================

✓  All dependencies found

test_add_comment (TestComment) ... ok
test_comment_text_in_xml (TestComment) ... ok
test_custom_author (TestComment) ... ok
...
test_demo_runs (TestCreateXlsx) ... ok

----------------------------------------------------------------------
Ran 42 tests in 6.3s

OK
```

**Run a single class:**
```bash
python test_office_tools_mini.py TestCreateDocx
python test_office_tools_mini.py TestCreatePptx
python test_office_tools_mini.py TestCreateXlsx
python test_office_tools_mini.py TestUnpackPack
python test_office_tools_mini.py TestExtractText
python test_office_tools_mini.py TestComment
```

---

## Step 5: Run the demos (optional)

Each creator script has a built-in demo. Run them to verify output looks right:

```bash
python create_docx.py    # creates demo_output.docx
python create_pptx.py    # creates demo_output.pptx
python create_xlsx.py    # creates demo_output.xlsx
python patterns.py        # creates thirteen pattern demos in output/
```

PowerPoint, Word and Excel saves automatically run their lightweight structural
QA checks. Open the outputs in Word / PowerPoint / Excel when a human review is
needed.

For real work, final artifacts are placed under
`projects/<project_name>/outputs/`; top-level `output/` contains technical
pattern demos.

Initialize a project folder explicitly when useful:

```bash
python project_workspace.py "Treasury Spread Model"
```

Run the Word-brief intake stage used by the AI assistant:

```bash
python deck.py model_brief.docx --project "Model Brief" --no-open
```

The command copies the input into the project, writes extracted context to
`working/`, generates a checked internal scaffold in `outputs/`, stores its QA
report, and records the generation run in `project.json`. In normal use, the AI
assistant then reads `skills/PATTERN_CATALOG.md`, writes
`working/slide_plan.md` with its pattern-selection reasoning, builds a refined
deck, repairs QA findings, records rendered-review status and gives the user
the refined PPTX.

The person using Copilot or Windsurf generally does not need to type this
command. They can provide the Word file in chat and ask for a presentation;
the assistant should run the full intake, refinement, QA-repair and
rendered-review status loop.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `ModuleNotFoundError: No module named 'docx'` | `pip install python-docx` |
| `ModuleNotFoundError: No module named 'pptx'` | `pip install python-pptx` |
| `ModuleNotFoundError: No module named 'openpyxl'` | `pip install openpyxl` |
| `ModuleNotFoundError: No module named 'PIL'` | `pip install Pillow` |
| `ModuleNotFoundError: No module named 'defusedxml'` | `pip install defusedxml` |
| Test creates file but Word won't open it | Run `python test_office_tools_mini.py TestUnpackPack -v` and share error |
| Demo runs but output looks wrong | Open an issue — share the file |

---

## What each script does

| Script | Purpose | Use it when... |
|--------|---------|---------------|
| `create_docx.py` | Create Word docs with DocxBuilder | Building reports, memos, proposals |
| `docx_patterns.py` | Reusable Word document patterns | Research notes and validation memos |
| `docx_qa.py` | Automatic Word quality gate | Checking final DOCX structure and draft tokens |
| `create_pptx.py` | Create PowerPoint with PptxBuilder | Building decks, presentations |
| `create_xlsx.py` | Create Excel with XlsxBuilder | Building models, data tables |
| `xlsx_patterns.py` | Reusable research workbooks | PMSE/model comparison and factor summary |
| `xlsx_qa.py` | Automatic workbook quality gate | Checking final XLSX structure and formula faults |
| `source_docx.py` | Extract Word source context | Turning notes/tables into deck inputs |
| `build_deck_from_docx.py` | Checked starter-deck build | Converting a Word brief into an initial PPTX |
| `deck.py` | Friendly Word intake entry point | Starting a DOCX-to-refined-deck chat workflow |
| `artifact_workflow.py` | Intake/workspace controller | Producing extracted context and an internal PPTX scaffold |
| `project_workspace.py` | Project folder/manifest manager | Keeping inputs, outputs, assets, QA and run history together |
| `artifact_utilities.py` | Retained utility facade | Accessing DOCX/slides/spreadsheet advanced tools |
| `unpack.py` | Unzip any Office file → XML | You need to edit raw formatting/XML |
| `pack.py` | Rezip XML → Office file | After editing XML in unpacked/ |
| `extract_text.py` | Dump text from any Office file | Reading/analysing existing files |
| `comment.py` | Add tracked comments to DOCX | Code review, annotation workflows |
| `patterns.py` | Reusable research slide layouts | Deck storytelling and visual comparison |
| `pptx_qa.py` | Automatic PowerPoint quality gate | Blocking definite slide delivery defects |
| `artifact_python_utilities_chatgpt/` | Retained advanced utilities | Rendering, audits, redlines, optional artifact workflows |

---

## Folder structure

```
office_scripts_mini/
├── .github/
│   └── copilot-instructions.md   ← Copilot rules (auto-loaded by VS Code)
├── create_docx.py                ← DocxBuilder — create Word docs
├── docx_patterns.py              ← Reusable Word deliverable patterns
├── docx_qa.py                    ← Mandatory structural DOCX quality gate
├── create_pptx.py                ← PptxBuilder — create PowerPoint decks
├── pptx_qa.py                    ← Mandatory structural PPTX quality gate
├── create_xlsx.py                ← XlsxBuilder — create Excel workbooks
├── xlsx_patterns.py              ← Reusable research workbook patterns
├── xlsx_qa.py                    ← Mandatory structural XLSX quality gate
├── source_docx.py                ← Word brief/table intake for presentations
├── build_deck_from_docx.py       ← Checked starter deck from a Word brief
├── artifact_workflow.py          ← DOCX intake/scaffold workflow controller
├── deck.py                       ← Friendly AI intake command for Word files
├── project_workspace.py          ← Project-scoped paths and run manifest
├── artifact_utilities.py         ← Supported facade for all retained utilities
├── examples/                      ← Reusable build scenarios and sample inputs
├── projects/                     ← Real deliverable workspaces
├── unpack.py                     ← Unzip Office file → XML
├── pack.py                       ← Rezip XML → Office file
├── extract_text.py               ← Extract text from any Office file
├── comment.py                    ← Add comments to DOCX
├── requirements.txt              ← pip dependencies (pip install -r ...)
├── test_office_tools_mini.py     ← Full test suite
└── SETUP_AND_TEST_GUIDE.md       ← This file
```

---

## Using with Copilot — example prompts

Once setup is complete, open Copilot Chat (`Ctrl + Alt + I`):

**Word document:**
```
Using DocxBuilder from create_docx.py, create a risk report with:
- A title page
- An executive summary section with 3 bullet points
- A table showing factor returns with a blue header row
Save to risk_report.docx
```

**PowerPoint:**
```
Read my attached model_brief.docx, create a project workspace, interpret the
talking points and tables, create a concise PowerPoint deck using the best
patterns, interpret and fix QA findings, record rendered-review status, and
return only the refined final deck with its QA evidence.
```

**Excel:**
```
Using XlsxBuilder from create_xlsx.py, create a P&L model with:
- An Assumptions sheet with growth rate and margin inputs (yellow cells)
- A P&L sheet with revenue forecasts using Excel formulas
- Blue font for inputs, black for formulas
Follow the financial colour conventions in copilot-instructions.md
```
