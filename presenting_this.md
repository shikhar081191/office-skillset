# Why Use This Toolkit Rather Than Claude Chat Directly?

This document answers the question a sceptical colleague will ask: *"Why not
just open Claude chat, paste my notes, and ask it to make a PowerPoint?"*

The short answer is: you can, and sometimes that is the right move. But this
toolkit adds seven things that a raw chat session cannot give you — and the
tradeoffs are real.

---

## What You Get Here That You Do Not Get From a Raw Chat Prompt

### 1. It works inside your corporate environment

Claude chat is blocked in many financial institutions. GitHub Copilot and
Windsurf are approved in far more corporate networks because they route
through enterprise licensing agreements. The entire value of this toolkit
is that it brings the same quality of AI-assisted deck building to the tools
your IT department has already signed off.

### 2. Twenty-three tested, reusable slide patterns

Claude chat will invent a layout each time. The shapes and coordinates might
be fine, or they might overflow the slide. There is no institutional memory.

This toolkit has 23 patterns — `numbered_steps_slide`, `heat_map_slide`,
`callout_bar_slide` and so on — that have been tested, QA'd and calibrated for
a 13.33" × 7.5" slide canvas. When you ask for a process flow you get the same
reliable geometry every time, not a fresh guess.

### 3. Story templates ensure the narrative is structured, not improvised

`STORY_TEMPLATES.md` defines five named narrative sequences (Approval Request,
Project Explainer, Exec Update, Research Findings, Status Update). Each one
specifies which slide types appear in which order, what the opening principle
is, and which patterns fit each slot.

Claude chat will make a reasonable attempt at structure. But it does not have
a committed opening principle (e.g. *"lead with the problem, not the solution"*
for a Project Explainer) unless you think to specify it in the prompt. Here,
that discipline is baked into the instructions the AI reads before it plans
the deck.

### 4. A mandatory QA gate runs automatically on every build

`pptx_qa.py` checks every saved file for:

- Objects placed off the visible canvas
- Placeholder text left unfilled
- Text boxes likely to overflow their bounds
- Low-contrast text (WCAG AA threshold: 4.5:1 for body text)
- Overlapping shapes
- Sparse or overly dense slides
- Uneven sizing or spacing in repeated panels

Claude chat has no equivalent. It may generate a script that produces a slide
with white text on a white background and you would not know until you opened
the file. Here, the QA report flags it and the AI is instructed to fix it
before delivery.

### 5. Every build leaves a full audit trail

Claude chat gives you a file. This toolkit gives you:

| File | What it records |
|---|---|
| `working/slide_plan.md` | Which pattern was chosen for each slide and why |
| `working/decision_log.md` | How the story was structured, what was inferred, what was taken directly from source |
| `qa/<name>.qa.json` | Structured QA results: pass/warn/fail per check |
| `qa/<name>.rendered_qa.json` | Whether a visual rendering review ran or why it was skipped |
| `scripts/build_<name>.py` | The exact Python script that produced the deck — re-runnable |

For a governance-sensitive environment this matters. If a committee asks "how
was that scenario analysis framed?", the answer is in `decision_log.md`, not
in someone's memory of what they typed into a chat box.

### 6. Outputs are reproducible and shareable

The build script that the AI writes is a real Python file. You can re-run it,
hand it to a colleague, modify the data and rebuild, or check it into version
control. The deck is not a black-box artefact from a conversation no one else
can see.

Claude chat produces a one-shot output. When the context window clears or the
session ends, the reasoning that produced the deck is gone.

### 7. A shared team library, versioned and improvable

When a new palette is added, a new pattern is built or a QA rule is tightened,
every person using the repo benefits immediately. Claude chat's quality depends
on what any individual types into their prompt.

---

## What This Toolkit Does Not Do (Limitations You Should Know)

### It requires setup

A Python virtual environment, `pip install -r requirements.txt`, and access to
GitHub Copilot or Windsurf. Claude chat has no setup — open a browser and go.

For a one-off, quick slide Claude chat wins on convenience. This toolkit earns
its overhead when you are doing this regularly or sharing work with a team.

### Rendered visual inspection is optional and not installed by default

The QA gate checks structure mathematically. It cannot open the file and look
at it the way a human would. Full rendered inspection (slide PNGs, overflow
smoke test) requires LibreOffice and Poppler to be installed separately. If
they are not, the system records that fact but you will not get slide previews
automatically.

Claude chat cannot render slides either. But it will not flag this gap — this
toolkit at least tells you the check was skipped.

### The AI reasoning still depends on prompt quality

The instructions in `AI_INSTRUCTIONS.md` and `.github/copilot-instructions.md`
guide Copilot and Windsurf, but they do not guarantee perfect output. A vague
brief will still produce a vague deck. The story templates and patterns reduce
the variance, but they do not eliminate it.

The toolkit makes it easier to give a good brief (see `BRIEF_TEMPLATE.md`) and
catches structural errors automatically. The editorial judgement — whether the
story is right for the audience, whether the emphasis is correct — still belongs
to you.

### Pattern coverage is finite

Twenty-three patterns cover the most common research and executive communication
needs. But there are layouts that fall outside them. Claude chat can produce
arbitrary custom layouts because it has no pre-committed geometry. This toolkit
will use `content_slide` as a fallback, which is a plain layout with less visual
impact.

### It will not help you with the hard judgment calls

No toolkit will. The speed gain comes from eliminating mechanical construction
work — choosing shapes, setting coordinates, running layout checks. The
intellectual work — what the scenario implies, whether the assumptions are
defensible, what the committee needs to hear — is yours.

---

## Summary

| Question | Claude chat | This toolkit |
|---|---|---|
| Works inside corporate network? | Often blocked | Yes (via Copilot/Windsurf) |
| Consistent tested slide geometry? | No — recreated each time | Yes — 23 patterns |
| Story template with opening principle? | Improvised | 5 named templates |
| Automatic QA on every build? | No | Yes — structural + contrast |
| Audit trail and decision log? | No | Yes — working/ and qa/ folders |
| Reproducible build script? | No | Yes — scripts/ folder |
| Rendered visual preview? | No | Optional (needs runtime) |
| Zero setup needed? | Yes | No — Python + Copilot/Windsurf |
| Arbitrary custom layouts? | Yes | Limited to 23 patterns |

Use Claude chat for a quick one-off you will not need to explain to anyone.
Use this toolkit when the output needs to be consistent, auditable, revisable
by a colleague, or good enough for a committee.
