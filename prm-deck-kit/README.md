> **Reference material only — not part of this repository's workflow.**
> This folder contains the original Claude chat toolkit that inspired the
> Python patterns in this repo. It uses **PptxGenJS (JavaScript)** and
> references paths such as `/mnt/skills/public/pptx/` that only exist in
> the Claude chat sandbox environment. None of it runs in this repo.
> The authoritative workflow for this project is in `AI_INSTRUCTIONS.md`,
> `STORY_TEMPLATES.md`, and `patterns.py`. Read those instead.

---

# PPTX Toolkit — Portfolio Risk Modelling (Reference)

A portable kit for instructing an AI to build presentation decks
in the style used by the BlackRock Portfolio Risk Modelling team.

---

## What's in this kit

| File              | Purpose |
|-------------------|---------|
| `INSTRUCTIONS.md` | Step-by-step process the AI follows: parse brief → plan slides → code → QA → deliver |
| `STYLE_GUIDE.md`  | Visual standards: palette, fonts, layout patterns, tone rules, common mistakes |

The underlying build infrastructure (PptxGenJS API reference and environment
scripts) lives separately at `/mnt/skills/public/pptx/`. The AI reads those
automatically when it reads `SKILL.md`.

---

## How to use this

### Give it to an AI in a new conversation

Paste or attach both files and say:

> "Read INSTRUCTIONS.md and STYLE_GUIDE.md. Then build me a deck from the following brief: [your brief]"

The AI will read the skill files, plan the slides, write the code, run visual
QA, and deliver the `.pptx`.

### What to include in your brief

You don't need to be precise. Just write naturally. The AI will extract:
- Audience (who's in the room)
- Topics to cover (what you want on each slide)
- Any hard constraints (slide count, things to avoid, specific names/dates)

If you don't specify slide count, it defaults to 8.
If you don't specify audience, it defaults to "senior internal, full jargon."

### Example brief format

```
Build me a deck for [audience] covering [topics].
Max [N] slides.

[Write the content in plain prose — no need to pre-structure it by slide.
The AI will figure out the slide breakdown.]
```

---

## What the AI produces

- A `.pptx` file at `/mnt/user-data/outputs/`
- Navy/teal palette, Calibri font, clean card-based layouts
- Visual QA already done (text checked for overflow/overlap)
- A short summary of what's in the deck and any assumptions made

---

## Customising the style

Edit `STYLE_GUIDE.md` to change:
- **Colors** — replace the hex values in the palette table
- **Layouts** — add new layout pattern descriptions
- **Tone** — update the language rules section
- **Slide sequence** — update the structural patterns section

The AI treats this file as a hard constraint document. Changes here
propagate to every deck built with this kit.

---

## Limitations

- The toolkit produces decks that look best in PowerPoint (Microsoft 365).
  LibreOffice rendering is used for QA but may show minor font differences.
- Images and icons require internet access or local file paths.
  Emoji characters render as OS-default glyphs, which may vary.
- The style guide is calibrated for internal risk/finance presentations.
  It may need adjustment for client-facing or marketing decks.
