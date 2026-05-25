# Instructions for AI: Building Presentation Decks

You are building `.pptx` presentation decks using PptxGenJS in a sandboxed
Linux environment. Follow these instructions precisely. They tell you *how*
to work, not just what to produce.

---

## Step 0: Always Read Skills First

Before writing any code, read these files in order:

1. `/mnt/skills/public/pptx/SKILL.md` — environment overview and QA process
2. `/mnt/skills/public/pptx/pptxgenjs.md` — full API reference
3. `STYLE_GUIDE.md` (this toolkit) — visual and tonal standards

Do not skip this. The skill files encode environment-specific constraints
(soffice paths, pdftoppm flags, shadow object mutation bugs) that are not
in your training data.

---

## Step 1: Parse the Brief

When given a deck brief, extract the following before writing a single line of code:

**Required:**
- What is the deck's purpose? (project explainer, exec update, stakeholder review, training)
- Who is the audience? (junior analyst, senior stakeholder, external client, team)
- How many slides maximum?
- What topics must be covered?

**Infer if not stated:**
- Audience expertise level → determines how much jargon to explain
- Tone → formal for exec/client, plain for internal/junior
- Slide count → default to the standard 8-slide sequence unless told otherwise

**If the brief is ambiguous on a critical point**, make a reasonable assumption
and state it at the top of your response. Do not ask for clarification before
attempting the deck — produce a draft and flag your assumptions.

---

## Step 2: Plan the Slide Sequence

Before coding, write out the planned slide list as a comment at the top of
your script. Example:

```
// Slides:
// 1. Title
// 2. The Problem
// 3. The Solution
// 4. How It Works (5-step pipeline)
// 5. Human in the Loop
// 6. The Team
// 7. Timeline
// 8. Why It Matters
```

Match this to the structural patterns in STYLE_GUIDE.md. For a project
explainer, use the standard 8-slide sequence. For other deck types, adapt.

---

## Step 3: Write the Code

### File setup

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
```

Output path: `/mnt/user-data/outputs/YourDeckName.pptx`

### Palette

Import and use the palette from STYLE_GUIDE.md as constants:

```javascript
const NAVY   = "0D1B2A";
const TEAL   = "0A9396";
const LTEAL  = "94D2BD";
const WHITE  = "FFFFFF";
const OFFWHT = "E9F5F5";
const GRAY   = "4A5568";
const LGRAY  = "E2E8F0";
const AMBER  = "EE9B00";
```

### Critical code rules

1. **Never use `#` in hex colors** — corrupts the file
2. **Never reuse option objects across addShape calls** — use a factory function:
   ```javascript
   const makeShadow = () => ({ type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.1 });
   ```
3. **Never use `lineSpacing` with bullets** — use `paraSpaceAfter` instead
4. **Set `margin:0`** on text boxes that must align with shape edges
5. **Keep description text width ≤ 6.2"** on process/step slides if there's a right-side callout
6. **Use `bullet: true`** in options — never unicode "•" characters

### Slide header pattern (use on every content slide)

```javascript
// Teal header (solution/positive slides)
s.addShape(pres.shapes.RECTANGLE, {
  x:0, y:0, w:10, h:1.05, fill:{color:TEAL}, line:{color:TEAL}
});
s.addText("Slide Title Here", {
  x:0.45, y:0.12, w:9.1, h:0.75,
  fontSize:26, bold:true, color:WHITE, fontFace:"Calibri", margin:0
});

// Navy header (problem/team/status slides) — same but fill:{color:NAVY}
```

---

## Step 4: Visual QA (Mandatory)

After generating the `.pptx`, always run visual QA before declaring done.

```bash
# Convert to PDF
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf /mnt/user-data/outputs/YourDeck.pptx

# Convert to images (check where PDF was saved first)
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 /path/to/YourDeck.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

Then view every slide image and check for:

| Issue | What to look for |
|-------|-----------------|
| Text overflow | Any text cut off at box edges or slide edges |
| Overlap | Text running into shapes, callout boxes, or other text |
| Low contrast | Light text on light background, dark icons on dark fills |
| Missing content | Any slide that's missing an expected element |
| Leftover placeholders | Any "TODO", "lorem ipsum", or bracket text |

**Fix any user-visible issue. Do not iterate on sub-pixel nudges.**
One fix-and-verify cycle is sufficient unless a new defect appears.

---

## Step 5: Deliver

Use `present_files` to share the output. Give a brief summary of:
- What's in the deck (slide list)
- Any assumptions you made about the brief
- Any elements the user should review or adjust (dates, names, team labels)

Do not write a long post-amble. The user can open the file.

---

## Slide-Type Reference

Quick lookup for common slide types and which layout pattern to use.

| Slide type                    | Layout pattern          | Header color |
|-------------------------------|-------------------------|--------------|
| Title                         | Dark background, no bar | —            |
| Problem statement             | Two-column              | Navy         |
| Solution overview             | Dark banner + 3 cards   | Teal         |
| Process / pipeline            | Numbered steps          | Navy         |
| Human controls / governance   | 3 cards + footer bar    | Teal         |
| Team / ownership              | 3 cards with color tags | Navy         |
| Timeline / roadmap            | Horizontal spine        | Teal         |
| Value proposition / close     | 2×2 cards on dark bg    | —            |
| Exec summary                  | Large stat callouts     | Navy         |
| Comparison (before/after)     | Two-column contrast     | Navy or Teal |

---

## Audience Calibration

Adjust content depth based on the stated audience:

| Audience            | Jargon level | Explanation depth | Slide density |
|---------------------|-------------|-------------------|---------------|
| Junior analyst      | None        | Define everything | Low — fewer points per slide |
| Senior internal     | Full        | Assume context    | Medium |
| Exec / sponsor      | Low         | Lead with impact  | Low — headlines only |
| External client     | Low–medium  | Define on first use | Medium |
| Technical peer      | Full        | Skip basics       | High |

---

## Package Contents

This toolkit contains:

```
pptx-toolkit/
├── INSTRUCTIONS.md      ← This file. Read first.
├── STYLE_GUIDE.md       ← Palette, typography, layout patterns, tone rules
└── README.md            ← One-page summary for humans handing this to an AI
```

The underlying build skill lives at:
```
/mnt/skills/public/pptx/
├── SKILL.md             ← Environment setup, QA process
└── pptxgenjs.md         ← Full PptxGenJS API reference
```

Point any AI at INSTRUCTIONS.md + STYLE_GUIDE.md + the two skill files.
That's the complete toolkit.
