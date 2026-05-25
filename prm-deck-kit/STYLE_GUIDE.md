# Deck Style Guide — Portfolio Risk Modelling, BlackRock

This file defines the visual and tonal standards for presentation decks
produced for or by the Portfolio Risk Modelling team. Any AI building
decks in this context should treat this as a hard constraint document,
not a suggestion list.

---

## Color Palette

Use this palette. Do not substitute without explicit instruction.

| Role            | Name              | Hex       | Usage                                      |
|-----------------|-------------------|-----------|--------------------------------------------|
| Primary dark    | Navy              | `0D1B2A`  | Title slides, header bars, dark backgrounds |
| Primary accent  | Teal              | `0A9396`  | Header bars on content slides, key accents  |
| Light teal      | Pale teal         | `94D2BD`  | Subtext on dark backgrounds, secondary text |
| Background      | Off-white         | `E9F5F5`  | Slide backgrounds (alternating with white)  |
| Surface         | White             | `FFFFFF`  | Cards, content areas                        |
| Warning/callout | Amber             | `EE9B00`  | Warnings, callouts, "WEEKS"-style punches   |
| Body text       | Charcoal          | `4A5568`  | Body text on white/light backgrounds        |
| Dividers        | Light gray        | `E2E8F0`  | Card borders, divider lines                 |

**Palette rhythm:** Alternate slide backgrounds between white (`FFFFFF`) and
off-white (`E9F5F5`). Title slide and closing slide use dark navy (`0D1B2A`).
This creates a "sandwich" — dark / light / light / ... / dark.

---

## Typography

| Element          | Font face  | Size  | Weight | Notes                            |
|------------------|------------|-------|--------|----------------------------------|
| Slide title      | Calibri    | 26–30pt | Bold  | In white on colored header bar   |
| Section header   | Calibri    | 14–17pt | Bold  | Navy on white/light backgrounds  |
| Body text        | Calibri    | 11–13pt | Regular | Charcoal                       |
| Callout text     | Calibri    | 10–12pt | Italic | Amber or warning-colored boxes  |
| Footer/meta      | Calibri    | 10–11pt | Regular | Muted, near bottom of slide    |

Never use Arial, Times, or other fonts unless the template explicitly requires it.

---

## Layout Patterns

### Slide header bar
Every content slide starts with a full-width colored header bar (10" × 1.05"):
- **Teal** (`0A9396`) for "solution / process / positive" slides
- **Navy** (`0D1B2A`) for "problem / team / status" slides
- White title text at 26–30pt bold, positioned at `x:0.45, y:0.12`

### Content area starts at y = 1.2"
All content begins below the header bar. Maintain 0.45" left margin minimum.

### Three-column card layout
Use for: team slides, feature comparisons, three-property breakdowns.
- Card width: ~2.95–3.0"
- Gap between cards: ~0.15–0.2"
- Cards start at x ≈ 0.4", 3.55", 6.7" (adjust slightly for centering)
- Each card gets a thin top accent bar in the appropriate color
- Shadow: `{ type:"outer", blur:5, offset:2, angle:135, color:"000000", opacity:0.1 }`

### Two-column layout
Use for: problem/solution contrast, before/after, left-context / right-detail.
- Left column: ~4.2" wide starting at x=0.45
- Right column: ~4.4" wide starting at x=5.1
- Use a border + fill contrast to distinguish (e.g., warm fill + amber border for "the problem")

### Timeline layout
- Horizontal spine at y ≈ 2.8", from x=0.9 to x=9.1
- Milestone dots: OVAL 0.4×0.4", centered on spine
- Completed milestones: teal fill; future milestones: light gray fill, gray border
- Labels above spine, descriptions below
- "NOW" marker: amber rectangle to the left of the current dot

### Process / numbered steps layout
- Numbered circle (OVAL, 0.58×0.58") at x=0.35
- Step title bold at x=1.1, step description at x=1.1 below
- Connector lines between dots (LINE shape, light gray)
- Keep description text width ≤ 6.2" to leave room for any right-side callout

### Full-width callout bar
For closing statements or governance notes:
- Navy rectangle spanning full slide width (x=0, w=10, h=0.55–0.65")
- Light teal italic text inside

---

## Structural Patterns (Slide Sequencing)

For a standard project explainer deck aimed at a non-specialist audience,
use this sequence:

1. **Title slide** — project name + one-line premise
2. **The Problem** — what pain exists today, why it matters
3. **The Solution** — the core idea in plain terms
4. **How It Works** — the process or pipeline
5. **The Human Element** — controls, overrides, governance
6. **The Team** — who owns what
7. **Where We Are** — timeline / status
8. **Why It Matters** — close with the deeper value proposition

This sequence works for technical AI/risk projects presented to non-specialists.
Adjust freely for different deck purposes (exec update, stakeholder review, etc.)

---

## Tone and Language Rules

These apply to all text written for this team's decks:

**Target audience default:** Assume the reader knows finance but not the specific
model or system being described. Avoid internal abbreviations without defining them
on first use (exception: "P&L" and "portfolio" are always safe).

**Sentence style:**
- Short, declarative sentences. One idea per sentence.
- Active voice. "The system reads your input" not "Your input is read by the system."
- No hedge-stacking. "This may potentially help to..." → "This helps..."
- No filler openings. Not "In today's rapidly evolving landscape..." Just start with the point.

**Technical terms to always explain:**
- Risk factors → "the measurable market variables (interest rates, equity prices, FX) that drive portfolio value"
- P&L → acceptable shorthand; no need to expand
- Covariance model → "a statistical model that captures how risk factors move together"
- Shocks → "estimated moves in each risk factor under the scenario"

**What to avoid:**
- "Leveraging" anything
- "Robust" as a filler adjective
- "End-to-end" without explaining what the ends are
- Bullet lists longer than 5 items — split into two groups or use a card layout instead

---

## Common Mistakes to Avoid

- **Never** put a decorative line under a title (hallmark of AI-generated slides)
- **Never** use cream/beige backgrounds — use white or the off-white from the palette
- **Never** encode opacity in hex color strings — use the `opacity` property
- **Never** reuse shape option objects across multiple `addShape` calls — PptxGenJS mutates them
- **Always** limit description text width on process slides to avoid callout overlap
- **Always** run visual QA (convert to PDF → images → inspect) before declaring done
- **Always** use `margin: 0` on text boxes that need to align precisely with shapes
