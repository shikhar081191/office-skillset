# Skill: Layout Symmetry, Overflow Prevention, and AI Labeling

## Rules

1. **No text overflow** — text must never escape its bounding box or render over adjacent content
2. **No shape overlap** — two text boxes must never overlap each other
3. **Auto-fit is on** — `content_slide` body boxes use `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`; PowerPoint shrinks font rather than overflowing; pass `auto_fit=True` to `b._add_text_box()` in custom scripts when text length is uncertain
4. **Symmetry** — shapes at the same visual level must share identical `w`, `h`, and spacing values

---

## QA checks that enforce this

| Check | Level | Fires when |
|---|---|---|
| `text_crowding` | warning; **error in final mode** when overflow > 1.5× capacity | Estimated text lines exceed box capacity |
| `shape_overlap` | warning | Two text boxes have overlapping bounding boxes (>5% area overlap) |
| `layout_symmetry` | warning | Three or more repeated large panels in an aligned row or column have unequal dimensions or uneven spacing |

These checks fire automatically on every `b.save(..., final=True)` call.

---

## Fixing a `text_crowding` warning or error

1. Shorten the text — distil bullets to ≤12 words each
2. Reduce bullet count — aim for ≤5 per slide
3. Switch to a pattern with more vertical space (e.g. `assertion_evidence_slide`, `process_slide`)
4. Do **not** reduce font size below 8pt to force text to fit — fix the content instead

## Fixing a `shape_overlap` warning

1. Increase the vertical gap between the two shapes
2. Reduce text in the upper shape so it fits within its assigned box
3. Check that `auto_fit=True` is set on the overflowing shape so PowerPoint can shrink it

## Symmetry rules for multi-shape layouts

- Two-column slides: both columns must use identical `w` and `h` values
- Card grids (status, scorecard): all cards must use the same `card_w_in` and `card_h_in`
- Timeline milestones: all label text boxes must use the same width
- Never nudge one shape without applying the same offset to its sibling shapes
- The automated check is deliberately limited to repeated panels; complex
  chart balance and optical alignment still require rendered slide review.

---

## AI-generated content labeling

When a slide contains content that was **synthesised, inferred, or not directly present** in the
user's source material, mark it immediately after building the slide:

```python
b.ai_label()  # adds a red "AI GENERATED — verify before use" badge to the last slide
```

The badge is a red rectangle (hex `C0392B`) at the bottom-right corner with white 8pt bold text.
It does not interfere with slide content and is visible in PowerPoint's normal and presenter views.

### When to call `ai_label()`

Call it on any slide where you:
- Added a number, statistic, or date not present in the source
- Wrote an assertion that goes beyond what the source explicitly states
- Added an example, analogy, or illustration not provided by the user
- Inferred a trend, conclusion, or recommendation from indirect signals

### When NOT to call it

Do **not** call it when content was directly extracted or distilled from the user's source,
even if you rephrased or condensed it.

---

## Source attribution

When slide content comes from a specific external source (report, dataset, model), add a
source badge at the bottom-left immediately after building the slide:

```python
b.source_label("Bloomberg, Q2 2025")  # or "MSCI Barra", "S&P LossStats", etc.
```

The badge is a small muted italic line at `y = 6.97"`. On a split slide where
the bottom-left is dark, position the label on a light panel so it remains
legible:

```python
b.source_label("Validation report, May 2025", x=6.45, width=6.1)
```

It does not conflict with `ai_label()` (which sits at bottom-right).

---

## Speaker notes

Add presenter context to a slide with:

```python
b.speaker_notes("Emphasise that the PMSE improvement holds on out-of-time validation — this was the committee's main concern in Q1.")
```

Notes are stored in the PowerPoint notes pane. They are not visible during screen sharing.
Use them for:
- Anticipated audience questions and suggested answers
- Data caveats the presenter should mention verbally
- References to supporting appendix slides

### Example

```python
from create_pptx import PptxBuilder
from patterns import assertion_evidence_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Model Review")
b = PptxBuilder(palette="blackrock")

b.title_slide("Model Review Q2 2025")

# Sourced directly from the brief — no label
assertion_evidence_slide(
    "Core Finding",
    "Bills drive almost all of the PMSE improvement",
    [{"stat": "46pp", "text": "lower PMSE in short-duration cohorts"}],
    builder=b,
)

# Context synthesised by the AI — label it
b.content_slide("Market Context", ["Rate volatility elevated throughout H1", "IG spreads tightened 40bps"])
b.ai_label()

deck_path = project.output_path("review.pptx")
b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
```
