# Skill: Layout Density

## Preference
Every content slide should fill the canvas. Empty space reads as unfinished.

**Ideal:** 1–2 visual elements (table, chart image, or named pattern) supported by short bullets.
**Acceptable fallback:** text-only slide using ≥18pt body font so the canvas looks intentional.
**Never:** small-font bullets floating above an empty lower half of the canvas.

---

## QA checks that enforce this

| Check | Fires when | Action required |
|---|---|---|
| `visual_density` | Content slide has no table, image, or structural pattern | Replace with a named pattern or add a visual element |
| `text_only_font` | Text-only slide has any font <18pt | Increase body font to ≥18pt, or switch to a visual pattern |
| `space_utilisation` | Text-only slide has <40 words total | Merge with adjacent slide, increase font, or add a visual |

The QA gate fires these as **warnings** — they do not block delivery, but the AI must address
each one before handing off the file.

---

## How to fix a `visual_density` warning

Choose the pattern that best fits the content type:

| If the content is… | Use this pattern |
|---|---|
| A key finding + supporting evidence | `assertion_evidence_slide` |
| A step-by-step method or process | `process_slide` |
| A roadmap or milestone schedule | `timeline_slide` |
| 3–4 headline statistics | `numbers_slide` |
| A numeric model or scenario comparison | `results_slide` or `heat_map_slide` |
| A chart with a headline and so-what | `chart_context_slide` |
| Workstream or programme health | `status_slide` |
| Positive/negative contribution breakdown | `waterfall_slide` |
| Criteria-based model or vendor selection | `scorecard_slide` |

Only keep `content_slide` when the content genuinely cannot be expressed as any of the above
**and** increase the body font to ≥18pt so the slide does not look sparse.

---

## How to fix a `text_only_font` warning

`PptxBuilder.content_slide()` auto-scales body font by bullet count:

| Bullets | Auto font |
|---|---|
| ≤3 | 20pt |
| 4–5 | 16pt |
| 6+ | 14pt |

If auto-scaling still produces a font below 18pt (i.e., 6+ bullets), either:
1. Reduce the bullet count (distil the content further), or
2. Switch to a visual pattern that handles the volume better.

---

## How to fix a `space_utilisation` warning

In order of preference:
1. Switch to a visual pattern — this resolves the warning and improves the slide.
2. Reduce bullet count so font auto-scales to 20pt.
3. Merge the slide with an adjacent slide that is also light on content.

---

## Design rule

Never deliver a slide where the lower 30% of the canvas is visually empty.
A `content_slide` with 3 or fewer short bullets must use ≥18pt font or include a visual element.
