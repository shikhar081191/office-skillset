# Adding a Skill (New Slide Pattern)

This guide explains how to document a new pattern you've built in `patterns.py`
so it is automatically picked up by AI assistants and reusable by the team.

---

## Step 1 — Write the pattern function in patterns.py

Add a new function to `patterns.py`. Follow the conventions of the existing patterns:

```python
def my_pattern_slide(title, my_arg, another_arg, builder=None, palette=None):
    """One-line docstring describing what this pattern produces."""
    b, p = _palette_builder(builder, palette)
    slide = b.prs.slides.add_slide(b._blank_layout)
    # ... build the slide using b._add_rectangle, b._add_text_box, _oval, etc.
    return b
```

**Rules:**
- First positional arg is always `title: str`
- Last two kwargs are always `builder=None, palette=None`
- Never hardcode hex colours — always reference `p["primary"]`, `p["secondary"]`,
  `p["accent"]`, `p["text_dark"]`, `p["text_light"]`
- Return the builder so the caller can chain or save
- Use the active builder palette for every fill, line and text colour. In
  addition to core roles, use `background`, `surface`, `surface_alt`, `border`,
  `muted_text`, `positive`, `warning`, `negative`, and `highlight`.
- When an existing builder is supplied, its palette is authoritative unless an
  explicit `palette=` override is provided.

---

## Step 2 — Add a demo call in demo_all()

At the bottom of `patterns.py`, add your pattern to `demo_all()`:

```python
my_pattern_slide(
    title="Demo — My Pattern",
    my_arg=...,
    another_arg=...,
).save("output/demo_my_pattern.pptx")
print("output/demo_my_pattern.pptx")
```

Run `python patterns.py` to confirm the output file is generated without errors.

---

## Step 3 — Create a skill file in skills/

Create `skills/my_pattern.md` using this structure:

```markdown
# Pattern: My Pattern Name

## When to use
[2–4 sentences on when this pattern is the right choice and what content it suits]

## Function signature
[Code block showing the function call with all arguments and their types]

## Example call
[A complete, runnable Python example using realistic domain data]

## Content-shaping rules
[Numbered list of rules for how an AI should distil raw user content into the
right arguments — what to extract, how to shorten, how to resolve ambiguity]
```

**Required sections** (all four must be present):
1. `When to use` — helps AI assistants pick this pattern over others
2. `Function signature` — exact API including all kwargs
3. `Example call` — copy-pasteable, uses realistic data
4. `Content-shaping rules` — teaches the AI how to translate user text into fields

---

## Step 4 — Update the README pattern table

Open `README.md` and add a row to the pattern table:

```markdown
| `my_pattern_slide()` | My Pattern | When to reach for it (one line) |
```

---

## Checklist

- [ ] Function added to `patterns.py` with `builder=None, palette=None` kwargs
- [ ] All colours reference palette keys (no hardcoded hex)
- [ ] Demo call added to `demo_all()` and runs without error
- [ ] `skills/my_pattern.md` created with all four required sections
- [ ] README pattern table updated
