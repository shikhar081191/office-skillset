# Examples

Reusable scenario builds and sample inputs live here instead of occupying the
library root.

```text
examples/
  decks/    Standalone presentation build examples
  inputs/   Sample source files for workflow experiments
```

Run a retained example from the repository root with module syntax:

```bash
python -m examples.decks.build_credit_factor_decay_deck
```

Real deliverables belong in `projects/<project>/`; the retained Q2 project has
its own build script within its project folder. For Word-input experiments, a
generated scaffold is only intermediate input for AI refinement and should not
be treated as the presentation delivered to a user.
