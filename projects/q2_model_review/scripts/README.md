# Q2 Model Review Scripts

This folder contains project-specific refined build scripts. Run the retained
builder from the repository root:

```bash
python -m projects.q2_model_review.scripts.build_refined_deck
```

For a copied workspace or dry run, target that project without overwriting the
canonical project output:

```bash
python -m projects.q2_model_review.scripts.build_refined_deck --project "Q2 Model Review Dry Run"
```

The project input, outputs, QA reports and decision log remain in their
existing sibling folders. The builder runs final PPTX QA and registers the
refined output with its slide plan and QA evidence. It also records rendered
review status using the optional render utilities; inspect slide previews when
they are available, or retain the explicit skipped-runtime report. Review and
repair any actionable QA warnings before delivery.
