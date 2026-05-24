# Workflow Skill: Word Quality Gate

## When to use
Use after creating a Word deliverable such as a research note, whitepaper,
short memo or model-validation paper.

## Required workflow
Final document scripts must save with the built-in structural audit:

```python
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("Research Note")
doc_path = project.output_path("research_note.docx")
b.save(doc_path, final=True, report_path=project.qa_path(doc_path.name))
```

Fix any blocking placeholder/content failures before delivery. For formal
deliverables, also use rendered page inspection when the optional DOCX render
runtime is available, because structural checks cannot see page-flow defects.
