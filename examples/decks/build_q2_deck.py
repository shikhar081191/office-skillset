from create_pptx import PptxBuilder
from patterns import BLACKROCK, numbers_slide, process_slide, status_slide, timeline_slide
from project_workspace import ProjectWorkspace

project = ProjectWorkspace("AI MDS Q2 Update")
b = PptxBuilder(palette=BLACKROCK)

# Slide 1 — Title
b.title_slide(
    title="AI MDS: Q2 Progress Update",
    subtitle="AFE Portfolio Risk Modelling | May 2026",
)

# Slide 2 — Numbers
numbers_slide(
    title="AI MDS at a Glance",
    stats=[
        {"number": "5",       "label": "Pipeline steps",        "context": "From plain language input to portfolio P&L"},
        {"number": "9",       "label": "Active Q2 workstreams", "context": "Across build, infra, validation and governance"},
        {"number": "Mid-Q2",  "label": "Beta target with RQA",  "context": "Conditional on node stability"},
        {"number": "H2 2026", "label": "Production release",    "context": "Pending governance sign-off"},
    ],
    builder=b,
)

# Slide 3 — Process
process_slide(
    title="How the Pipeline Works",
    steps=[
        {"name": "Build Narrative",        "description": "Translate market situation into structured briefing"},
        {"name": "Historical Analogues",   "description": "Anchor scenario to precedent, not clone it"},
        {"name": "Factor Identification",  "description": "Select the risk factors this scenario actually moves"},
        {"name": "Factor Shocks",          "description": "Calibrate shock magnitudes using historical precedent"},
        {"name": "Covariance Propagation", "description": "GARCH-t-DCC spreads shocks across thousands of Aladdin factors"},
    ],
    footnote="HITL checkpoints run at every step — human judgment stays in the loop",
    builder=b,
)

# Slide 4 — Status
status_slide(
    title="Q2 Workstream Status",
    workstreams=[
        {"name": "Node Improvements",    "status": "Four core nodes being improved per spec",               "rag": "green", "owner": "Ai"},
        {"name": "Infrastructure",       "status": "RockAI onboarding underway; Surface API fallback ready", "rag": "green", "owner": "Sa"},
        {"name": "Validation Framework", "status": "Actor-critic prototype scoped; precedes case studies",   "rag": "green", "owner": "B"},
        {"name": "HITL Scoping",         "status": "Blocked — cannot start until node spec is finalised",   "rag": "amber", "owner": "Ai"},
        {"name": "Beta Testing",         "status": "Not started — depends on node stability",               "rag": "amber", "owner": "RQA"},
        {"name": "GARCH-DCC",            "status": "Promising early results; isolated from main timeline",  "rag": "amber", "owner": "D"},
        {"name": "Governance",           "status": "Scoping this quarter; production framework next",       "rag": "amber", "owner": "Sh"},
    ],
    builder=b,
)

# Slide 5 — Timeline
# "Current marker: May 2026" → validation framework (last completed item) marked current
timeline_slide(
    title="Q2 Delivery Schedule",
    milestones=[
        {"date": "May 26", "label": "Node spec complete",       "state": "past"},
        {"date": "May 26", "label": "Validation framework",     "state": "current"},
        {"date": "Jun 26", "label": "Actor-critic prototype",   "state": "future"},
        {"date": "Jun 26", "label": "RQA beta testing",         "state": "future"},
        {"date": "Jul 26", "label": "Analytics onboarded",      "state": "future"},
        {"date": "Sep 26", "label": "Prod readiness review",    "state": "future"},
    ],
    builder=b,
)

deck_path = project.output_path("ai_mds_q2_update.pptx")
out = b.save(deck_path, final=True, report_path=project.qa_path(deck_path.name))
print(f"Saved: {out}  ({b.prs.slides.__len__()} slides)")
