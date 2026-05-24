from patterns import numbers_slide, process_slide, status_slide, timeline_slide
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace


def build_deck():
    project = ProjectWorkspace("AI MDS Q2 Progress Update")
    b = PptxBuilder(palette="midnight_executive")

    b.title_slide(
        title="AI MDS: Q2 Progress Update",
        subtitle="AFE Portfolio Risk Modelling | May 2026"
    )

    # Strong opening headline to land the key message for Claire
    b.content_slide(
        title="Headline — On track; Beta mid‑Q2",
        body=[
            "End-to-end prototype live; UI operational; RQA using it",
            "Beta targeted mid-Q2 — gating item: node stability (Aidan, on track)",
            "Validation (actor-critic) must precede RQA case studies (Benedek; on track)",
            "GARCH-t-DCC research progressing in parallel (Dave); governance planning started",
        ],
        dark_bg=True,
    )

    numbers_slide(
        title="Q2 numbers",
        stats=[
            {
                "number": "5",
                "label": "Pipeline steps end-to-end",
                "context": "From plain language input to portfolio P&L"
            },
            {
                "number": "9",
                "label": "Active Q2 workstreams",
                "context": "Across build, infra, validation and governance"
            },
            {
                "number": "Mid-Q2",
                "label": "Beta target with RQA",
                "context": "Conditional on node stability"
            },
            {
                "number": "H2 2026",
                "label": "Production release",
                "context": "Pending governance sign-off"
            },
        ],
        builder=b,
    )

    process_slide(
        title="How the Pipeline Works",
        builder=b,
        steps=[
            {
                "name": "Build Narrative",
                "description": "Translate market situation into structured briefing."
            },
            {
                "name": "Historical Analogues",
                "description": "Anchor scenario to precedent, not clone it."
            },
            {
                "name": "Factor Identification",
                "description": "Select the risk factors this scenario actually moves."
            },
            {
                "name": "Factor Shocks",
                "description": "Calibrate shock magnitudes using historical precedent."
            },
            {
                "name": "Covariance Propagation",
                "description": "GARCH-t-DCC spreads shocks across thousands of Aladdin factors."
            },
        ],
        footnote="HITL checkpoints run at every step — human judgment stays in the loop"
    )

    status_slide(
        title="Q2 Workstream Status",
        builder=b,
        workstreams=[
            {
                "name": "Node Improvements",
                "status": "Four core nodes being improved per spec",
                "rag": "green",
                "owner": "A"
            },
            {
                "name": "Infrastructure",
                "status": "RockAI onboarding underway, Surface API fallback ready",
                "rag": "green",
                "owner": "S"
            },
            {
                "name": "Validation Framework",
                "status": "Actor-critic prototype scoped, precedes case studies",
                "rag": "green",
                "owner": "B"
            },
            {
                "name": "HITL Scoping",
                "status": "Blocked — cannot start until node spec is finalised",
                "rag": "amber",
                "owner": "A"
            },
            {
                "name": "Beta Testing",
                "status": "Not started — depends on node stability",
                "rag": "amber",
                "owner": "R"
            },
            {
                "name": "GARCH-DCC",
                "status": "Promising early results, isolated from main timeline",
                "rag": "amber",
                "owner": "D"
            },
            {
                "name": "Governance",
                "status": "Scoping this quarter, production framework next",
                "rag": "amber",
                "owner": "S"
            },
        ]
    )

    timeline_slide(
        title="Q2 Delivery Schedule",
        builder=b,
        milestones=[
            {"date": "May 2026", "label": "Node spec complete", "state": "past"},
            {"date": "May 2026", "label": "Validation framework in place", "state": "past"},
            {"date": "June 2026", "label": "Actor-critic prototype live", "state": "future"},
            {"date": "June 2026", "label": "RQA beta testing begins", "state": "future"},
            {"date": "July 2026", "label": "Analytics Product onboarded", "state": "future"},
            {"date": "September 2026", "label": "Production readiness review", "state": "future"},
        ]
    )

    output_path = project.output_path("ai_mds_q2_progress_update.pptx")
    b.save(output_path, final=True, report_path=project.qa_path(output_path.name))
    print(f"Saved deck to {output_path}")


if __name__ == "__main__":
    build_deck()
