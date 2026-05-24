from patterns import status_slide, process_slide, timeline_slide
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace


def build_deck():
    project = ProjectWorkspace("AI MDS Q2 Status")
    b = PptxBuilder(palette="midnight_executive")

    b.title_slide(
        title="AI Market-Driven Scenario (AI MDS) — Q2 Status",
        subtitle="For Claire | Q2 2026 update"
    )

    b.stat_slide(
        title="Q2 snapshot",
        stats=[
            ("5", "Pipeline steps"),
            ("9", "Active workstreams"),
            ("Mid-Q2", "Beta target"),
            ("H2 2026", "Production release"),
        ]
    )

    process_slide(
        title="AI MDS pipeline",
        steps=[
            {
                "name": "Build Narrative",
                "description": "Create news briefing and catalyst events to generate qualitative insight."
            },
            {
                "name": "Historical Analogues",
                "description": "Anchor scenarios in precedent; use history as context, not a clone."
            },
            {
                "name": "Factor Identification",
                "description": "Select the most critical drivers; move from economist insight to selector execution."
            },
            {
                "name": "Factor Shocks",
                "description": "Calibrate shocks using analogue precedent to keep scenarios plausible."
            },
            {
                "name": "Covariance Propagation",
                "description": "Use GARCH-t-DCC to propagate shocks across thousands of correlated factors."
            },
        ],
        highlight=2,
        footnote="Beta testing timing is conditional on node stability and RQA readiness.",
        builder=b,
    )

    status_slide(
        title="Workstream status — Q2 2026",
        workstreams=[
            {
                "name": "Node Improvements",
                "status": "In progress; on track",
                "rag": "green",
                "owner": "A"
            },
            {
                "name": "Infrastructure",
                "status": "In progress; on track",
                "rag": "green",
                "owner": "S"
            },
            {
                "name": "Validation Framework",
                "status": "In progress; on track",
                "rag": "green",
                "owner": "B"
            },
            {
                "name": "GARCH-DCC",
                "status": "In progress; isolated from main timeline",
                "rag": "green",
                "owner": "D"
            },
            {
                "name": "HITL Scoping",
                "status": "Blocked; pending node spec",
                "rag": "amber"
            },
            {
                "name": "Beta + Governance",
                "status": "Amber; beta depends on node stability; governance scoping this quarter",
                "rag": "amber"
            },
        ],
        builder=b,
    )

    timeline_slide(
        title="Q2 milestones",
        milestones=[
            {"date": "May", "label": "Node spec complete", "state": "past"},
            {"date": "June", "label": "Actor-critic prototype live", "state": "current"},
            {"date": "June", "label": "RQA beta testing begins", "state": "future"},
            {"date": "July", "label": "Analytics Product onboarded", "state": "future"},
            {"date": "Sept", "label": "Production readiness review", "state": "future"},
        ],
        builder=b,
    )

    output_path = project.output_path("ai_mds_q2_status.pptx")
    b.save(output_path, final=True, report_path=project.qa_path(output_path.name))
    print(f"Saved deck to {output_path}")


if __name__ == "__main__":
    build_deck()
