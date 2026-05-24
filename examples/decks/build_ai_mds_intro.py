from decision_log import DecisionLog
from patterns import numbers_slide, process_slide
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace


def build_deck():
    project = ProjectWorkspace("AI MDS Intro for Junior Analyst")
    instructions_path = project.working_path("user_instructions.txt")
    instructions_path.write_text(
        "AI assistant should act as the decision layer: choose slide patterns, document reasoning, and log the build in working/decision_log.md.\n",
        encoding="utf-8",
    )

    log = DecisionLog(project.working_path("decision_log.md"))
    log.record_slide_plan([
        {"section": "Cover", "pattern": "title_slide", "assertion": "Introduce AI MDS in a simple, non-technical way."},
        {"section": "Problem", "pattern": "two_column_slide", "assertion": "Show why the current market-shock process is too slow and manual."},
        {"section": "Solution", "pattern": "two_column_slide", "assertion": "Explain how AI MDS turns plain language into portfolio impact."},
        {"section": "Pipeline", "pattern": "process_slide", "assertion": "Break the system into five clear steps."},
        {"section": "Human review", "pattern": "two_column_slide", "assertion": "Make the human-in-the-loop oversight explicit and transparent."},
        {"section": "Teams", "pattern": "two_column_slide", "assertion": "Summarise the three groups responsible for build and adoption."},
        {"section": "Status", "pattern": "numbers_slide", "assertion": "Capture the current prototype, UI, beta and production timing."},
        {"section": "Value", "pattern": "two_column_slide", "assertion": "Explain why consistency and auditability matter more than speed."},
    ])

    b = PptxBuilder(palette="midnight_executive")

    b.title_slide(
        title="AI MDS — Quick market shock analysis",
        subtitle="What it does, how it works, and why it matters"
    )

    b.two_column_slide(
        title="Why the old approach is too slow",
        left_heading="The issue",
        left_body=[
            "Market shocks need fast answers.",
            "The existing process takes weeks.",
        ],
        right_heading="What happens now",
        right_body=[
            "Experts debate the scenario and choose risk drivers.",
            "They manually estimate shocks and run the portfolio.",
            "By then, the market has already moved.",
        ],
    )

    b.two_column_slide(
        title="What AI MDS does",
        left_heading="Input",
        left_body=[
            "Type the market event in plain English.",
            "Describe it the way you would tell a colleague.",
        ],
        right_heading="Output",
        right_body=[
            "A structured scenario, the relevant risk factors, and portfolio P&L.",
            "The whole analysis is delivered in minutes.",
        ],
    )

    process_slide(
        title="Five steps in the AI MDS pipeline",
        steps=[
            {
                "name": "Read the situation",
                "description": "Turn plain language into a clear scenario."
            },
            {
                "name": "Check history",
                "description": "Find similar episodes as a plausibility check."
            },
            {
                "name": "Pick factors",
                "description": "Choose the risk drivers the scenario actually moves."
            },
            {
                "name": "Estimate shocks",
                "description": "Set how much each chosen factor should move."
            },
            {
                "name": "Propagate impact",
                "description": "Run the shocks through the portfolio model."
            },
        ],
        highlight=2,
        builder=b,
    )

    b.two_column_slide(
        title="Human review is built in",
        left_heading="What the expert does",
        left_body=[
            "Reviews the scenario output at every step.",
            "Can override the factor selection or shock sizes.",
        ],
        right_heading="What the system keeps",
        right_body=[
            "Stores overrides and explanation notes.",
            "Makes every decision visible and auditable.",
        ],
    )

    b.two_column_slide(
        title="Who is building this",
        left_heading="Core team",
        left_body=[
            "Architecture & modelling: builds the pipeline and risk models.",
            "Analytics Product: deploys the UI and infrastructure.",
        ],
        right_heading="Business owners",
        right_body=[
            "RQA: uses the system and shapes requirements every two weeks.",
        ],
    )

    numbers_slide(
        title="Where we are today",
        stats=[
            {
                "number": "Prototype",
                "label": "End-to-end flow works",
                "context": "The full pipeline is operational.",
            },
            {
                "number": "Live UI",
                "label": "Interface is available",
                "context": "Users can try the system now.",
            },
            {
                "number": "Mid 2026",
                "label": "Beta testing starts",
                "context": "RQA will begin broader trials.",
            },
            {
                "number": "H2 2026",
                "label": "Production target",
                "context": "Full rollout is planned later this year.",
            },
        ],
        builder=b,
    )

    b.two_column_slide(
        title="Why this matters more than speed",
        left_heading="Consistency",
        left_body=[
            "Every scenario follows the same structured process.",
            "Every assumption and override is recorded.",
        ],
        right_heading="Value",
        right_body=[
            "Governance and audit become simpler.",
            "Experts can focus on judgement, not mechanical work.",
        ],
    )

    output_path = project.output_path("ai_mds_intro_for_junior.pptx")
    qa_path = project.qa_path(output_path.name)
    b.save(output_path, final=True, report_path=qa_path)

    log.record_qa(qa_path)
    log.record_output(output_path)

    project.record_workflow(
        "ai_mds_intro_build",
        inputs=[],
        outputs=[output_path],
        working_files=[instructions_path, project.working_path("decision_log.md")],
        qa_reports=[qa_path],
        palette="midnight_executive",
        patterns=["title_slide", "two_column_slide", "process_slide", "numbers_slide"],
        settings={"ai_layer": True, "source": "chat"},
    )

    print(f"Saved deck to {output_path}")


if __name__ == "__main__":
    build_deck()
