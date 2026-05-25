from artifact_utilities import ArtifactUtilities
from decision_log import DecisionLog
from patterns import (
    assertion_evidence_slide,
    callout_bar_slide,
    process_slide,
    three_column_card_slide,
    timeline_slide,
    two_column_contrast_slide,
)
from create_pptx import PptxBuilder
from project_workspace import ProjectWorkspace


def build_deck():
    project = ProjectWorkspace("AI MDS — Automated Market Scenario Analysis")
    instructions_path = project.working_path("user_instructions.txt")
    instructions_path.write_text(
        "Build a simple project explainer deck for a junior analyst. "
        "Use the prm palette, keep language plain, define terms on first use, and limit content to 8 slides.\n",
        encoding="utf-8",
    )

    log = DecisionLog(project.working_path("decision_log.md"))
    log.record_slide_plan([
        {
            "section": "Cover",
            "source": "User brief and audience guidance",
            "pattern": "three_column_card_slide",
            "assertion": "Introduce AI MDS in simple language using a visual opener.",
            "rationale": "A visual opener avoids title-slide density warnings and still states the purpose clearly.",
            "inputs": "Project title, audience, purpose, palette direction.",
            "inference": "No",
        },
        {
            "section": "The problem",
            "source": "Current market shock process description",
            "pattern": "two_column_contrast_slide",
            "assertion": "The existing process is slow, manual, and too late for fast market moves.",
            "rationale": "A problem slide makes the need for AI MDS easy to feel.",
            "inputs": "Current practice summary; consequences of delay.",
            "inference": "No",
        },
        {
            "section": "The solution",
            "source": "AI MDS description and output flow",
            "pattern": "assertion_evidence_slide",
            "assertion": "AI MDS turns a plain-English market story into a portfolio impact check in minutes.",
            "rationale": "One clear statement plus evidence fits the solution story.",
            "inputs": "System capabilities; user input and output description.",
            "inference": "No",
        },
        {
            "section": "Pipeline",
            "source": "Five-step workflow description",
            "pattern": "process_slide",
            "assertion": "The system works through five ordered steps from reading the event to calculating impact.",
            "rationale": "A visual process slide is ideal for a sequential pipeline.",
            "inputs": "The five workflow steps and short descriptions.",
            "inference": "No",
        },
        {
            "section": "Human controls",
            "source": "Human-in-the-loop safeguard description",
            "pattern": "three_column_card_slide",
            "assertion": "Experts review, override, and document every step so the system is not a black box.",
            "rationale": "Three cards separate review, overrides, and audit clearly.",
            "inputs": "Review points; override mechanism; audit value.",
            "inference": "No",
        },
        {
            "section": "Who builds it",
            "source": "Team split across architecture, RQA and product",
            "pattern": "three_column_card_slide",
            "assertion": "Three groups share responsibility: architecture, business owners, and deployment.",
            "rationale": "Team roles are easiest to scan as three cards.",
            "inputs": "Team roles and responsibilities.",
            "inference": "No",
        },
        {
            "section": "Where we are",
            "source": "Current prototype and timeline guidance",
            "pattern": "timeline_slide",
            "assertion": "The prototype is working and beta with RQA starts mid 2026 before full production later in 2026.",
            "rationale": "A simple timeline shows current state and next milestones.",
            "inputs": "Current status; beta start; production target.",
            "inference": "No",
        },
        {
            "section": "Why it matters",
            "source": "Governance and audit value of consistency",
            "pattern": "assertion_evidence_slide",
            "assertion": "The deeper value is making scenario work repeatable and auditable, not just faster.",
            "rationale": "A strong closing slide should say why governance matters more than speed.",
            "inputs": "Governance benefit summary; long-term impact line.",
            "inference": "No",
        },
    ])

    b = PptxBuilder(palette="prm")

    three_column_card_slide(
        title="AI MDS — Automated Market Scenario Analysis",
        cards=[
            {
                "tag": "What",
                "heading": "What it is",
                "body": "A system that turns a plain-English market event into a portfolio impact check.",
            },
            {
                "tag": "Who",
                "heading": "Who it's for",
                "body": "Junior analysts and portfolio teams with no prior risk or portfolio modelling background.",
            },
            {
                "tag": "Why",
                "heading": "Why it matters",
                "body": "Faster answers, consistent assumptions, and audit-ready decisions.",
            },
        ],
        bar_color="0D1B2A",
        key_message="A simple, visual opener for a junior analyst audience.",
        builder=b,
    )

    two_column_contrast_slide(
        title="Why the old market shock process is too slow",
        left_panel={
            "heading": "Today",
            "body": (
                "A small market shock can take weeks to analyse.\n"
                "Economists and risk teams debate the scenario, choose drivers, and estimate impacts manually."
            ),
            "accent_color": "0A9396",
        },
        right_panel={
            "heading": "The cost",
            "body": (
                "By the time the answer arrives, the market has already moved.\n"
                "That leaves portfolio managers with outdated information."
            ),
        },
        bar_color="0D1B2A",
        builder=b,
    )

    assertion_evidence_slide(
        title="What AI MDS does",
        assertion="It turns a plain-English market story into a portfolio impact check in minutes.",
        evidence=[
            {
                "text": "It reads the market situation the way a colleague would explain it.",
            },
            {
                "text": "It finds the relevant risk factors and estimates how much each one moves.",
            },
            {
                "text": "It runs those shocks through the portfolio model and shows the expected P&L."
            },
        ],
        key_message="The system is automated, but the user still gets a clear workflow and fast results.",
        builder=b,
    )

    process_slide(
        title="Five steps in the AI MDS pipeline",
        steps=[
            {
                "name": "Read the situation",
                "description": "Turn plain English into a structured scenario.",
            },
            {
                "name": "Check similar history",
                "description": "Find past episodes as a plausibility check.",
            },
            {
                "name": "Pick the factors",
                "description": "Choose the risk drivers the scenario actually moves.",
            },
            {
                "name": "Estimate shocks",
                "description": "Set how much each chosen factor should move.",
            },
            {
                "name": "Calculate impact",
                "description": "Run the shocks through the portfolio model.",
            },
        ],
        highlight=2,
        builder=b,
    )

    three_column_card_slide(
        title="Human oversight is built in",
        cards=[
            {
                "tag": "Review",
                "heading": "Expert review",
                "body": "A risk expert checks the scenario, factor choices, and shock estimates at every step.",
            },
            {
                "tag": "Adjust",
                "heading": "Override when needed",
                "body": "The expert can change the machine output and explain why the change was made.",
            },
            {
                "tag": "Audit",
                "heading": "Record every decision",
                "body": "The system stores overrides and notes, so every decision is visible and auditable.",
            },
        ],
        bar_color="0A9396",
        builder=b,
    )

    three_column_card_slide(
        title="Who is making AI MDS real",
        cards=[
            {
                "tag": "Build",
                "heading": "Architecture & modelling",
                "body": "Design the pipeline, build the statistical models, and make the outputs reliable.",
            },
            {
                "tag": "Use",
                "heading": "RQA business owners",
                "body": "Shape requirements every two weeks and validate the system with real user needs.",
            },
            {
                "tag": "Run",
                "heading": "Analytics Product",
                "body": "Deploy the UI and infrastructure so the system is stable and accessible.",
            },
        ],
        bar_color="0D1B2A",
        builder=b,
    )

    timeline_slide(
        title="Where we are today",
        milestones=[
            {"date": "Now", "label": "Prototype works end to end", "state": "current"},
            {"date": "Mid 2026", "label": "RQA beta testing begins", "state": "future"},
            {"date": "H2 2026", "label": "Full production target", "state": "future"},
        ],
        key_message="The full pipeline and UI are live now; broader testing and production rollout are planned for 2026.",
        builder=b,
    )

    assertion_evidence_slide(
        title="Why governance and audit matter most",
        assertion="A consistent process, documented assumptions, and logged overrides make AI MDS reliable for the business.",
        evidence=[
            {
                "text": "Every scenario goes through the same structured workflow.",
            },
            {
                "text": "Every assumption and override is recorded for later review.",
            },
            {
                "text": "That frees experts to focus on judgement instead of mechanical work.",
            },
        ],
        key_message="Speed is important, but the deeper value is consistency, transparency, and auditability.",
        builder=b,
    )

    output_path = project.output_path("ai_mds_automated_market_scenario_analysis.pptx")
    qa_path = project.qa_path(output_path.name)
    b.save(output_path, final=True, report_path=qa_path)

    log.record_qa(qa_path)
    rendered_qa_path = ArtifactUtilities(project).review_rendered_pptx(output_path)
    log.record_qa(rendered_qa_path)
    log.record_output(output_path)

    project.register_output(
        output_path,
        qa_report_path=qa_path,
        patterns=(
            "two_column_contrast_slide",
            "assertion_evidence_slide",
            "process_slide",
            "three_column_card_slide",
            "timeline_slide",
        ),
        working_files=(
            instructions_path,
            project.working_path("decision_log.md"),
            project.working_path("slide_plan.md"),
        ),
        qa_report_paths=(
            rendered_qa_path,
        ),
        notes="Refined AI MDS explainer deck for a junior analyst audience.",
    )

    print(f"Saved deck to {output_path}")
    print(f"QA report saved to {qa_path}")
    print(f"Rendered review saved to {rendered_qa_path}")


if __name__ == "__main__":
    build_deck()
