from artifact_utilities import ArtifactUtilities
from decision_log import DecisionLog
from patterns import (
    assertion_evidence_slide,
    callout_bar_slide,
    numbered_steps_slide,
    status_slide,
    three_column_card_slide,
    timeline_slide,
    two_column_contrast_slide,
)
from create_pptx import PptxBuilder, PALETTES
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
            "pattern": "title_slide",
            "assertion": "Introduce AI MDS in simple language and state the deck purpose.",
            "rationale": "A junior analyst needs a clear opening statement and context.",
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
            "pattern": "numbered_steps_slide",
            "assertion": "The system works through five ordered steps from reading the event to calculating impact.",
            "rationale": "numbered_steps_slide is the best fit for a sequential vertical pipeline — process_slide is horizontal and cramped at 5 steps.",
            "inputs": "The five workflow steps and short descriptions.",
            "inference": "No",
        },
        {
            "section": "Human controls",
            "source": "Human-in-the-loop safeguard description",
            "pattern": "three_column_card_slide",
            "assertion": "Experts review, override, and document every step so the system is not a black box.",
            "rationale": "Three cards separate review, overrides, and audit clearly — genuinely three parallel equal-weight items.",
            "inputs": "Review points; override mechanism; audit value.",
            "inference": "No",
        },
        {
            "section": "Who builds it",
            "source": "Team split across architecture, RQA and product",
            "pattern": "status_slide",
            "assertion": "Three groups share responsibility: architecture, business owners, and deployment.",
            "rationale": "status_slide avoids repeating three_column_card_slide back-to-back and suits team/ownership content naturally.",
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
            "pattern": "callout_bar_slide",
            "assertion": "The deeper value is making scenario work repeatable and auditable, not just faster.",
            "rationale": "callout_bar_slide is the correct closing pattern — one strong statement plus supporting bullets.",
            "inputs": "Governance benefit summary; long-term impact line.",
            "inference": "No",
        },
    ])

    prm = PALETTES["prm"]
    b = PptxBuilder(palette="prm")

    # ── 1. Title ──────────────────────────────────────────────────────────────
    b.title_slide(
        "AI MDS — Automated Market Scenario Analysis",
        "Junior Analyst Briefing · 2026",
    )

    # ── 2. The problem ────────────────────────────────────────────────────────
    two_column_contrast_slide(
        title="Why the old market shock process is too slow",
        left_panel={
            "heading": "Today",
            "body": (
                "A small market shock can take weeks to analyse.\n"
                "Economists and risk teams debate the scenario, choose drivers, and estimate impacts manually."
            ),
            "accent_color": prm["primary"],
        },
        right_panel={
            "heading": "The cost",
            "body": (
                "By the time the answer arrives, the market has already moved.\n"
                "That leaves portfolio managers with outdated information."
            ),
        },
        bar_color=prm["primary"],
        key_message="Manual scenario analysis is too slow to be actionable when markets move fast.",
        builder=b,
    )

    # ── 3. The solution ───────────────────────────────────────────────────────
    assertion_evidence_slide(
        title="What AI MDS does",
        assertion="It turns a plain-English market story into a portfolio impact check in minutes.",
        evidence=[
            {"text": "It reads the market situation the way a colleague would explain it."},
            {"text": "It finds the relevant risk factors and estimates how much each one moves."},
            {"text": "It runs those shocks through the portfolio model and shows the expected P&L."},
        ],
        key_message="The system is automated, but the user stays in control at every step.",
        builder=b,
    )

    # ── 4. Pipeline ───────────────────────────────────────────────────────────
    numbered_steps_slide(
        title="Five steps in the AI MDS pipeline",
        steps=[
            {"title": "Read the situation",   "description": "Turn plain English into a structured scenario."},
            {"title": "Check similar history", "description": "Find past episodes as a plausibility check."},
            {"title": "Pick the factors",      "description": "Choose the risk drivers the scenario actually moves."},
            {"title": "Estimate shocks",       "description": "Set how much each chosen factor should move."},
            {"title": "Calculate impact",      "description": "Run the shocks through the portfolio model."},
        ],
        bar_color=prm["primary"],
        key_message="Step 3 — identifying the right risk factors — is the most critical; wrong factors make everything downstream wrong.",
        builder=b,
    )

    # ── 5. Human controls ─────────────────────────────────────────────────────
    three_column_card_slide(
        title="Human oversight is built in at every step",
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
        bar_color=prm["secondary"],
        key_message="This is not a black box — every decision is visible, logged, and explainable.",
        builder=b,
    )

    # ── 6. Team ───────────────────────────────────────────────────────────────
    status_slide(
        title="Who is building AI MDS",
        workstreams=[
            {
                "name": "Architecture & Modelling",
                "status": "Pipeline design, statistical models, and output reliability",
                "rag": "green",
                "owner": "Your team",
            },
            {
                "name": "RQA — Business Owners",
                "status": "Shape requirements every two weeks; validate with real use cases",
                "rag": "green",
                "owner": "RQA",
            },
            {
                "name": "Analytics Product",
                "status": "UI deployment and infrastructure; keeps the system stable and accessible",
                "rag": "green",
                "owner": "Analytics Product",
            },
        ],
        key_message="Three teams, one shared cadence — requirements are reviewed and validated every two weeks.",
        builder=b,
    )

    # ── 7. Timeline ───────────────────────────────────────────────────────────
    timeline_slide(
        title="Where we are today",
        milestones=[
            {"date": "Now",      "label": "Prototype works end to end; UI is live", "state": "current"},
            {"date": "Mid 2026", "label": "RQA beta testing begins",                "state": "future"},
            {"date": "H2 2026",  "label": "Full production target",                 "state": "future"},
        ],
        key_message="The full pipeline and UI are live now; broader testing and production rollout are planned for 2026.",
        builder=b,
    )

    # ── 8. Why it matters ─────────────────────────────────────────────────────
    callout_bar_slide(
        title="Why governance and audit matter most",
        body=[
            "Every scenario follows the same structured process — no shortcuts, no ad hoc shortcuts.",
            "Every assumption is documented. Every override is logged with a reason.",
            "That frees experts to focus on the hard judgement calls instead of mechanical construction.",
        ],
        callout_text=(
            "Speed is the headline. Consistency, transparency, and auditability are the deeper value."
        ),
        bar_color=prm["secondary"],
        builder=b,
    )

    # ── Save, QA, rendered review, output registration ────────────────────────
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
            "title_slide",
            "two_column_contrast_slide",
            "assertion_evidence_slide",
            "numbered_steps_slide",
            "three_column_card_slide",
            "status_slide",
            "timeline_slide",
            "callout_bar_slide",
        ),
        working_files=(
            instructions_path,
            project.working_path("decision_log.md"),
            project.working_path("slide_plan.md"),
        ),
        qa_report_paths=(rendered_qa_path,),
        notes="Refined AI MDS explainer deck for a junior analyst audience.",
    )

    print(f"Saved deck to {output_path}")
    print(f"QA report saved to {qa_path}")
    print(f"Rendered review saved to {rendered_qa_path}")


if __name__ == "__main__":
    build_deck()
