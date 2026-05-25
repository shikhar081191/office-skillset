# Deck Brief Template

**How to use:**
1. Copy the blank brief below and fill it in (takes ~2 minutes)
2. Paste the filled brief into Copilot or Windsurf chat
3. Add your raw content below it — bullet points, notes, data, whatever you have
4. The AI reads `AI_INSTRUCTIONS.md` and `STORY_TEMPLATES.md`, picks the right
   template, and builds the deck

---

## BLANK BRIEF — copy and fill this in

```
DECK PURPOSE:       [approval request / project explainer / exec update / research findings / status update]
AUDIENCE:           [who will see this — role and expertise level]
KEY MESSAGE:        [the one thing they must leave knowing or deciding — one sentence]
DECISION NEEDED:    [what you need from them, if anything — or "none"]
MAX SLIDES:         [number, or "use template default"]
PALETTE:            [prm / blackrock / midnight_executive / charcoal_minimal — or "default"]

TOPICS TO COVER:
  - [Topic or section 1]
  - [Topic or section 2]
  - [Topic or section 3]

KEY DATA POINTS:
  - [Metric: value — context, e.g. "PMSE: 18% improvement vs baseline"]
  - [Metric: value — context]
  - [Metric: value — context]

TIMELINE MILESTONES:    [dates and labels, or "none"]
TEAM / OWNERS:          [names, roles — or "none"]
RISKS OR CAVEATS:       [anything the audience will push back on — or "none"]

NOTES:
[Paste your raw thinking here — paragraphs, bullet points, data tables, anything.
The AI will distill this into the deck. You do not need to structure it.]
```

---

## EXAMPLE — filled brief for an approval request

```
DECK PURPOSE:       approval request
AUDIENCE:           Risk Committee — senior internal, knows the model domain
KEY MESSAGE:        Model v3 is ready for production; approve deployment by end of June
DECISION NEEDED:    Sign-off on production deployment of Model v3
MAX SLIDES:         6
PALETTE:            prm

TOPICS TO COVER:
  - Recommendation and ask
  - Performance improvement vs current model
  - Sensitivity analysis across asset cohorts
  - Risks and mitigations
  - Deployment timeline

KEY DATA POINTS:
  - PMSE improvement: 18.2% on holdout data vs current production model
  - Statistical significance: p < 0.01 across all three asset cohorts
  - Runtime: 6.4 minutes end-to-end, within the 10-minute SLA
  - Cohorts validated: Rates, Credit, Equities

TIMELINE MILESTONES:
  - Jan 2026: Development complete (past)
  - Mar 2026: Backtesting and validation (past)
  - May 2026: Risk Committee approval (current)
  - Jun 2026: Shadow running begins (future)
  - Aug 2026: Full production deployment (future)

TEAM / OWNERS:      none (committee deck, no team slide needed)
RISKS OR CAVEATS:
  - Model has not been tested on FX cohort — excluded from v3 scope
  - Legal sign-off on data lineage required before go-live

NOTES:
The main improvement comes from a new factor decomposition for short-duration
sovereigns (bills). Bills drove 46pp of the PMSE gain. The credit cohort also
improved materially (+22pp) but equities was flat vs prior model.

Sensitivity: the improvement holds under +/-1 sigma rate shocks and a 2008-style
credit stress scenario. The only scenario where v3 underperforms is a rapid FX
dislocation, which is why we excluded FX from v3 scope.

The committee previously asked about explainability. We added a factor attribution
output that breaks PMSE by risk factor — this satisfies the governance requirement.
```
