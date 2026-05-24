# Executive Summary

Source: projects\q2_model_review\inputs\q2_model_review_brief.docx

## Source notes
- Q2 2025 Credit Model Review — Decision Briefing
- For: Credit Risk Committee  |  24 May 2025  |  Prepared by: Credit Risk Modelling Team

## Executive Summary
<!-- DISTIL: rewrite long paragraphs below into concise slide copy before using -->
- [DISTIL] The Credit Risk Modelling team recommends approval of Credit Model v3 for full production deployment in Q3 2025. The new model delivers an 18 percentage point improvement in PMSE on the holdout validation set, passes all regulatory stress tests, and has received independent model validation sign-off from Accenture Risk Advisory on 14 May 2025. A 90-day parallel run against live book positions confirms no material change to capital requirements. The primary residual risk is vintage-level drift in the short-duration cohort, mitigated by a monthly monitoring protocol agreed with the model risk team.
- Management recommendation: Approve Credit Model v3. Target regulatory submission 28 May and go-live August 2025.

## Agenda
- Today's committee meeting covers five areas.
- 1. Executive Summary — Recommendation and key findings
- 2. Model Performance — Quantitative results and holdout validation
- 3. Stress Testing and Assumptions — Macro scenario parameters
- 4. Implementation Plan — Go-live timeline and resource requirements
- 5. Appendix — Methodology and supporting data

## Q2 Performance Snapshot
- Six key metrics summarise Q2 2025 model and portfolio performance vs prior quarter.
- PMSE (holdout, Model v3): 0.046 — down from 0.060 baseline, improvement of 23%
- Sharpe Ratio: 1.18 — up 0.08 vs Q1 2025
- AUM covered: $4.7 billion
- Model coverage: 98.3% of live positions
- Information Ratio: 0.71
- Stress test pass rate: 100% across all macro scenarios

## Programme Status
<!-- DISTIL: rewrite long paragraphs below into concise slide copy before using -->
- Six workstreams are tracked on the path to production deployment.
- [DISTIL] Model Development: Complete. Credit Model v3 signed off by internal validation on 28 April 2025. All factor specifications finalised.
- [DISTIL] Independent Validation: Complete. External validator Accenture confirmed sign-off 14 May 2025 following four-week review.
- Regulatory Submission: In progress. MRM submission package drafted; awaiting legal review sign-off expected 28 May.
- [DISTIL] Technology Integration: At risk. Two-week delay caused by API versioning incompatibility with core banking system. Vendor patch in testing.
- Parallel Run: Complete. 90-day parallel run ended 10 May. All cohort results within agreed tolerance bands.
- Documentation: In progress. Model card and user guide 80% complete. Target sign-off 7 June ahead of go-live.

## Model Performance Comparison
- PMSE by model version and maturity cohort on the out-of-time holdout set (January–March 2025).

### table_1
Cohort | Model v1 | Model v2 | Model v3
--- | --- | --- | ---
Short (<1Y) | 0.078 | 0.048 | 0.031
Medium (1-5Y) | 0.085 | 0.061 | 0.049
Long (>5Y) | 0.091 | 0.072 | 0.058
All Cohorts | 0.084 | 0.060 | 0.046
Numeric columns detected: Model v1, Model v2, Model v3

## Implementation Plan
- The implementation follows a six-milestone schedule from model completion to full production go-live.
- February 2025: Model development complete — all factor specifications finalised and version-controlled.
- March 2025: Internal validation sign-off — passed all internal Model Risk Management criteria.
- May 2025: External validation sign-off — Accenture confirmed; current milestone reached.
- June 2025: Regulatory submission — MRM package submitted to PRA under SS1/23 requirements.
- July 2025: Technology integration — API migration, UAT and user acceptance sign-off complete.
- August 2025: Production go-live — Model v3 active across all credit portfolios.

## Stress Test Assumptions
<!-- DISTIL: rewrite long paragraphs below into concise slide copy before using -->
- [DISTIL] Model validated against a 1-in-20 adverse macro scenario. Assumptions represent a coherent stress, not individual extremes.

### table_2
Assumption | Value | Source | Sensitivity
--- | --- | --- | ---
GDP growth (2025) | -1.5% | IMF WEO | High
Rate path (10Y UST) | +75bps | Fed dots | High
IG spread widening | +120bps | Internal | Medium
Equity drawdown | -20% | Internal | High
FX (EURUSD) | 0.95 | Bloomberg | Low
Recovery rate | 40% | S&P LossStats | Medium
Numeric columns detected: Value

## Factor Return Attribution
- Q2 2025 return attribution by risk factor, comparing Model v2 and Model v3 contributions in basis points.

### table_3
Factor | Model v2 (bps) | Model v3 (bps)
--- | --- | ---
Equity Beta | 28 | 32
Duration | -12 | -8
Credit | 15 | 18
FX Carry | 6 | 6
Residual | -4 | -2
Numeric columns detected: Model v2 (bps), Model v3 (bps)

## Historical PMSE Trend
- Rolling quarterly PMSE for Model v2 and Model v3 over the 2023–2025 parallel run period. Lower is better.

### table_4
Quarter | Model v2 | Model v3
--- | --- | ---
Q3 23 | 0.065 | 0.058
Q4 23 | 0.063 | 0.055
Q1 24 | 0.061 | 0.052
Q2 24 | 0.060 | 0.049
Q3 24 | 0.062 | 0.048
Q4 24 | 0.061 | 0.047
Q1 25 | 0.060 | 0.047
Q2 25 | 0.060 | 0.046
Numeric columns detected: Model v2, Model v3

## Core Finding
<!-- DISTIL: rewrite long paragraphs below into concise slide copy before using -->
- [DISTIL] Bills and short-duration instruments drive almost all of the PMSE improvement. The model's treatment of prepayment risk and vintage effects in the short-duration cohort explains 72% of the total PMSE gain — equivalent to 14 percentage points of the 18pp total improvement. The remaining 4pp comes from improved credit spread factor specification in the medium and long maturity cohorts. External validation using an independent data panel (January 2020 – December 2024) confirms the finding is robust to different time periods.
- Short cohort PMSE reduced by 35%: from 0.048 to 0.031
- Vintage effect accounts for 72% of the total PMSE gain across all cohorts
- External validation on independent panel confirms finding holds out-of-sample
- Stress test capital impact: immaterial — within tolerance agreed with PRA

## Implementation Risk Prioritisation
<!-- DISTIL: rewrite long paragraphs below into concise slide copy before using -->
- Four risks require committee attention, assessed by potential impact and resolution urgency.
- [DISTIL] API versioning delay: High impact, High urgency. Two-week delay on technology integration. Extended parallel run covers any gap.
- [DISTIL] Vintage drift monitoring: High impact, Low urgency. Post-launch monthly monitoring protocol agreed. Recalibration triggers if 5% threshold breached.
- User training coverage: Medium impact, Medium urgency. Three workshops scheduled June. All users trained before go-live.
- [DISTIL] Documentation completeness: Low impact, Low urgency. Model card 80% done. No blocker to go-live; can complete post-launch.

## Candidate slide visuals
- table_1: comparison table, heat map or diverging bars using Model v1, Model v2, Model v3
- table_2: bar chart or ranked result view using Value
- table_3: comparison table, heat map or diverging bars using Model v2 (bps), Model v3 (bps)
- table_4: comparison table, heat map or diverging bars using Model v2, Model v3
