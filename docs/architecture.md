# Architecture

## Decision: RAG + tool-calling agents, not fine-tuning

Fine-tuning is for baking in style/format; it doesn't solve the actual problem here,
which is that the underlying facts (CPI, PMI, FAAC figures) change every month. RAG
solves that directly — numbers come from retrieved/fetched source data, not model
memory, and style consistency is handled with few-shot prompting against the existing
bulletin archive instead of retraining.

## Pipeline

```
Orchestrator
  ├─ Global Data Agent
  ├─ Nigeria Real-Sector Agent
  ├─ Price/CPI Agent
  ├─ Fiscal/FAAC Agent
  └─ Country-Focus Agent
       │  (each returns structured JSON per schemas/section_data_schema.json,
       │   never prose — this keeps fact-gathering separate from writing)
       ▼
  Section Writer Agent(s)
       │  (structured data + prompts/style_guide.md + prompts/sections/*.md)
       ▼
  Chart/Figure Agent
       │  (Figures 1-7 generated from the same JSON, so text and charts can't disagree)
       ▼
  Editor/Consistency Agent
       │  (cross-checks numbers across sections, flags anything not traceable
       │   to the input JSON, produces a diff vs. last month)
       ▼
  Doc Assembly → Human review → publish
```

## Fixed bulletin structure (from March/April/May/July 2026 editions)

1. Highlights
2. Global Economic Update
3. Global Economic Outlook
4. Nigeria's Output Growth (PMI, GDP)
5. Output Growth Outlook
6. Price Update (headline/food/core CPI, NBS)
7. Fiscal Operations Update (FAAC allocations)
8. Conclusion
9. Country in Focus (rotates) + CAPE EPU index

## Data sources

NBS, Stanbic IBTC PMI, FAAC, CBN, CAPE's own EPU index/research, plus global data
(IMF/World Bank/market data) for the global sections.

## Human-in-the-loop

Full autonomy is not appropriate for a numbers-heavy publication — a wrong CPI figure
is a credibility problem. Minimal input, not zero:

- Raw source docs (NBS release, PMI report, FAAC circular) are pulled automatically
  where published on predictable URLs, or pasted in where not.
- The "country in focus" is an editorial choice supplied manually each month.
- A human reviews the near-final draft before publish — non-negotiable.

## Roadmap

- **Phase 1 — Templating & prompts**: convert bulletin archive into a prompt
  library (style rules per section) and finalize the JSON schema each data agent
  must output.
- **Phase 2 — Data connectors**: identify which sources are scrapeable/API-accessible
  on a schedule vs. need manual paste-in; build fetchers (fetch + LLM-extraction
  where there's no clean API).
- **Phase 3 — Agent pipeline**: wire up orchestrator → data agents → writer agents →
  editor agent with tool use; chart agent recreates Figures 1-7 from the shared JSON.
- **Phase 4 — Assembly & QA**: auto-populate the docx template; editor agent outputs
  a diff vs. last month plus a list of unverified claims for human review.
- **Phase 5 — Human-in-the-loop review, then iterate** on prompts/validation rules
  based on what the agents consistently get right/wrong.

## Why 2023-to-date data matters (once the full archive is loaded)

- **Style/voice profile**: real recurring phrases and rhetorical patterns across
  ~36+ editions, not inferred from 4 samples — becomes the actual system prompt.
- **Trend/history database**: lets the writer note things like "fourth consecutive
  month above the 50-point PMI threshold" instead of just restating the current
  figure.
- **Validation/drift detection**: sanity-checks new figures against historical
  ranges to catch transcription errors from source docs before they reach prose.
