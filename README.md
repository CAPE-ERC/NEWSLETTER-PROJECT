# CAPE Newsletter Automation

Multi-agent pipeline that drafts CAPE's monthly *Economic Performance and Prospect
Bulletin* from minimal manual input, using retrieval-augmented generation rather than
fine-tuning (see [docs/architecture.md](docs/architecture.md) for the full rationale).

## Status

Working with the 4 available editions (March, April, May, July 2026) while the full
2023-to-date archive is pending upload. Structured data has been curated from all 4
bulletins into `data/processed/`, merged into `data/trend_db/trend.json`, and a first
style guide derived from real cross-edition patterns. The **Price Update** section has
a complete prompt template and a working writer agent wired to the Anthropic API — the
other sections' templates are still Phase 1 placeholders. A local dashboard
(`web/`) visualizes the trend data, lets you inspect each edition's extracted figures,
and test draft generation.

## Layout

```
data/
  raw/bulletins/     source .docx editions (style exemplars + trend data)
  processed/         curated structured data per edition, matches schemas/section_data_schema.json
  trend_db/          trend.json (merged monthly series) + data_notes.md (source inconsistencies found)
prompts/
  style_guide.md     house style rules, derived from the 4 available bulletins
  sections/          one prompt template per bulletin section (price_update.md is complete)
schemas/
  section_data_schema.json   JSON schema every data-collection agent must satisfy
src/cape_newsletter/
  agents/            orchestrator, data agents, writer agents (real Anthropic API call), editor, chart agent
  extraction/        docx -> text extraction, incl. a raw-XML fallback for the corrupted May 2026 file
  api.py             FastAPI backend for the local dashboard
web/                 static dashboard frontend (trend charts, edition explorer, draft-generation test UI)
docs/
  architecture.md    pipeline design, decision rationale, roadmap
```

## Setup

```
pip install -e ".[dev]"
cp .env.example .env   # fill in ANTHROPIC_API_KEY
```

## Run the dashboard

```
uvicorn cape_newsletter.api:app --app-dir src --reload
```

Then open http://127.0.0.1:8000. The trend charts and edition explorer work without
an API key; "Generate draft" (Price Update only, for now) calls the Anthropic API and
needs `ANTHROPIC_API_KEY` set in `.env`.

## Run tests

```
pytest
```

## Next steps

1. Load the full 2023-to-date bulletin archive into `data/raw/bulletins/` — expands
   the trend database and lets the style guide be derived from real patterns across
   ~36 editions instead of 4.
2. Fill in the remaining prompt templates in `prompts/sections/` (currently only
   `price_update.md` is complete) using the same approach as the reference example.
3. Build data connectors for NBS/PMI/FAAC sources (Phase 2 in the roadmap).
4. Wire up the orchestrator, editor, and chart agents (currently stubs).
