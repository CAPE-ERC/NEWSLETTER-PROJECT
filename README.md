# CAPE Newsletter Automation

Multi-agent pipeline that drafts CAPE's monthly *Economic Performance and Prospect
Bulletin* from minimal manual input, using retrieval-augmented generation rather than
fine-tuning (see [docs/architecture.md](docs/architecture.md) for the full rationale).

## Status

Scaffolding stage. Four bulletins (March, April, May, July 2026) are in
`data/raw/bulletins/`; the full 2023-to-date archive is pending upload before the
style guide and trend database can be built from real historical patterns.

## Layout

```
data/
  raw/bulletins/     source .docx editions (style exemplars + trend data)
  processed/         extracted structured data per edition (generated)
  trend_db/          historical time series across editions (generated)
prompts/
  style_guide.md     house style rules, extracted from the bulletin archive
  sections/          one prompt template per bulletin section
schemas/
  section_data_schema.json   JSON schema every data-collection agent must satisfy
src/cape_newsletter/
  agents/            orchestrator, data agents, writer agents, editor, chart agent
  extraction/         docx -> structured data extraction utilities
docs/
  architecture.md    pipeline design, decision rationale, roadmap
```

## Setup

```
pip install -e ".[dev]"
cp .env.example .env   # fill in ANTHROPIC_API_KEY
```

## Next steps

1. Load the full 2023-to-date bulletin archive into `data/raw/bulletins/`.
2. Run extraction to populate `data/processed/` and `data/trend_db/`.
3. Derive `prompts/style_guide.md` from real cross-edition patterns.
4. Build data connectors for NBS/PMI/FAAC sources (Phase 2 in the roadmap).
