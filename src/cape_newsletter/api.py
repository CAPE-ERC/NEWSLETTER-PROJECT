"""FastAPI backend for the local test dashboard.

Serves the curated edition data and trend series read-only, plus a
/api/generate endpoint that exercises the real writer agent (agents/writer_agents.py)
against the Anthropic API for sections that already have a real prompt
template (see prompts/sections/). This is a local development/testing surface,
not the production pipeline described in docs/architecture.md.
"""

from __future__ import annotations

import json

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from cape_newsletter.agents import writer_agents
from cape_newsletter.config import DATA_PROCESSED_DIR, ROOT_DIR, TREND_DB_DIR

app = FastAPI(title="CAPE Newsletter Dashboard API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Sections with a real, filled-in prompt template (see prompts/sections/).
# The rest are still Phase 1 placeholders — listed so the dashboard can show
# them as "not ready" rather than silently generating from an empty template.
READY_SECTIONS = {"price_update"}

# Maps a section name to the key in a processed edition's JSON that holds its
# structured data, per schemas/section_data_schema.json.
SECTION_DATA_KEY = {
    "global_economic_update": "global",
    "global_economic_outlook": "global",
    "nigeria_output_growth": "nigeria_output",
    "output_growth_outlook": "nigeria_output",
    "price_update": "prices",
    "fiscal_operations": "fiscal",
    "country_in_focus": "country_in_focus",
}


def _edition_id(path) -> str:
    return path.stem  # e.g. "2026-07"


@app.get("/api/editions")
def list_editions():
    editions = []
    for path in sorted(DATA_PROCESSED_DIR.glob("*.json")):
        record = json.loads(path.read_text(encoding="utf-8"))
        editions.append(
            {
                "id": _edition_id(path),
                "month": record["edition"]["month"],
                "year": record["edition"]["year"],
                "country_in_focus": record["country_in_focus"]["country"],
            }
        )
    return editions


@app.get("/api/editions/{edition_id}")
def get_edition(edition_id: str):
    path = DATA_PROCESSED_DIR / f"{edition_id}.json"
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"No edition '{edition_id}'")
    return json.loads(path.read_text(encoding="utf-8"))


@app.get("/api/trend")
def get_trend():
    path = TREND_DB_DIR / "trend.json"
    if not path.exists():
        raise HTTPException(status_code=404, detail="trend.json not found")
    return json.loads(path.read_text(encoding="utf-8"))


@app.get("/api/sections")
def list_sections():
    return {
        "ready": sorted(READY_SECTIONS),
        "all": sorted(SECTION_DATA_KEY),
    }


class GenerateRequest(BaseModel):
    edition_id: str
    section: str


@app.post("/api/generate")
def generate_section(req: GenerateRequest):
    if req.section not in SECTION_DATA_KEY:
        raise HTTPException(status_code=400, detail=f"Unknown section '{req.section}'")
    if req.section not in READY_SECTIONS:
        raise HTTPException(
            status_code=409,
            detail=(
                f"prompts/sections/{req.section}.md is still a Phase 1 placeholder — "
                "only price_update has a real prompt template so far."
            ),
        )

    path = DATA_PROCESSED_DIR / f"{req.edition_id}.json"
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"No edition '{req.edition_id}'")
    record = json.loads(path.read_text(encoding="utf-8"))
    section_data = record[SECTION_DATA_KEY[req.section]]

    try:
        text = writer_agents.write_section(req.section, section_data)
    except Exception as exc:  # anthropic.AuthenticationError, missing key, etc.
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return {"section": req.section, "edition_id": req.edition_id, "text": text}


WEB_DIR = ROOT_DIR / "web"
if WEB_DIR.exists():
    app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
