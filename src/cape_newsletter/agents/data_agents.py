"""Data-collection agents.

Each agent fetches one section's facts for the current month and returns
structured data matching schemas/section_data_schema.json — never prose. This
keeps fact-gathering separate from writing so the writer agents can't invent
numbers (see docs/architecture.md).

Phase 2 work: wire each function to its real source (web fetch + LLM
extraction for PDFs/press releases, or a manual paste-in fallback).
"""

from __future__ import annotations

from typing import Any


def fetch_global_data() -> dict[str, Any]:
    """IMF/World Bank/market data for the Global Economic Update/Outlook sections."""
    raise NotImplementedError


def fetch_nigeria_output_data() -> dict[str, Any]:
    """Stanbic IBTC PMI + NBS GDP for Nigeria's Output Growth section."""
    raise NotImplementedError


def fetch_price_data() -> dict[str, Any]:
    """NBS headline/food/core CPI for the Price Update section."""
    raise NotImplementedError


def fetch_fiscal_data() -> dict[str, Any]:
    """FAAC allocation figures for the Fiscal Operations Update section."""
    raise NotImplementedError


def fetch_country_in_focus_data(country: str) -> dict[str, Any]:
    """CAPE EPU index + notes for the manually-chosen country in focus."""
    raise NotImplementedError
