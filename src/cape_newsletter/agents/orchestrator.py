"""Orchestrator: runs the full monthly pipeline.

data agents -> writer agents -> chart agent -> editor agent -> doc assembly.
See docs/architecture.md for the pipeline diagram and rationale.
"""

from __future__ import annotations

from cape_newsletter.agents import chart_agent, data_agents, editor_agent, writer_agents

SECTIONS = [
    "highlights",
    "global_economic_update",
    "global_economic_outlook",
    "nigeria_output_growth",
    "output_growth_outlook",
    "price_update",
    "fiscal_operations",
    "conclusion",
    "country_in_focus",
]


def run(country_in_focus: str) -> None:
    """Produce a draft bulletin for the current month.

    Phase 3 work: wire data_agents -> writer_agents.write_section per section
    in SECTIONS -> chart_agent.generate_figures -> editor_agent.review_draft
    -> hand off to doc assembly for human review.
    """
    raise NotImplementedError
