"""Editor/consistency agent.

Cross-checks drafted sections against the structured input data: flags any
figure in the prose not traceable to the input JSON, checks MoM/YoY
consistency, and confirms the Highlights section reflects the body. Produces
a diff vs. last month's edition and a list of unverified claims for human
review (see docs/architecture.md, Phase 4).
"""

from __future__ import annotations

from typing import Any


def review_draft(sections: dict[str, str], source_data: dict[str, Any]) -> dict[str, Any]:
    """Return {"flags": [...], "diff_vs_last_month": str} for human review."""
    raise NotImplementedError
