"""Section writer agents.

Each function drafts one bulletin section from structured data (never raw
source docs) plus the style guide and prompt template for that section. See
prompts/sections/ for the per-section prompt templates.
"""

from __future__ import annotations

from typing import Any

from cape_newsletter.config import PROMPTS_DIR


def write_section(section_name: str, data: dict[str, Any]) -> str:
    """Draft one section's prose from its structured data.

    Loads prompts/sections/<section_name>.md, injects `data` plus the shared
    style guide, and calls the model. Phase 3 work.
    """
    prompt_path = PROMPTS_DIR / "sections" / f"{section_name}.md"
    if not prompt_path.exists():
        raise FileNotFoundError(f"No prompt template for section '{section_name}'")
    raise NotImplementedError
