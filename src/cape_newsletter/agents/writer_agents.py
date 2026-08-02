"""Section writer agents.

Each function drafts one bulletin section from structured data (never raw
source docs) plus the style guide and prompt template for that section. See
prompts/sections/ for the per-section prompt templates.
"""

from __future__ import annotations

import json
from typing import Any

import anthropic

from cape_newsletter.config import PROMPTS_DIR

MODEL = "claude-opus-5"


def write_section(section_name: str, data: dict[str, Any]) -> str:
    """Draft one section's prose from its structured data.

    Loads prompts/sections/<section_name>.md, injects `data` plus the shared
    style guide, and calls the model. The prompt template instructs the model
    to never state a figure absent from `data` — this is the guard against
    invented numbers described in docs/architecture.md.
    """
    prompt_path = PROMPTS_DIR / "sections" / f"{section_name}.md"
    if not prompt_path.exists():
        raise FileNotFoundError(f"No prompt template for section '{section_name}'")
    section_prompt = prompt_path.read_text(encoding="utf-8")
    style_guide = (PROMPTS_DIR / "style_guide.md").read_text(encoding="utf-8")

    system = f"{section_prompt}\n\n---\n\n{style_guide}"

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=MODEL,
        max_tokens=2000,
        system=system,
        output_config={"effort": "medium"},
        messages=[
            {
                "role": "user",
                "content": (
                    "Draft this section from the following structured data. "
                    "Do not state any figure that is not present in this JSON:\n\n"
                    f"{json.dumps(data, indent=2)}"
                ),
            }
        ],
    )
    return next(block.text for block in response.content if block.type == "text")
