"""Chart/figure agent.

Regenerates Figures 1-7 from the same structured JSON used by the writer
agents, so numbers in text and charts can never disagree.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any


def generate_figures(source_data: dict[str, Any], output_dir: Path) -> list[Path]:
    """Render each figure to output_dir and return the file paths."""
    raise NotImplementedError
