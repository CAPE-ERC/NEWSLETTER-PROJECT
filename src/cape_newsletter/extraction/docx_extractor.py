"""Extract text and section boundaries from bulletin .docx files.

Feeds two downstream uses (docs/architecture.md, "Why 2023-to-date data
matters"):
  - style_guide.md derivation (cross-edition phrase/pattern mining)
  - data/trend_db/ population (historical figures per section, per edition)
"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from cape_newsletter.config import DATA_RAW_DIR

SECTION_HEADINGS = [
    "Highlights",
    "Global Economic Update",
    "Global Economic Outlook",
    "Nigeria's Output Growth",
    "Output Growth Outlook",
    "Price Update",
    "Fiscal Operations Update",
    "Conclusion",
    "Country in Focus",
]


def extract_text(docx_path: Path) -> str:
    doc = Document(docx_path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def split_into_sections(text: str) -> dict[str, str]:
    """Split full bulletin text into {heading: body} using SECTION_HEADINGS.

    Phase 1 work: headings may vary slightly across editions (e.g. numbering,
    punctuation) — validate against the full archive once available.
    """
    raise NotImplementedError


def iter_bulletins() -> list[Path]:
    return sorted(DATA_RAW_DIR.glob("*.docx"))
