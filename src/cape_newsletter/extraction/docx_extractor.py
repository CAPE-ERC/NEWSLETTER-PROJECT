"""Extract text and section boundaries from bulletin .docx files.

Feeds two downstream uses (docs/architecture.md, "Why 2023-to-date data
matters"):
  - style_guide.md derivation (cross-edition phrase/pattern mining)
  - data/trend_db/ population (historical figures per section, per edition)
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

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

_WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _extract_text_via_raw_xml(docx_path: Path) -> str:
    """Fallback for .docx files with a corrupted embedded part (e.g. a bad-CRC
    image) that trips up python-docx's full package load. Reads
    word/document.xml directly out of the zip, which is enough to recover
    paragraph text even when other parts of the archive are damaged.
    """
    with zipfile.ZipFile(docx_path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ElementTree.fromstring(xml_bytes)
    lines = []
    for p_el in root.iter(f"{_WORD_NS}p"):
        line = "".join(t.text or "" for t in p_el.iter(f"{_WORD_NS}t"))
        if line.strip():
            lines.append(line)
    return "\n".join(lines)


def extract_text(docx_path: Path) -> str:
    try:
        doc = Document(docx_path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except (zipfile.BadZipFile, PackageNotFoundError):
        return _extract_text_via_raw_xml(docx_path)


def split_into_sections(text: str) -> dict[str, str]:
    """Split full bulletin text into {heading: body} using SECTION_HEADINGS.

    Headings appear on their own paragraph/line except "Country in Focus",
    which is followed inline by the country name (e.g. "Country in Focus
    –Republic of Benin") in all four reference editions.
    """
    sections: dict[str, str] = {}
    current: str | None = None
    buffer: list[str] = []

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        matched = None
        if line.startswith("Country in Focus"):
            matched = "Country in Focus"
        else:
            normalized = line.replace("’", "'")
            for heading in SECTION_HEADINGS:
                if heading == "Country in Focus":
                    continue
                if normalized == heading.replace("’", "'"):
                    matched = heading
                    break

        if matched:
            if current is not None:
                sections[current] = "\n".join(buffer).strip()
            current = matched
            buffer = [line] if matched == "Country in Focus" else []
        elif current is not None:
            buffer.append(line)

    if current is not None:
        sections[current] = "\n".join(buffer).strip()
    return sections


def iter_bulletins() -> list[Path]:
    return sorted(DATA_RAW_DIR.glob("*.docx"))
