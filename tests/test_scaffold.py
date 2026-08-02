import json

from cape_newsletter.config import DATA_PROCESSED_DIR, SCHEMAS_DIR
from cape_newsletter.extraction.docx_extractor import (
    SECTION_HEADINGS,
    extract_text,
    iter_bulletins,
    split_into_sections,
)


def test_bulletins_present():
    bulletins = iter_bulletins()
    assert len(bulletins) == 4


def test_all_bulletins_extract_text_including_corrupted_may_edition():
    for bulletin in iter_bulletins():
        text = extract_text(bulletin)
        assert len(text) > 0, f"{bulletin.name} produced no text"


def test_split_into_sections_finds_every_heading():
    for bulletin in iter_bulletins():
        text = extract_text(bulletin)
        sections = split_into_sections(text)
        for heading in SECTION_HEADINGS:
            assert heading in sections, f"{bulletin.name} missing '{heading}'"
            assert len(sections[heading]) > 0


def test_schema_is_valid_json():
    schema_path = SCHEMAS_DIR / "section_data_schema.json"
    with open(schema_path, encoding="utf-8") as f:
        schema = json.load(f)
    assert schema["title"] == "CAPE Bulletin Monthly Data"


def test_processed_data_matches_schema_required_fields():
    schema_path = SCHEMAS_DIR / "section_data_schema.json"
    with open(schema_path, encoding="utf-8") as f:
        schema = json.load(f)
    required = schema["required"]

    processed_files = sorted(DATA_PROCESSED_DIR.glob("*.json"))
    assert len(processed_files) == 4
    for path in processed_files:
        with open(path, encoding="utf-8") as f:
            record = json.load(f)
        for field in required:
            assert field in record, f"{path.name} missing required field '{field}'"
