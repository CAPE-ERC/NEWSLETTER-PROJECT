import json

from cape_newsletter.config import DATA_RAW_DIR, SCHEMAS_DIR
from cape_newsletter.extraction.docx_extractor import extract_text, iter_bulletins


def test_bulletins_present():
    bulletins = iter_bulletins()
    assert len(bulletins) == 4


def test_bulletin_text_extracts():
    bulletins = iter_bulletins()
    text = extract_text(bulletins[0])
    assert len(text) > 0


def test_schema_is_valid_json():
    schema_path = SCHEMAS_DIR / "section_data_schema.json"
    with open(schema_path, encoding="utf-8") as f:
        schema = json.load(f)
    assert schema["title"] == "CAPE Bulletin Monthly Data"
