import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

ROOT_DIR = Path(__file__).resolve().parents[2]
DATA_RAW_DIR = ROOT_DIR / "data" / "raw" / "bulletins"
DATA_PROCESSED_DIR = ROOT_DIR / "data" / "processed"
TREND_DB_DIR = ROOT_DIR / "data" / "trend_db"
PROMPTS_DIR = ROOT_DIR / "prompts"
SCHEMAS_DIR = ROOT_DIR / "schemas"
