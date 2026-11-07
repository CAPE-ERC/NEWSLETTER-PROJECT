"""Run the project-level GDP nowcast and newsletter workflow from this folder."""
from pathlib import Path
import runpy
import sys


if __name__ == "__main__":
    sys.dont_write_bytecode = True
    launcher = Path(__file__).resolve().parent.parent / "GDP_Nowcast_and_Newsletter.py"
    if not launcher.is_file():
        raise FileNotFoundError(f"Project launcher not found: {launcher}")
    runpy.run_path(str(launcher), run_name="__main__")
