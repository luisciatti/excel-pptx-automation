from pathlib import Path
import sys

PACKAGE_DIR = Path(__file__).resolve().parent
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = PACKAGE_DIR

DATA_DIR = BASE_DIR / "data"
INPUTS_DIR = DATA_DIR / "inputs"
PPTX_DIR = DATA_DIR / "pptx"
LOGS_DIR = DATA_DIR / "logs"

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_PPTX_DIR = OUTPUT_DIR / "pptx"
OUTPUT_REPORTS_DIR = OUTPUT_DIR / "reports"

ALL_DATA_DIRS = [INPUTS_DIR, PPTX_DIR, LOGS_DIR, OUTPUT_PPTX_DIR, OUTPUT_REPORTS_DIR]


def ensure_directories():
    for folder in ALL_DATA_DIRS:
        folder.mkdir(parents=True, exist_ok=True)
