from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = PROJECT_ROOT / "data" / "raw" / "30120 RF.pdf"
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "ocr"

MODEL_NAME = "gpt-4.1"

SUBTESTS_TO_EXTRACT = [
    "Similarities",
    "Vocabulary",
    "Comprehension",
]

OUTPUT_JSON = OUTPUT_DIR / "AISI_captured_protocol.json"
OUTPUT_EXCEL = OUTPUT_DIR / "AISI_captured_protocol.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "AISI_captured_protocol.csv"

ENV_FILE = PROJECT_ROOT / ".env"
