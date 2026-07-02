from pathlib import Path
import base64
import json
import os

import pandas as pd
from openai import OpenAI


# =====================================================
# CONFIG
# =====================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

IMAGE_FILE = PROJECT_ROOT / "data" / "raw" / "fake_data_wisc5_protocol.png"

OUTPUT_DIR = PROJECT_ROOT / "outputs" / "ocr"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

JSON_OUT = OUTPUT_DIR / "pass_protocol_extraction.json"
EXCEL_OUT = OUTPUT_DIR / "pass_protocol_extraction.xlsx"
CSV_OUT = OUTPUT_DIR / "pass_protocol_extraction.csv"


# =====================================================
# ENV
# =====================================================

def load_env_file(env_path: Path | None = None) -> None:
    if env_path is None:
        env_path = PROJECT_ROOT / ".env"

    if not env_path.exists():
        raise FileNotFoundError(f"Could not find .env at {env_path}")

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()

        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def encode_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


# =====================================================
# EXTRACTION
# =====================================================

def extract_protocol_data(image_file: Path) -> dict:
    if not image_file.exists():
        raise FileNotFoundError(image_file)

    load_env_file()
    client = OpenAI()

    image_b64 = encode_image(image_file)

    response = client.responses.create(
        model="gpt-4.1",
        input=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": """
You are extracting data from a fake administered cognitive assessment protocol.

Extract both printed and handwritten information.

Return JSON with this structure:
{
  "examinee_info": {
    "candidate_id": "",
    "test_date": "",
    "birth_date": "",
    "test_age": "",
    "examiner_name": "",
    "examiner_id": "",
    "site_id": ""
  },
  "subtests": [
    {
      "subtest_name": "",
      "items": [
        {
          "item_number": "",
          "printed_prompt": "",
          "handwritten_response": "",
          "dk_marked": false,
          "nr_marked": false,
          "score_marked": "",
          "confidence": 0.0,
          "notes": ""
        }
      ]
    }
  ],
  "global_notes": ""
}

Rules:
- Do not infer missing values.
- Use "UNCLEAR" for hard-to-read handwriting.
- Preserve examiner-written responses as closely as possible.
- If a score is circled or visibly marked, capture it.
- If the protocol area is blank, use an empty string.
- Return valid JSON only.
"""
                    },
                    {
                        "type": "input_image",
                        "image_url": f"data:image/png;base64,{image_b64}",
                    },
                ],
            }
        ],
    )

    text = response.output_text
    return json.loads(text)


# =====================================================
# FLATTEN TO TABLE
# =====================================================

def flatten_protocol(data: dict) -> pd.DataFrame:
    rows = []
    info = data.get("examinee_info", {})

    for subtest in data.get("subtests", []):
        subtest_name = subtest.get("subtest_name", "")

        for item in subtest.get("items", []):
            row = {
                **info,
                "subtest_name": subtest_name,
                "item_number": item.get("item_number", ""),
                "printed_prompt": item.get("printed_prompt", ""),
                "handwritten_response": item.get("handwritten_response", ""),
                "dk_marked": item.get("dk_marked", False),
                "nr_marked": item.get("nr_marked", False),
                "score_marked": item.get("score_marked", ""),
                "confidence": item.get("confidence", ""),
                "notes": item.get("notes", ""),
            }
            rows.append(row)

    return pd.DataFrame(rows)


# =====================================================
# MAIN
# =====================================================

def main() -> None:
    print("Reading image:")
    print(IMAGE_FILE)

    data = extract_protocol_data(IMAGE_FILE)

    JSON_OUT.write_text(json.dumps(data, indent=2), encoding="utf-8")

    df = flatten_protocol(data)
    df.to_excel(EXCEL_OUT, index=False)
    df.to_csv(CSV_OUT, index=False)

    print("\nDone.")
    print(JSON_OUT)
    print(EXCEL_OUT)
    print(CSV_OUT)


if __name__ == "__main__":
    main()