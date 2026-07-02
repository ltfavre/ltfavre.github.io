import base64
import json
import mimetypes
import os
import sys
from pathlib import Path

import pandas as pd
from openai import OpenAI

sys.path.append(str(Path(__file__).resolve().parents[1]))

from config.AISI_protocol_config import (
    INPUT_FILE,
    OUTPUT_DIR,
    OUTPUT_JSON,
    OUTPUT_EXCEL,
    OUTPUT_CSV,
    ENV_FILE,
    MODEL_NAME,
    SUBTESTS_TO_EXTRACT,
)

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}
SUPPORTED_PDF_EXTENSIONS = {".pdf"}


def load_env_file():
    if not ENV_FILE.exists():
        raise FileNotFoundError(f"Missing .env file: {ENV_FILE}")

    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def encode_file(input_path: Path) -> str:
    return base64.b64encode(input_path.read_bytes()).decode("utf-8")


def build_input_content(input_path: Path) -> dict:
    if not input_path.exists():
        raise FileNotFoundError(input_path)

    suffix = input_path.suffix.lower()
    encoded = encode_file(input_path)

    if suffix in SUPPORTED_IMAGE_EXTENSIONS:
        mime_type = mimetypes.guess_type(input_path.name)[0] or "image/png"
        return {
            "type": "input_image",
            "image_url": f"data:{mime_type};base64,{encoded}",
        }

    if suffix in SUPPORTED_PDF_EXTENSIONS:
        return {
            "type": "input_file",
            "filename": input_path.name,
            "file_data": f"data:application/pdf;base64,{encoded}",
        }

    allowed = sorted(SUPPORTED_IMAGE_EXTENSIONS | SUPPORTED_PDF_EXTENSIONS)
    raise ValueError(f"Unsupported input type '{suffix}'. Use one of: {allowed}")


def parse_json_response(text: str) -> dict:
    text = text.strip()

    if text.startswith("```"):
        text = text.removeprefix("```json").removeprefix("```").strip()
        text = text.removesuffix("```").strip()

    return json.loads(text)


def capture_protocol_with_vision(input_path: Path) -> dict:
    file_content = build_input_content(input_path)

    load_env_file()
    client = OpenAI()

    prompt = f"""
You are extracting data from a fake administered WISC-style protocol.

Extract printed and handwritten information.

Subtests to extract:
{SUBTESTS_TO_EXTRACT}

Return valid JSON only with this structure:

{{
  "examinee_info": {{
    "name": "",
    "id_number": "",
    "date_of_birth": "",
    "age_at_testing": "",
    "date_of_testing": "",
    "examiner_name": "",
    "examiner_id": "",
    "site_id": ""
  }},
  "subtests": [
    {{
      "subtest_name": "",
      "start_time": "",
      "end_time": "",
      "total_raw_score": "",
      "items": [
        {{
          "item_number": "",
          "printed_prompt": "",
          "handwritten_response": "",
          "score_marked": "",
          "dk_marked": false,
          "nr_marked": false,
          "confidence": 0.0,
          "notes": ""
        }}
      ]
    }}
  ],
  "global_notes": "",
  "examiner_comments": ""
}}

Rules:
- Preserve handwriting as closely as possible.
- Keep line breaks in handwritten responses using "\\n".
- Use "UNCLEAR" if handwriting is difficult to read.
- Do not infer missing information.
- If a score is circled, written, or marked, capture it.
- If a field is blank, return an empty string.
"""

    response = client.responses.create(
        model=MODEL_NAME,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    file_content,
                ],
            }
        ],
    )

    return parse_json_response(response.output_text)


def flatten_protocol(data: dict) -> pd.DataFrame:
    rows = []
    info = data.get("examinee_info", {})

    for subtest in data.get("subtests", []):
        for item in subtest.get("items", []):
            rows.append(
                {
                    **info,
                    "subtest_name": subtest.get("subtest_name", ""),
                    "subtest_start_time": subtest.get("start_time", ""),
                    "subtest_end_time": subtest.get("end_time", ""),
                    "subtest_total_raw_score": subtest.get("total_raw_score", ""),
                    "item_number": item.get("item_number", ""),
                    "printed_prompt": item.get("printed_prompt", ""),
                    "handwritten_response": item.get("handwritten_response", ""),
                    "score_marked": item.get("score_marked", ""),
                    "dk_marked": item.get("dk_marked", False),
                    "nr_marked": item.get("nr_marked", False),
                    "confidence": item.get("confidence", ""),
                    "notes": item.get("notes", ""),
                    "global_notes": data.get("global_notes", ""),
                    "examiner_comments": data.get("examiner_comments", ""),
                }
            )

    return pd.DataFrame(rows)


def write_outputs(data: dict, df: pd.DataFrame):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    OUTPUT_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    df.to_excel(OUTPUT_EXCEL, index=False)
    df.to_csv(OUTPUT_CSV, index=False)


def main():
    print("AISI protocol capture + parser")
    print(f"Input file: {INPUT_FILE}")

    data = capture_protocol_with_vision(INPUT_FILE)
    df = flatten_protocol(data)
    write_outputs(data, df)

    print("\nDone.")
    print(OUTPUT_JSON)
    print(OUTPUT_EXCEL)
    print(OUTPUT_CSV)


if __name__ == "__main__":
    main()
