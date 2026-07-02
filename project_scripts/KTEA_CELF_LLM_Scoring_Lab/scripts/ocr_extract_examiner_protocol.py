import base64
import json
from pathlib import Path
import mimetypes

import pandas as pd

from dotenv import load_dotenv
from openai import OpenAI

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")
client = OpenAI()

INPUT_FILE = Path(r"C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab\data\raw\30120 RF.pdf")
OUTPUT_DIR = Path(r"C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab\outputs\ocr")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}
SUPPORTED_PDF_EXTENSIONS = {".pdf"}


def encode_file(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def build_file_content(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(path)

    suffix = path.suffix.lower()
    encoded = encode_file(path)

    if suffix in SUPPORTED_IMAGE_EXTENSIONS:
        mime_type = mimetypes.guess_type(path.name)[0] or "image/png"
        return {
            "type": "input_image",
            "image_url": f"data:{mime_type};base64,{encoded}",
        }

    if suffix in SUPPORTED_PDF_EXTENSIONS:
        return {
            "type": "input_file",
            "filename": path.name,
            "file_data": f"data:application/pdf;base64,{encoded}",
        }

    allowed = sorted(SUPPORTED_IMAGE_EXTENSIONS | SUPPORTED_PDF_EXTENSIONS)
    raise ValueError(f"Unsupported input type '{suffix}'. Use one of: {allowed}")


schema = {
    "type": "object",
    "properties": {
        "examinee_info": {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "dob": {"type": "string"},
                "test_date": {"type": "string"},
                "age": {"type": "string"},
                "examiner": {"type": "string"},
                "school": {"type": "string"},
                "grade": {"type": "string"}
            },
            "required": ["name", "dob", "test_date", "age", "examiner", "school", "grade"],
            "additionalProperties": False
        },
        "subtests": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "subtest": {"type": "string"},
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "item_number": {"type": "string"},
                                "item_prompt": {"type": "string"},
                                "handwritten_response": {"type": "string"},
                                "captured_score": {"type": "string"},
                                "confidence": {"type": "number"}
                            },
                            "required": [
                                "item_number",
                                "item_prompt",
                                "handwritten_response",
                                "captured_score",
                                "confidence"
                            ],
                            "additionalProperties": False
                        }
                    }
                },
                "required": ["subtest", "items"],
                "additionalProperties": False
            }
        },
        "notes": {"type": "string"}
    },
    "required": ["examinee_info", "subtests", "notes"],
    "additionalProperties": False
}


file_content = build_file_content(INPUT_FILE)

response = client.responses.create(
    model="gpt-4.1",
    input=[
        {
            "role": "user",
            "content": [
                {
                    "type": "input_text",
                    "text": """
You are extracting data from a fake WISC-V-style administered protocol for an OCR/data-capture test.

Extract:
1. Examinee information from the header.
2. Each visible subtest.
3. Each item number.
4. Printed item prompt.
5. Handwritten examiner response.
6. Circled or written score.
7. Confidence from 0 to 1.

Do not infer scores if unreadable. Use "UNCLEAR" when unsure.
Return only structured JSON matching the schema.
"""
                },
                file_content
            ]
        }
    ],
    text={
        "format": {
            "type": "json_schema",
            "name": "wisc_protocol_extraction",
            "schema": schema,
            "strict": True
        }
    }
)

data = json.loads(response.output_text)

output_stem = INPUT_FILE.stem

json_file = OUTPUT_DIR / f"{output_stem}.json"
json_file.write_text(json.dumps(data, indent=2), encoding="utf-8")

rows = []
for subtest in data["subtests"]:
    for item in subtest["items"]:
        rows.append({
            "subtest": subtest["subtest"],
            **item
        })

df = pd.DataFrame(rows)

excel_file = OUTPUT_DIR / f"{output_stem}.xlsx"
csv_file = OUTPUT_DIR / f"{output_stem}.csv"

df.to_excel(excel_file, index=False)
df.to_csv(csv_file, index=False)

print("Done.")
print(json_file)
print(excel_file)
print(csv_file)
