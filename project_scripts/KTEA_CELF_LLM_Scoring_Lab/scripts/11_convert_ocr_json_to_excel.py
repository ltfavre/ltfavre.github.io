import argparse
import json
from datetime import datetime
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_JSON = PROJECT_ROOT / "data" / "raw" / "30120 RF.pdf"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "ocr"


def score_to_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Input JSON not found: {path}")

    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def flatten_items(data: dict) -> pd.DataFrame:
    examinee = data.get("examinee_info", {})
    rows = []

    for subtest_index, subtest in enumerate(data.get("subtests", []), start=1):
        subtest_name = (
            subtest.get("subtest")
            or subtest.get("subtest_name")
            or f"Subtest {subtest_index}"
        )

        for item in subtest.get("items", []):
            score = item.get("captured_score", item.get("score_marked", ""))
            row = {
                "source_name": examinee.get("name", ""),
                "dob": examinee.get("dob", examinee.get("birth_date", "")),
                "test_date": examinee.get("test_date", ""),
                "age": examinee.get("age", examinee.get("test_age", "")),
                "examiner": examinee.get("examiner", examinee.get("examiner_name", "")),
                "school": examinee.get("school", ""),
                "grade": examinee.get("grade", ""),
                "subtest": subtest_name,
                "item_number": item.get("item_number", ""),
                "item_prompt": item.get("item_prompt", item.get("printed_prompt", "")),
                "handwritten_response": item.get("handwritten_response", ""),
                "captured_score": score,
                "score_numeric": score_to_number(score),
                "confidence": item.get("confidence", ""),
                "notes": item.get("notes", ""),
            }
            rows.append(row)

    return pd.DataFrame(rows)


def summarize_subtests(items_df: pd.DataFrame) -> pd.DataFrame:
    if items_df.empty:
        return pd.DataFrame(
            columns=[
                "subtest",
                "item_count",
                "scored_item_count",
                "total_score",
                "mean_confidence",
                "unclear_item_count",
            ]
        )

    summary = (
        items_df.groupby("subtest", dropna=False)
        .agg(
            item_count=("item_number", "count"),
            scored_item_count=("score_numeric", "count"),
            total_score=("score_numeric", "sum"),
            mean_confidence=("confidence", "mean"),
        )
        .reset_index()
    )

    unclear = (
        items_df.assign(
            is_unclear=items_df["captured_score"]
            .astype(str)
            .str.upper()
            .eq("UNCLEAR")
        )
        .groupby("subtest", dropna=False)["is_unclear"]
        .sum()
        .reset_index(name="unclear_item_count")
    )

    return summary.merge(unclear, on="subtest", how="left")


def examinee_info_table(data: dict) -> pd.DataFrame:
    info = data.get("examinee_info", {})
    return pd.DataFrame(
        [{"field": key, "value": value} for key, value in info.items()]
    )


def notes_table(data: dict) -> pd.DataFrame:
    notes = data.get("notes", data.get("global_notes", ""))
    return pd.DataFrame([{"notes": notes}])


def autofit_columns(writer, sheet_name: str, df: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]

    for col_idx, column_name in enumerate(df.columns, start=1):
        values = [str(column_name)]
        values.extend(str(v) for v in df[column_name].fillna("").head(200))
        width = min(max(len(v) for v in values) + 2, 60)
        worksheet.column_dimensions[worksheet.cell(1, col_idx).column_letter].width = width

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def write_excel(data: dict, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    items_df = flatten_items(data)
    summary_df = summarize_subtests(items_df)
    info_df = examinee_info_table(data)
    notes_df = notes_table(data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        items_df.to_excel(writer, sheet_name="Items", index=False)
        summary_df.to_excel(writer, sheet_name="Subtest Summary", index=False)
        info_df.to_excel(writer, sheet_name="Examinee Info", index=False)
        notes_df.to_excel(writer, sheet_name="Notes", index=False)

        for sheet_name, df in {
            "Items": items_df,
            "Subtest Summary": summary_df,
            "Examinee Info": info_df,
            "Notes": notes_df,
        }.items():
            autofit_columns(writer, sheet_name, df)

    return output_path


def build_output_path(input_json: Path, output_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_dir / f"{input_json.stem}_usable_{timestamp}.xlsx"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Convert OCR protocol JSON into a usable Excel workbook."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT_JSON,
        help=f"Input JSON path. Default: {DEFAULT_INPUT_JSON}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx path. Default: timestamped workbook in outputs/ocr.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_json = args.input.resolve()
    output_path = args.output.resolve() if args.output else build_output_path(
        input_json, DEFAULT_OUTPUT_DIR
    )

    data = load_json(input_json)
    saved_path = write_excel(data, output_path)
    print(f"Saved Excel workbook: {saved_path}")


if __name__ == "__main__":
    main()
