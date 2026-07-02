from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import json
import re
import traceback
import typing

import pandas as pd

SOURCE_TYPES = {"auto", "azure_document_intelligence", "gpt_protocol"}


# ============================================================
# USER CONFIG
# EDIT THIS SECTION ONLY 
# ============================================================
CONFIG_FILE = Path(
    r"C:\Users\UFAVRLA\PythonProjects\KTEA_CELF_LLM_Scoring_Lab\config\AISI_json_parser_config.xlsx"
)


# ============================================================
# CONFIG LOADING
# ============================================================
@dataclass
class ParserConfig:
    input_path: Path
    output_excel_path: Path
    fields_to_extract: list[str]
    always_raw_fields: set[str]
    final_column_order: list[str]
    source_type: str = "auto"
    add_timestamp: bool = True
    preserve_line_breaks: bool = True
    line_break_token: str = "\\n"
    include_confidence_columns: bool = True
    file_id_column: str = "FileID"
    section_labels: set[str] | None = None
    allow_s_section_prefixes: set[str] | None = None
    numeric_trial_section_prefixes: set[str] | None = None
    nested_sample_section_prefixes: set[str] | None = None
    lowercase_verbatim_prefixes: tuple[str, ...] = ("si_p", "vc_p", "co_p", "in_p12")


def _truthy(value: typing.Any, default: bool = False) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    return str(value).strip().upper() in {"TRUE", "T", "YES", "Y", "1"}


def _split_list(value: typing.Any) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    return [x.strip() for x in str(value).split(",") if x.strip()]


def _path_from_setting(value: typing.Any) -> Path:
    return Path(str(value).strip().strip('"').strip("'"))


def _source_type_from_setting(value: typing.Any) -> str:
    source_type = str(value or "auto").strip().lower()
    if source_type not in SOURCE_TYPES:
        allowed = ", ".join(sorted(SOURCE_TYPES))
        raise ValueError(f"Invalid source_type '{source_type}'. Use one of: {allowed}")
    return source_type


def load_config(config_file: Path = CONFIG_FILE) -> ParserConfig:
    if not config_file.exists():
        raise FileNotFoundError(f"Config workbook not found: {config_file}")

    settings_df = pd.read_excel(config_file, sheet_name="settings")
    settings = {
        str(row["setting"]).strip(): row["value"]
        for _, row in settings_df.iterrows()
        if pd.notna(row.get("setting"))
    }

    fields_df = pd.read_excel(config_file, sheet_name="fields_to_extract")
    fields_to_extract = [
        str(x).strip()
        for x in fields_df.loc[
            fields_df.get("extract", "TRUE").astype(str).str.upper().isin(["TRUE", "YES", "1"]),
            "field_name",
        ].dropna()
        if str(x).strip()
    ]

    raw_df = pd.read_excel(config_file, sheet_name="always_raw_fields")
    always_raw_fields = {
        str(x).strip()
        for x in raw_df.loc[
            raw_df.get("always_raw", "TRUE").astype(str).str.upper().isin(["TRUE", "YES", "1"]),
            "field_name",
        ].dropna()
        if str(x).strip()
    }

    order_df = pd.read_excel(config_file, sheet_name="final_column_order")
    final_column_order = [
        str(x).strip()
        for x in order_df.sort_values("order")["column_name"].dropna()
        if str(x).strip()
    ]
    final_column_order = list(dict.fromkeys(final_column_order))

    rules_df = pd.read_excel(config_file, sheet_name="parser_rules")
    rules = {
        str(row["rule"]).strip(): row["value"]
        for _, row in rules_df.iterrows()
        if pd.notna(row.get("rule"))
    }

    return ParserConfig(
        input_path=_path_from_setting(settings["input_path"]),
        output_excel_path=_path_from_setting(settings["output_excel_path"]),
        fields_to_extract=fields_to_extract,
        always_raw_fields=always_raw_fields,
        final_column_order=final_column_order,
        source_type=_source_type_from_setting(settings.get("source_type", "auto")),
        add_timestamp=_truthy(settings.get("add_timestamp"), True),
        preserve_line_breaks=_truthy(settings.get("preserve_line_breaks"), True),
        line_break_token=str(settings.get("line_break_token", "\\n")),
        include_confidence_columns=_truthy(settings.get("include_confidence_columns"), True),
        file_id_column=str(settings.get("file_id_column", "FileID")),
        section_labels=set(_split_list(rules.get("section_labels"))) or {"DA", "DB", "DC", "SA", "SB", "Sample", "d", "ra", "rb", "sa", "sb"},
        allow_s_section_prefixes=set(_split_list(rules.get("allow_s_section_prefixes"))) or {"vp_l", "ar_l"},
        numeric_trial_section_prefixes=set(_split_list(rules.get("numeric_trial_section_prefixes"))) or {"df", "ds", "lns", "db", "db_p14", "db_p15"},
        nested_sample_section_prefixes=set(_split_list(rules.get("nested_sample_section_prefixes"))) or {"ds", "lns", "db", "db_p14", "db_p15"},
        lowercase_verbatim_prefixes=tuple(_split_list(rules.get("lowercase_verbatim_prefixes"))) or ("si_p", "vc_p", "co_p", "in_p12"),
    )


# ============================================================
# FIELD EXTRACTION
# ============================================================
def best_value(field_obj: typing.Any) -> typing.Any:
    """Extract a usable value from a Document Intelligence field object."""
    if not isinstance(field_obj, dict):
        return field_obj

    val = field_obj.get("content") or field_obj.get("valueString")

    if val is None:
        for key in (
            "value", "valueNumber", "valueInteger", "valueDate", "valueTime",
            "valuePhoneNumber", "valueBoolean", "valueSelectionGroup",
        ):
            if field_obj.get(key) is not None:
                val = field_obj[key]
                break

    if val is None and isinstance(field_obj.get("valueArray"), list):
        parts = [str(best_value(v)) for v in field_obj["valueArray"]]
        val = " | ".join(parts) if parts else None

    if val is None and isinstance(field_obj.get("valueObject"), dict):
        parts = [f"{k}: {best_value(v)}" for k, v in field_obj["valueObject"].items()]
        val = " | ".join(parts) if parts else None

    return val


def extract_from_document_fields(data: dict, field_names: list[str]) -> dict:
    extracted: dict[str, typing.Any] = {}
    docs = data.get("analyzeResult", {}).get("documents", [])

    for doc in docs:
        fields = doc.get("fields", {})
        if not isinstance(fields, dict):
            continue

        for name in field_names:
            if name not in fields or name in extracted:
                continue

            field_obj = fields[name]
            value = best_value(field_obj)
            conf = field_obj.get("confidence") if isinstance(field_obj, dict) else None
            conf_pct = round(conf * 100, 2) if isinstance(conf, (int, float)) else None
            extracted[name] = value
            extracted[f"{name}_confidence"] = conf_pct

    return extracted


# ============================================================
# HELPERS
# ============================================================
ROW_HEADER_RE = re.compile(r"(?m)(^|\n|\|)\s*(\d+|Sample|SA|SB|DA|DB|DC)\s*:", re.IGNORECASE)
KV_START_RE = re.compile(r"^\s*([A-Za-z0-9_\-]+)\s*:\s*(.*)$")
INLINE_SECTION_RE = re.compile(r"(^|\|)\s*([A-Za-z0-9_]{1,20})\s*:\s*[A-Za-z0-9_]+\s*:")
SECTION_KEY_RE = re.compile(r"^[A-Z]{2}$")

LEGIT_ROW_KEYS = {
    "time", "Response", "response", "Verbatim", "verbatim",
    "r1", "r2", "r3", "r4", "r5", "r6",
    "dk", "nr", "ir", "de", "re",
    "s", "s0", "s1", "s2", "s4", "s5", "s6", "s7",
    "t0", "t1", "t1_2", "time_1", "time_2",
}


def normalize_text_breaks(value: typing.Any) -> typing.Any:
    """
    Normalize all line-break representations before parsing.

    This handles:
    - real CRLF/CR/LF breaks
    - literal text "\\n" emitted by models
    - Azure pipe separators
    - unicode line/paragraph separators
    """
    if not isinstance(value, str):
        return value

    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\\n", "\n")
    text = text.replace("\u2028", "\n").replace("\u2029", "\n")
    text = re.sub(r"\s*\|\s*", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def format_embedded_line_breaks(value: typing.Any, cfg: ParserConfig) -> typing.Any:
    """Store internal line breaks consistently in Excel/CSV output."""
    if not isinstance(value, str):
        return value
    text = normalize_text_breaks(value)
    if cfg.preserve_line_breaks:
        return text
    return text.replace("\n", " ")


def looks_like_table_string(value: typing.Any) -> bool:
    text = normalize_text_breaks(value)
    return isinstance(text, str) and bool(ROW_HEADER_RE.search(text))


def looks_like_inline_section(value: typing.Any) -> bool:
    text = normalize_text_breaks(value)
    return isinstance(text, str) and bool(INLINE_SECTION_RE.search(text))


def normalise_key(prefix: str, key: str, cfg: ParserConfig) -> str:
    if not isinstance(key, str):
        return key

    key = key.strip()
    if prefix.startswith(cfg.lowercase_verbatim_prefixes) and key.lower() == "verbatim":
        return "verbatim"
    return key


def is_numeric_trial_label(label: str) -> bool:
    return bool(re.match(r"^\d+_\d+$", label))


def is_sample_trial_label(label: str) -> bool:
    return bool(re.match(r"^(sa|sb|s)_\d+$", label, flags=re.IGNORECASE))


def is_section_label(prefix: str, possible_section: str, cfg: ParserConfig) -> bool:
    if possible_section in cfg.section_labels:
        return True
    if possible_section == "s" and prefix in cfg.allow_s_section_prefixes:
        return True
    if prefix in cfg.numeric_trial_section_prefixes and is_numeric_trial_label(possible_section):
        return True
    if prefix in cfg.nested_sample_section_prefixes and is_sample_trial_label(possible_section):
        return True
    return False


def split_key_value_lines(chunk: str) -> list[tuple[str, str]]:
    """
    Split a chunk into key/value pairs while preserving continuation lines.

    Example:
        Verbatim: first line
        second handwritten line
        dk: None

    becomes:
        [("Verbatim", "first line\nsecond handwritten line"), ("dk", "None")]
    """
    pairs: list[tuple[str, str]] = []
    current_key: str | None = None
    current_lines: list[str] = []

    for raw_line in normalize_text_breaks(chunk).split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        match = KV_START_RE.match(line)
        if match:
            # If this looks like a section marker, let caller handle section boundaries.
            key, val = match.groups()
            if current_key is not None:
                pairs.append((current_key, "\n".join(current_lines).strip()))
            current_key = key.strip()
            current_lines = [val.strip()]
        else:
            if current_key is not None:
                current_lines.append(line)
            else:
                # Leading unkeyed text is preserved under a generic value key.
                current_key = "value"
                current_lines = [line]

    if current_key is not None:
        pairs.append((current_key, "\n".join(current_lines).strip()))

    return pairs


def parse_table_string(prefix: str, value: typing.Any, cfg: ParserConfig) -> dict:
    out: dict[str, typing.Any] = {}
    text = normalize_text_breaks(value)
    if not isinstance(text, str) or not text:
        return out

    rows = list(ROW_HEADER_RE.finditer(text))
    for idx, marker in enumerate(rows):
        row = marker.group(2).strip()
        start = marker.end()
        end = rows[idx + 1].start() if idx + 1 < len(rows) else len(text)
        chunk = text[start:end].strip()

        current_sub = None
        for key, val in split_key_value_lines(chunk):
            key = normalise_key(prefix, key, cfg)

            if SECTION_KEY_RE.match(str(key)) and key not in LEGIT_ROW_KEYS:
                # Treat section markers as boundaries, not values.
                break

            if prefix == "sr_p19" and row == "12" and key in ("DA", "DB", "DC"):
                current_sub = key
                continue

            if prefix == "sr_p19" and row == "12" and current_sub:
                out[f"{prefix}_{row}_{current_sub}_{key}"] = format_embedded_line_breaks(val, cfg)
                continue

            out[f"{prefix}_{row}_{key}"] = format_embedded_line_breaks(val, cfg)

    return out


def parse_inline_sections(prefix: str, value: typing.Any, cfg: ParserConfig) -> dict:
    out: dict[str, typing.Any] = {}
    text = normalize_text_breaks(value)
    if not isinstance(text, str) or not text:
        return out

    # Important: after normalize_text_breaks, pipes and literal \n have become real newlines.
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    current_section = None
    current_key = None
    current_value_lines: list[str] = []

    def flush_current():
        nonlocal current_key, current_value_lines, current_section
        if current_section is not None and current_key is not None:
            out[f"{prefix}_{current_section}_{current_key}"] = format_embedded_line_breaks(
                "\n".join(current_value_lines).strip(), cfg
            )
        current_key = None
        current_value_lines = []

    for part in parts:
        working = part

        while ":" in working:
            possible_section, remainder = [x.strip() for x in working.split(":", 1)]

            if is_section_label(prefix, possible_section, cfg):
                flush_current()
                current_section = possible_section
                working = remainder
                if not working:
                    break
                continue

            if (
                prefix in cfg.nested_sample_section_prefixes
                and current_section in {"sa", "sb"}
                and re.match(r"^\d+$", possible_section)
            ):
                flush_current()
                current_section = f"{current_section}_{possible_section}"
                working = remainder
                if not working:
                    break
                continue

            break

        if not working:
            continue

        kv = KV_START_RE.match(working)
        if kv and current_section is not None:
            flush_current()
            key, val = kv.groups()
            current_key = normalise_key(prefix, key, cfg)
            current_value_lines = [":selected:" if ":selected:" in val else val.strip()]
        elif current_section is not None and current_key is not None:
            # Continuation line for previous key.
            current_value_lines.append(working)

    flush_current()
    return out


# ============================================================
# OUTPUT
# ============================================================
def make_file_id(stem: str) -> str:
    for suffix in ("_result", "_output", "_analysis", "_analyze"):
        if stem.endswith(suffix):
            return stem[:-len(suffix)]
    return stem


def apply_final_column_order(df: pd.DataFrame, cfg: ParserConfig) -> pd.DataFrame:
    desired = cfg.final_column_order
    missing = [c for c in desired if c not in df.columns]
    if missing:
        print(f"Info: {len(missing)} configured columns had no captured values and will be blank.")
        print("   First 30 missing:", missing[:30])

    remainder = [c for c in df.columns if c not in desired]
    return df.reindex(columns=desired + remainder)


def resolve_input_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".json":
            raise ValueError(f"Input file must be a .json file: {input_path}")
        return [input_path]

    if input_path.is_dir():
        return sorted(input_path.glob("*.json"))

    raise FileNotFoundError(f"Input path not found: {input_path}")


def output_path_with_timestamp(path: Path, add_timestamp: bool) -> Path:
    if not add_timestamp:
        return path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{ts}{path.suffix}")


# ============================================================
# PARSING
# ============================================================
GPT_PROTOCOL_COLUMNS = [
    "name",
    "id_number",
    "date_of_birth",
    "age_at_testing",
    "date_of_testing",
    "examiner_name",
    "examiner_id",
    "site_id",
    "subtest_name",
    "subtest_start_time",
    "subtest_end_time",
    "subtest_total_raw_score",
    "item_number",
    "printed_prompt",
    "handwritten_response",
    "score_marked",
    "dk_marked",
    "nr_marked",
    "confidence",
    "notes",
    "global_notes",
    "examiner_comments",
]


def is_gpt_protocol_json(data: dict) -> bool:
    return isinstance(data.get("subtests"), list)


def is_azure_document_intelligence_json(data: dict) -> bool:
    docs = data.get("analyzeResult", {}).get("documents")
    return isinstance(docs, list)


def detect_source_type(data: dict) -> str:
    if is_gpt_protocol_json(data):
        return "gpt_protocol"
    if is_azure_document_intelligence_json(data):
        return "azure_document_intelligence"
    return "unsupported"


def flatten_gpt_protocol_rows(data: dict, cfg: ParserConfig) -> list[dict[str, typing.Any]]:
    rows = []
    info = data.get("examinee_info", {})
    if not isinstance(info, dict):
        info = {}

    for subtest in data.get("subtests", []):
        if not isinstance(subtest, dict):
            continue

        for item in subtest.get("items", []):
            if not isinstance(item, dict):
                continue

            rows.append(
                {
                    **info,
                    "subtest_name": subtest.get("subtest_name", ""),
                    "subtest_start_time": subtest.get("start_time", ""),
                    "subtest_end_time": subtest.get("end_time", ""),
                    "subtest_total_raw_score": subtest.get("total_raw_score", ""),
                    "item_number": item.get("item_number", ""),
                    "printed_prompt": item.get("printed_prompt", ""),
                    "handwritten_response": format_embedded_line_breaks(
                        item.get("handwritten_response", ""), cfg
                    ),
                    "score_marked": item.get("score_marked", ""),
                    "dk_marked": item.get("dk_marked", False),
                    "nr_marked": item.get("nr_marked", False),
                    "confidence": item.get("confidence", ""),
                    "notes": item.get("notes", ""),
                    "global_notes": data.get("global_notes", ""),
                    "examiner_comments": data.get("examiner_comments", ""),
                }
            )

    return rows


def gpt_item_column_base(subtest_name: str, item_number: typing.Any) -> str | None:
    item = str(item_number).strip()
    name = str(subtest_name).strip().lower()

    if name == "similarities":
        if item.upper() in {"S", "SAMPLE"}:
            return "si_p1_Sample"
        if item.isdigit():
            item_int = int(item)
            page = "p1" if item_int <= 11 else "p2"
            return f"si_{page}_{item_int}"

    if name == "vocabulary" and item.isdigit():
        item_int = int(item)
        if 1 <= item_int <= 3:
            return f"vc_p7_1-3_{item_int}"
        page = "p7" if item_int <= 12 else "p8"
        return f"vc_{page}_{item_int}"

    if name == "comprehension" and item.isdigit():
        item_int = int(item)
        page = "p17" if item_int <= 9 else "p18"
        return f"co_{page}_{item_int}"

    return None


def parse_gpt_protocol_json(json_file: Path, data: dict, cfg: ParserConfig) -> dict:
    row: dict[str, typing.Any] = {cfg.file_id_column: make_file_id(json_file.stem)}

    info = data.get("examinee_info", {})
    if isinstance(info, dict):
        row["id"] = info.get("id_number") or info.get("id") or ""

    for subtest in data.get("subtests", []):
        if not isinstance(subtest, dict):
            continue

        name = str(subtest.get("subtest_name", "")).strip().lower()
        prefix = {
            "similarities": "si",
            "vocabulary": "vc",
            "comprehension": "co",
        }.get(name)

        if prefix:
            row[f"{prefix}_st_raw"] = subtest.get("start_time", "")
            row[f"{prefix}_et_raw"] = subtest.get("end_time", "")

        for item in subtest.get("items", []):
            if not isinstance(item, dict):
                continue

            base = gpt_item_column_base(subtest.get("subtest_name", ""), item.get("item_number", ""))
            if not base:
                continue

            row[f"{base}_verbatim"] = format_embedded_line_breaks(item.get("handwritten_response", ""), cfg)
            row[f"{base}_dk"] = bool(item.get("dk_marked", False))
            row[f"{base}_nr"] = bool(item.get("nr_marked", False))

            score = str(item.get("score_marked", "")).strip()
            if score.isdigit():
                row[f"{base}_s{score}"] = True

            confidence = item.get("confidence")
            if cfg.include_confidence_columns and isinstance(confidence, (int, float)):
                row[f"{base}_confidence"] = round(confidence * 100, 2)

    row["global_notes"] = data.get("global_notes", "")
    row["examiner_comments"] = data.get("examiner_comments", "")
    return row


def parse_one_json_file(json_file: Path, cfg: ParserConfig) -> dict:
    with open(json_file, encoding="utf-8") as handle:
        data = json.load(handle)

    source_type = detect_source_type(data)
    if source_type == "gpt_protocol":
        return parse_gpt_protocol_json(json_file, data, cfg)
    if source_type != "azure_document_intelligence":
        raise ValueError(
            f"Unsupported JSON source format in {json_file}. Expected either "
            "Azure Document Intelligence JSON with analyzeResult.documents or "
            "GPT protocol JSON with subtests."
        )

    extracted = extract_from_document_fields(data, cfg.fields_to_extract)
    row: dict[str, typing.Any] = {cfg.file_id_column: make_file_id(json_file.stem)}

    for field_name in cfg.fields_to_extract:
        value = extracted.get(field_name)
        value = normalize_text_breaks(value)

        if field_name in cfg.always_raw_fields:
            row[f"{field_name}_raw"] = format_embedded_line_breaks(value, cfg)
        elif looks_like_table_string(value):
            row[f"{field_name}_raw"] = format_embedded_line_breaks(value, cfg)
            row.update(parse_table_string(field_name, value, cfg))
            row.update(parse_inline_sections(field_name, value, cfg))
        elif looks_like_inline_section(value):
            row[f"{field_name}_raw"] = format_embedded_line_breaks(value, cfg)
            row.update(parse_inline_sections(field_name, value, cfg))
        else:
            row[field_name] = format_embedded_line_breaks(value, cfg)

        if cfg.include_confidence_columns:
            row[f"{field_name}_confidence"] = extracted.get(f"{field_name}_confidence")

    return row


def main() -> None:
    cfg = load_config(CONFIG_FILE)
    files = resolve_input_files(cfg.input_path)
    if not files:
        raise FileNotFoundError(f"No JSON files found in {cfg.input_path}")

    loaded_files = []
    for json_file in files:
        print(f"Parsing: {json_file.name}")
        with open(json_file, encoding="utf-8") as handle:
            loaded_files.append((json_file, json.load(handle)))

    source_types = {detect_source_type(data) for _, data in loaded_files}
    if "unsupported" in source_types:
        unsupported = [str(path) for path, data in loaded_files if detect_source_type(data) == "unsupported"]
        raise ValueError(
            "Unsupported JSON source format. Expected either Azure Document Intelligence "
            "JSON with analyzeResult.documents or GPT protocol JSON with subtests. "
            f"Unsupported files: {unsupported}"
        )
    if len(source_types) > 1:
        raise ValueError(
            "Mixed JSON source formats found in one run. Put Azure Document Intelligence "
            "files and GPT protocol files in separate folders/runs."
        )

    output_file = output_path_with_timestamp(cfg.output_excel_path, cfg.add_timestamp)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    detected_source_type = next(iter(source_types))
    if cfg.source_type == "auto":
        source_type = detected_source_type
    else:
        source_type = cfg.source_type
        if detected_source_type != source_type:
            raise ValueError(
                f"Configured source_type is '{source_type}', but the input files look like "
                f"'{detected_source_type}'. Set source_type to 'auto' or update the input files."
            )

    if source_type == "gpt_protocol":
        rows = []
        for _, data in loaded_files:
            rows.extend(flatten_gpt_protocol_rows(data, cfg))

        df = pd.DataFrame(rows).reindex(columns=GPT_PROTOCOL_COLUMNS)
        df.to_excel(output_file, index=False)
        print(f"Saved {len(df)} rows and {len(df.columns)} columns to {output_file}")
        return

    rows = [parse_one_json_file(json_file, cfg) for json_file, _ in loaded_files]

    df = pd.DataFrame(rows)
    df = apply_final_column_order(df, cfg)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="parsed_data", index=False)

        summary = pd.DataFrame(
            {
                "metric": ["files_parsed", "columns_output", "rows_output"],
                "value": [len(files), len(df.columns), len(df)],
            }
        )
        summary.to_excel(writer, sheet_name="run_summary", index=False)

    print(f"Saved {len(df)} rows and {len(df.columns)} columns to {output_file}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
        raise
