"""Refresh the reviewed CELF-6 stratification workbook after targeted swaps.

The workbook supplied by the CELF-6 team is the reporting authority.  This
module deliberately updates values in place instead of rebuilding worksheets,
so sheet names, labels, merges, dimensions, conditional formatting, and print
settings remain unchanged.
"""
from __future__ import annotations

from pathlib import Path
import re
import shutil

import pandas as pd
from openpyxl import load_workbook


REPORT_SHEETS = (
    "byage", "form_1", "form_2", "edl X sex", "edl X reg",
    "eth X sex", "eth X edl", "eth X reg", "reg X sex",
    "osep", "clinical_sources", "TEID_List",
)

AGE_BANDS = (
    ("5:0-5:5", lambda frame: frame["agey"].eq(5) & frame["agem"].between(0, 5)),
    ("5:6-5:11", lambda frame: frame["agey"].eq(5) & frame["agem"].between(6, 11)),
    ("6:0-6:5", lambda frame: frame["agey"].eq(6) & frame["agem"].between(0, 5)),
    ("6:6-6:11", lambda frame: frame["agey"].eq(6) & frame["agem"].between(6, 11)),
    *((f"{age}:0-{age}:11", lambda frame, age=age: frame["agey"].eq(age)) for age in range(7, 17)),
    ("17:0-21:11", lambda frame: frame["agey"].between(17, 21)),
)

CROSS_TABS = {
    "edl X sex": ("ped", "sex", True),
    "edl X reg": ("ped", "region", False),
    "eth X sex": ("eth", "sex", True),
    "eth X edl": ("eth", "ped", False),
    "eth X reg": ("eth", "region", False),
    "reg X sex": ("region", "sex", True),
}

HEADER_LEVELS = {
    "ped": {"ped_1": 1, "ped_2": 2, "ped_3": 3, "ped_4": 4,
            "edu_1": 1, "edu_2": 2, "edu_3": 3, "edu_4": 4},
    "eth": {"AS": "AS", "BL": "BL", "HI": "HI", "O": "OT", "OT": "OT", "W": "WH", "WH": "WH"},
    "region": {"MW": "MW", "NE": "NE", "SO": "SO", "W": "WE", "WE": "WE"},
    "sex": {"Female": "F", "Male": "M", "Other": "OTHER"},
}

GROUP_COLUMNS = {
    "LD": ("C_LD", ("C_LD_RF1", "C_LD_RF2")),
    "HH": ("C_HH", ("C_HH_RF1", "C_HH_RF2")),
    "ASD": ("C_ASDL", ("C_ASDL_RF1", "C_ASDL_RF2")),
    "SLD-RW": ("C_SLDRW", ("C_SLDRW_RF1", "C_SLDRW_RF2")),
    "DD/ID": ("OSEP_DDID", ("OSEP_DD_RF1", "OSEP_DD_RF2", "OSEP_ID_RF1", "OSEP_ID_RF2")),
    "SI": ("OSEP_SI", ("OSEP_SI_RF1", "OSEP_SI_RF2")),
    "ADHD": ("OSEP_ADHD", ("OSEP_ADHD_RF1", "OSEP_ADHD_RF2")),
    "SLD": ("OSEP_SLD", ("OSEP_SLD_RF1", "OSEP_SLD_RF2")),
    "GT": ("OSEP_GT", ("OSEP_GT_RF1", "OSEP_GT_RF2")),
}

DIAGNOSIS_PATTERNS = {
    "LD": r"\b(?:language disorder|language impairment|language delay|expressive language|receptive.language)\b",
    "DD/ID": r"\b(?:intellectual disability|intellectual developmental disorder)\b",
    "ADHD": r"\b(?:attention[- ]deficit(?:/hyperactivity)?(?: disorder)?|adhd)\b",
    "SLD": r"\b(?:specific learning (?:disability|disorder)|dyslexia|developmental learning disorder)\b",
    "GT": r"\b(?:gifted|talented)\b",
}


def is_celf6_stratification_template(path: str | Path) -> bool:
    """Return whether *path* has the reviewed CELF-6 workbook topology."""
    candidate = Path(path)
    if not candidate.is_file():
        return False
    try:
        workbook = load_workbook(candidate, read_only=True, data_only=False)
        return all(name in workbook.sheetnames for name in REPORT_SHEETS)
    except Exception:
        return False


def _clean_teid(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\.0$", "", str(value).strip())


def _label(value: object) -> str:
    return (
        str(value or "").strip().replace("–", "-").replace("—", "-")
        .replace("�", "-").replace(".0-", ":0-").replace(":00", ":0")
    )


def _prepare(frame: pd.DataFrame) -> pd.DataFrame:
    data = frame.copy()
    if "teid" not in data:
        raise ValueError("CELF-6 stratification requires a teid column.")
    data["teid"] = data["teid"].map(_clean_teid)
    age_source = "agey" if "agey" in data else "age" if "age" in data else None
    if age_source is None:
        raise ValueError("CELF-6 stratification requires agey (or canonical age).")
    data["agey"] = pd.to_numeric(data[age_source], errors="coerce")
    data["agem"] = pd.to_numeric(data.get("agem", 0), errors="coerce").fillna(0)
    if data["agey"].isna().any():
        raise ValueError("CELF-6 stratification found missing/non-numeric ages.")
    data["agey"] = data["agey"].astype(int)
    data["agem"] = data["agem"].astype(int)
    if "form" not in data:
        data["form"] = data["agey"].map(lambda value: "form_1" if value <= 8 else "form_2")
    else:
        normalized = data["form"].astype(str).str.strip().str.casefold()
        data["form"] = normalized.map(
            lambda value: "form_1" if value in {"1", "1.0", "form 1", "form_1"}
            else "form_2" if value in {"2", "2.0", "form 2", "form_2"} else value
        )
    if "ped" in data:
        data["ped"] = pd.to_numeric(data["ped"], errors="coerce")
    age_group = pd.Series("", index=data.index, dtype="object")
    for name, selector in AGE_BANDS:
        age_group.loc[selector(data)] = name
    data["_scope_age_group"] = age_group
    return data


def _band_mask(frame: pd.DataFrame, value: object) -> pd.Series:
    wanted = _label(value)
    return frame["_scope_age_group"].eq(wanted)


def _find_header_row(sheet, label: str) -> int:
    wanted = label.casefold()
    for row in range(1, sheet.max_row + 1):
        if str(sheet.cell(row, 1).value or "").strip().casefold() == wanted:
            return row
    raise ValueError(f"Could not find {label!r} on {sheet.title}.")


def _header_pairs(sheet, major: str, minor: str, major_row: int, minor_row: int):
    current = ""
    for column in range(3, sheet.max_column + 1):
        raw_major = sheet.cell(major_row, column).value
        if raw_major is not None and str(raw_major).strip():
            current = str(raw_major).strip()
        raw_minor = sheet.cell(minor_row, column).value
        if not current or raw_minor is None or not str(raw_minor).strip():
            continue
        major_value = HEADER_LEVELS[major].get(current)
        minor_value = HEADER_LEVELS[minor].get(str(raw_minor).strip())
        if major_value is not None and minor_value is not None:
            yield column, major_value, minor_value


def _write_byage(sheet, sample: pd.DataFrame) -> None:
    for row, (_, selector) in zip(range(3, 18), AGE_BANDS):
        group = sample.loc[selector(sample)]
        sheet.cell(row, 2).value = int(group["form"].eq("form_1").sum()) or None
        sheet.cell(row, 3).value = int(group["form"].eq("form_2").sum()) or None
    form_1_n = int(sample["form"].eq("form_1").sum())
    form_2_n = int(sample["form"].eq("form_2").sum())
    sheet["B18"], sheet["C18"] = form_1_n, form_2_n
    sheet["B19"] = form_1_n + form_2_n


def _write_form_sheet(sheet, sample: pd.DataFrame, form: str) -> None:
    headers = {
        str(sheet.cell(6, column).value or "").strip(): column
        for column in range(1, sheet.max_column + 1)
    }
    source_column = headers["Source"]
    age_column = headers["Age Group"]
    category_columns = []
    for header, column in headers.items():
        if header.startswith("ped_"):
            category_columns.append((column, "ped", HEADER_LEVELS["ped"].get(header)))
        elif header in HEADER_LEVELS["eth"]:
            category_columns.append((column, "eth", HEADER_LEVELS["eth"][header]))
        elif header in HEADER_LEVELS["region"]:
            category_columns.append((column, "region", HEADER_LEVELS["region"][header]))
    sex_columns = [(column, HEADER_LEVELS["sex"][header]) for header, column in headers.items() if header in HEADER_LEVELS["sex"]]
    for row in range(7, sheet.max_row + 1):
        if str(sheet.cell(row, source_column).value or "").strip().casefold() != "sample":
            continue
        group = sample.loc[sample["form"].eq(form) & _band_mask(sample, sheet.cell(row, age_column).value)]
        denominator = len(group)
        for column, field, value in category_columns:
            count = int(group[field].eq(value).sum()) if denominator and field in group else 0
            sheet.cell(row, column).value = count / denominator if denominator else 0
            sheet.cell(row, column).number_format = "0.0%"
        counts = [int(group["sex"].eq(value).sum()) if "sex" in group else 0 for _, value in sex_columns]
        sex_denominator = sum(counts)
        for (column, _), count in zip(sex_columns, counts):
            sheet.cell(row, column).value = f"{100 * count / sex_denominator:.1f}% ({count})" if sex_denominator else "0.0% (0)"


def _write_cross_sheet(sheet, sample: pd.DataFrame, major: str, minor: str, exact_counts: bool) -> None:
    form_header = _find_header_row(sheet, "Form")
    age_header = _find_header_row(sheet, "Age")
    form_pairs = list(_header_pairs(sheet, major, minor, form_header, form_header + 1))
    age_pairs = list(_header_pairs(sheet, major, minor, age_header, age_header + 1))
    current_form = ""
    current_age = ""
    for row in range(form_header + 2, age_header):
        first = str(sheet.cell(row, 1).value or "").strip().casefold()
        if first.startswith("form_1"):
            current_form = "form_1"
        elif first.startswith("form_2"):
            current_form = "form_2"
        if str(sheet.cell(row, 2).value or "").strip().casefold() == "sample" and current_form:
            _write_cross_row(sheet, row, sample.loc[sample["form"].eq(current_form)], form_pairs, major, minor, exact_counts)
    for row in range(age_header + 2, sheet.max_row + 1):
        if sheet.cell(row, 1).value is not None:
            current_age = _label(sheet.cell(row, 1).value)
        if str(sheet.cell(row, 2).value or "").strip().casefold() == "sample" and current_age:
            _write_cross_row(sheet, row, sample.loc[sample["_scope_age_group"].eq(current_age)], age_pairs, major, minor, exact_counts)


def _write_cross_row(sheet, row: int, group: pd.DataFrame, pairs, major: str, minor: str, exact_counts: bool) -> None:
    denominator = len(group)
    values = []
    for column, major_value, minor_value in pairs:
        count = int((group[major].eq(major_value) & group[minor].eq(minor_value)).sum()) if denominator else 0
        values.append((sheet.cell(row, column), count))
    if exact_counts:
        for cell, count in values:
            cell.value = f"{100 * count / denominator:.1f}% ({count})" if denominator else "0.0% (0)"
    else:
        for cell, count in values:
            # The reviewed workbook rounds displayed cross-tab percentages to
            # two percentage points before storing the Excel fraction.
            cell.value = round(100 * count / denominator, 2) / 100 if denominator else 0
            cell.number_format = "0.00%"


def _baseline_ids(sheet) -> set[str]:
    return {
        _clean_teid(sheet.cell(row, 1).value)
        for row in range(2, sheet.max_row + 1)
        if _clean_teid(sheet.cell(row, 1).value)
    }


def _numeric_flag(frame: pd.DataFrame, column: str) -> pd.Series:
    if column not in frame:
        return pd.Series(False, index=frame.index)
    return pd.to_numeric(frame[column], errors="coerce").fillna(0).eq(1)


def _group_masks(frame: pd.DataFrame, label: str) -> tuple[pd.Series, pd.Series, pd.Series]:
    canonical, raw_columns = GROUP_COLUMNS[label]
    official = _numeric_flag(frame, canonical)
    for column in raw_columns:
        official |= _numeric_flag(frame, column)
    report = pd.Series(False, index=frame.index)
    for prefix in ("SEED_", "REPORT_", "ELIGIBLE_"):
        report |= _numeric_flag(frame, f"{prefix}{canonical}")
    diagnosis = frame.get("diagnosis", pd.Series("", index=frame.index)).fillna("").astype(str)
    diagnosis_only = diagnosis.str.contains(DIAGNOSIS_PATTERNS[label], case=False, regex=True, na=False) & ~official if label in DIAGNOSIS_PATTERNS else pd.Series(False, index=frame.index)
    effective = report | official | diagnosis_only
    return effective, official & effective, diagnosis_only & effective


def _delta(frame: pd.DataFrame, baseline_ids: set[str], final_ids: set[str], label: str, form: str | None = None, age: str | None = None, kind: str = "effective") -> int:
    effective, official, diagnosis = _group_masks(frame, label)
    mask = {"effective": effective, "official": official, "diagnosis": diagnosis}[kind]
    if form is not None:
        mask &= frame["form"].eq(form)
    if age is not None:
        mask &= frame["_scope_age_group"].eq(age)
    added = frame["teid"].isin(final_ids - baseline_ids)
    removed = frame["teid"].isin(baseline_ids - final_ids)
    return int((mask & added).sum() - (mask & removed).sum())


def _write_osep_and_clinical(
    workbook,
    sample: pd.DataFrame,
    available: pd.DataFrame,
    baseline: set[str],
    baseline_form_n: dict[str, int],
) -> None:
    final_ids = set(sample["teid"])
    if final_ids == baseline:
        return
    universe = pd.concat([sample, available], ignore_index=True).drop_duplicates("teid", keep="first")
    sheet = workbook["osep"]
    final_form_n = sample["form"].value_counts().to_dict()
    osep_rows = {str(sheet.cell(row, 1).value): row for row in range(4, 9)}
    clinical_rows = {str(sheet.cell(row, 9).value): row for row in range(4, 8)}

    for label, row in osep_rows.items():
        for column, form in ((2, "form_1"), (3, "form_2")):
            base_count = round(float(sheet.cell(row, column).value or 0) * baseline_form_n[form])
            new_count = base_count + _delta(universe, baseline, final_ids, label, form=form)
            sheet.cell(row, column).value = new_count / int(final_form_n.get(form, 0)) if final_form_n.get(form, 0) else 0
            sheet.cell(row, column).number_format = "0.0%"
        base_count = round(float(sheet.cell(row, 4).value or 0) * len(baseline))
        sheet.cell(row, 4).value = (base_count + _delta(universe, baseline, final_ids, label)) / len(sample) if len(sample) else 0
        sheet.cell(row, 4).number_format = "0.0%"
    for column in (2, 3, 4):
        sheet.cell(9, column).value = sum(float(sheet.cell(row, column).value or 0) for row in osep_rows.values())
        sheet.cell(9, column).number_format = "0.0%"

    for label, row in clinical_rows.items():
        for column, form in ((10, "form_1"), (11, "form_2")):
            base_count = round(float(sheet.cell(row, column).value or 0) * baseline_form_n[form])
            new_count = base_count + _delta(universe, baseline, final_ids, label, form=form)
            sheet.cell(row, column).value = new_count / int(final_form_n.get(form, 0)) if final_form_n.get(form, 0) else 0
            sheet.cell(row, column).number_format = "0.0%"
        base_count = round(float(sheet.cell(row, 12).value or 0) * len(baseline))
        sheet.cell(row, 12).value = (base_count + _delta(universe, baseline, final_ids, label)) / len(sample) if len(sample) else 0
        sheet.cell(row, 12).number_format = "0.0%"
    for column in (10, 11, 12):
        sheet.cell(8, column).value = sum(float(sheet.cell(row, column).value or 0) for row in clinical_rows.values())
        sheet.cell(8, column).number_format = "0.0%"

    for row in range(14, 29):
        age = _label(sheet.cell(row, 1).value)
        for column, label in zip(range(2, 7), ("DD/ID", "SI", "ADHD", "SLD", "GT")):
            sheet.cell(row, column).value = int(sheet.cell(row, column).value or 0) + _delta(universe, baseline, final_ids, label, age=age)
    for column in range(2, 7):
        sheet.cell(29, column).value = sum(int(sheet.cell(row, column).value or 0) for row in range(14, 29))
    for row in range(13, 28):
        age = _label(sheet.cell(row, 9).value)
        for column, label in ((10, "LD"), (12, "HH"), (13, "ASD"), (14, "SLD-RW")):
            sheet.cell(row, column).value = int(sheet.cell(row, column).value or 0) + _delta(universe, baseline, final_ids, label, age=age)
    for column in (10, 12, 13, 14):
        sheet.cell(28, column).value = sum(int(sheet.cell(row, column).value or 0) for row in range(13, 28))

    sources = workbook["clinical_sources"]
    source_rows = {str(sources.cell(row, 1).value): row for row in range(3, sources.max_row)}
    for label, row in source_rows.items():
        if label not in GROUP_COLUMNS:
            continue
        sources.cell(row, 4).value = int(sources.cell(row, 4).value or 0) + _delta(universe, baseline, final_ids, label, kind="official")
        sources.cell(row, 5).value = int(sources.cell(row, 5).value or 0) + _delta(universe, baseline, final_ids, label, kind="diagnosis")
        sources.cell(row, 6).value = int(sources.cell(row, 4).value or 0) + int(sources.cell(row, 5).value or 0)
        shortfall = int(sources.cell(row, 3).value or 0) - int(sources.cell(row, 6).value or 0)
        sources.cell(row, 7).value = "Met" if shortfall <= 0 else f"Short by {shortfall}"


def _write_teids(sheet, sample: pd.DataFrame, available: pd.DataFrame) -> None:
    selected_ids = sample["teid"].tolist()
    available_ids = available["teid"].tolist()
    last_row = max(sheet.max_row, len(selected_ids) + 1, len(available_ids) + 1)
    for row in range(2, last_row + 1):
        sheet.cell(row, 1).value = None
        sheet.cell(row, 2).value = None
    sheet.cell(1, 1).value = f"Selected TEIDs (N={len(selected_ids)})"
    sheet.cell(1, 2).value = f"Available Pool TEIDs (N={len(available_ids)})"
    for index in range(max(len(selected_ids), len(available_ids))):
        if index < len(selected_ids):
            sheet.cell(index + 2, 1).value = selected_ids[index]
        if index < len(available_ids):
            sheet.cell(index + 2, 2).value = available_ids[index]
    sheet.auto_filter.ref = f"A1:B{max(len(selected_ids), len(available_ids)) + 1}"


def render_celf6_stratification_workbook(
    output: str | Path,
    sample: pd.DataFrame,
    available: pd.DataFrame,
    template: str | Path,
) -> Path:
    """Copy the reviewed workbook and refresh all case-dependent CELF-6 cells."""
    output_path, template_path = Path(output), Path(template)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    workbook = load_workbook(output_path)
    missing = [name for name in REPORT_SHEETS if name not in workbook.sheetnames]
    if missing:
        raise ValueError("CELF-6 stratification template is missing: " + ", ".join(missing))
    selected = _prepare(sample)
    pool = _prepare(available)
    baseline = _baseline_ids(workbook["TEID_List"])
    baseline_form_n = {
        "form_1": int(workbook["byage"]["B18"].value or 0),
        "form_2": int(workbook["byage"]["C18"].value or 0),
    }
    # The reviewed template already represents its embedded selected roster.
    # Avoid needlessly rewriting those exact values (some cells intentionally
    # retain more precision than their neighbors). Any different roster is
    # recalculated from participant data and deltas to that reviewed baseline.
    if set(selected["teid"]) != baseline:
        _write_byage(workbook["byage"], selected)
        _write_form_sheet(workbook["form_1"], selected, "form_1")
        _write_form_sheet(workbook["form_2"], selected, "form_2")
        for name, (major, minor, exact) in CROSS_TABS.items():
            _write_cross_sheet(workbook[name], selected, major, minor, exact)
        _write_osep_and_clinical(workbook, selected, pool, baseline, baseline_form_n)
    _write_teids(workbook["TEID_List"], selected, pool)
    workbook.active = workbook.sheetnames.index("byage")
    workbook.save(output_path)
    return output_path
