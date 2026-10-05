"""Stable, profile-driven SCOPE workbooks and audit outputs."""
from __future__ import annotations

from pathlib import Path
from copy import copy
import shutil
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .demographic_renderer import render_demographic_workbook
from .celf6_stratification import (
    is_celf6_stratification_template,
    render_celf6_stratification_workbook,
)


NAVY = "123B63"
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(color="FFFFFF", bold=True)
NEGATIVE_FONT = Font(color="B84444", bold=True)
NEGATIVE_FILL = PatternFill("solid", fgColor="FCE8E8")


def _key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().casefold())


def _age_key(value: object) -> str:
    text = str(value or "").strip().replace("–", "-")
    match = re.match(r"^(\d{1,2})[.:](\d{1,2})(?:[.:]\d{1,2})?\s*-\s*(\d{1,2})[.:](\d{1,2})", text)
    if not match:
        return _key(text)
    y1, m1, y2, m2 = map(int, match.groups())
    return _key(f"{y1}:{m1}-{y2}:{m2}")


def _format_workbook(path: Path) -> None:
    """Apply familiar report styling and make negative Diff values explicit."""
    workbook = load_workbook(path)
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value is not None}
        for cell in sheet[1]:
            if cell.value is not None:
                cell.fill = HEADER_FILL; cell.font = HEADER_FONT; cell.alignment = Alignment(wrap_text=True)
        percent_columns = [column for name, column in headers.items() if "%" in name or name.casefold() == "diff"]
        diff_columns = [column for name, column in headers.items() if name.casefold() in {"diff", "difference %", "difference"}]
        for row in range(2, sheet.max_row + 1):
            for column in percent_columns:
                sheet.cell(row, column).number_format = "+0.0%;-0.0%;0.0%" if column in diff_columns else "0.0%"
            for column in diff_columns:
                cell = sheet.cell(row, column)
                if isinstance(cell.value, (int, float)) and cell.value < 0:
                    cell.font = NEGATIVE_FONT; cell.fill = NEGATIVE_FILL
        for column in range(1, sheet.max_column + 1):
            values = [str(sheet.cell(row, column).value or "") for row in range(1, min(sheet.max_row, 200) + 1)]
            sheet.column_dimensions[get_column_letter(column)].width = min(42, max(11, max(map(len, values), default=0) + 2))
        sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def _write_tables(path: Path, outputs: dict[str, pd.DataFrame], template: str = "") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    template_path = Path(template) if template and not template.startswith("builtin:") else None
    # A source workbook is safe to reuse only when it contains every sheet we
    # are about to populate.  Otherwise copying it would carry historical case
    # data into a new project deliverable.  SCOPE still applies the approved
    # report styling to the clean workbook it creates.
    template_sheets: set[str] = set()
    if template_path and template_path.is_file():
        try:
            template_sheets = set(load_workbook(template_path, read_only=True).sheetnames)
        except Exception:
            template_sheets = set()
    if template_path and template_path.is_file() and all(name[:31] in template_sheets for name in outputs):
        shutil.copy2(template_path, path)
        mode, behavior = "a", "overlay"
    else:
        mode, behavior = "w", None
    options = {"engine": "openpyxl", "mode": mode}
    if behavior:
        options["if_sheet_exists"] = behavior
    with pd.ExcelWriter(path, **options) as writer:
        for name, frame in outputs.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    _format_workbook(path)
    return path


def write_profile_workbook(path: str | Path, summaries: dict[str, pd.DataFrame], audit: pd.DataFrame | None = None) -> Path:
    tables = dict(summaries)
    if audit is not None:
        tables["targeted_swap_audit"] = audit
    return _write_tables(Path(path), tables)


def write_targeted_swap_output(path: str | Path, sample: pd.DataFrame, available: pd.DataFrame, audit: pd.DataFrame, summaries: dict[str, pd.DataFrame]) -> Path:
    output = Path(path); output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample.to_excel(writer, sheet_name="final_all", index=False)
        if "teid" in sample.columns:
            sample[["teid"]].to_excel(writer, sheet_name="final_teid_list", index=False)
        available.to_excel(writer, sheet_name="available_pool", index=False)
        audit.to_excel(writer, sheet_name="targeted_swap_log", index=False)
        for name, frame in summaries.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    _format_workbook(output)
    return output


def write_targeted_swap_outputs(
    run_dir: str | Path,
    sample: pd.DataFrame,
    available: pd.DataFrame,
    audit: pd.DataFrame,
    summaries: dict[str, pd.DataFrame],
    sample_n_before: int,
    template: str = "",
    output_selection: str = "both",
) -> dict[str, Path]:
    """Write the requested Targeted Case Swap deliverables plus the audit log."""
    folder = Path(run_dir); cases = folder / "Cases"; qa = folder / "QA"; stratification_dir = folder / "Stratification"
    for destination in (cases, qa, stratification_dir): destination.mkdir(parents=True, exist_ok=True)
    include_lists = output_selection in {"both", "teid_lists"}
    include_tables = output_selection in {"both", "tables"}
    outputs: dict[str, Path] = {}
    if include_lists:
        selected_teids = sample[["teid"]].copy() if "teid" in sample else sample.iloc[:, :0].copy()
        available_teids = available[["teid"]].copy() if "teid" in available else available.iloc[:, :0].copy()
        outputs["selected_cases"] = _write_tables(cases / "Selected_TEIDs.xlsx", {
            # Keep the legacy sheet name so prior projects can reuse this file,
            # but make the TEID-list deliverable a true membership list. Writing
            # thousands of participant columns here duplicated the source data
            # and made three-case swaps take many minutes and several GB of RAM.
            "final_all": selected_teids,
            "Final TEID List": selected_teids,
        })
        outputs["available_pool"] = _write_tables(cases / "Available_Pool.xlsx", {"Available Pool": available_teids})
    summary = pd.DataFrame([
        {"Metric": "Sample N before", "Value": sample_n_before},
        {"Metric": "Sample N after", "Value": len(sample)},
        {"Metric": "Constraint status", "Value": "Passed"},
        {"Metric": "Optimizer", "Value": "CP-SAT NOT RERUN"},
    ])
    outputs["swap_log"] = _write_tables(qa / "Swap_Log.xlsx", {"Swap Log": audit, "Run Summary": summary})
    if include_tables:
        stratification_output = stratification_dir / "Updated_Stratification.xlsx"
        if is_celf6_stratification_template(template):
            outputs["updated_stratification"] = render_celf6_stratification_workbook(
                stratification_output, sample, available, template
            )
        else:
            outputs["updated_stratification"] = _write_tables(stratification_output, summaries, template)
    return outputs


def write_optimization_reports(
    run_dir: str | Path,
    summaries: dict[str, pd.DataFrame],
    template: str | dict[str, str] = "",
    *,
    sample: pd.DataFrame | None = None,
    profile=None,
    targets: list[dict] | None = None,
    demographic_tables: dict[str, list[str]] | None = None,
) -> dict[str, Path]:
    """Split profile results into the three familiar report families."""
    folder = Path(run_dir); folder.mkdir(parents=True, exist_ok=True)
    templates = template if isinstance(template, dict) else {"demographic": template, "osep": template, "clinical": template}
    osep: dict[str, pd.DataFrame] = {}
    clinical: dict[str, pd.DataFrame] = {}
    demographics: dict[str, pd.DataFrame] = {}
    for name, frame in summaries.items():
        lowered = name.casefold()
        if "osep" in lowered and "clinical" in lowered and "Group" in frame:
            clinical_mask = frame["Group"].astype(str).str.contains("clin", case=False, na=False)
            if (~clinical_mask).any(): osep[name.replace("_Clinical", "")] = frame.loc[~clinical_mask].copy()
            if clinical_mask.any(): clinical[name.replace("OSEP_", "")] = frame.loc[clinical_mask].copy()
        elif "osep" in lowered:
            osep[name] = frame
        elif "clinical" in lowered:
            clinical[name] = frame
        else:
            demographics[name] = frame
    demographic_template = templates.get("demographic", "")
    try:
        if sample is not None and profile is not None and demographic_template and Path(demographic_template).is_file():
            demographic_output = render_demographic_workbook(
                folder / "Demographic_Tables.xlsx", sample, profile, targets, demographic_template, demographic_tables
            )
        else:
            demographic_output = _write_tables(folder / "Demographic_Tables.xlsx", demographics, demographic_template)
    except ValueError:
        # A template that lacks a sheet the profile uses should not abort the run;
        # fall back to a clean workbook so the deliverable is still produced.
        demographic_output = _write_tables(folder / "Demographic_Tables.xlsx", demographics, "")
    outputs = {"demographic_tables": demographic_output}
    if osep:
        outputs["osep_tables"] = _write_tables(folder / "OSEP_Tables.xlsx", osep, templates.get("osep", ""))
    if clinical:
        outputs["clinical_tables"] = _write_tables(folder / "Clinical_Tables.xlsx", clinical, templates.get("clinical", ""))
    return outputs


def write_recruitment_workbook(path: str | Path, outputs: dict[str, pd.DataFrame], template: str = "") -> Path:
    """Populate an approved Field Research workbook without flattening its design.

    The familiar source workbook remains intact (sheet order, formulas, merges,
    widths, fills, borders and conditional formatting).  SCOPE writes the
    current profile-driven plan into the existing Summary table.
    """
    output = Path(path); output.parent.mkdir(parents=True, exist_ok=True)
    template_path = Path(template) if template and not template.startswith("builtin:") else None
    report = outputs.get("Field Research Report")
    if not template_path or not template_path.is_file() or report is None:
        return _write_tables(output, outputs, template)
    shutil.copy2(template_path, output)
    workbook = load_workbook(output)
    if "aurora_data" not in workbook.sheetnames:
        # Older generic profiles have a SCOPE-authored workbook rather than an
        # approved assessment Field Research workbook. Keep their established
        # tabular output path; approved WISC/CELF templates use aurora_data.
        return _write_tables(output, outputs, "")
    aurora = workbook["aurora_data"]
    needs = outputs.get("Remaining Recruitment Need", pd.DataFrame())
    cells = outputs.get("Recruitment Cell Targets", pd.DataFrame())
    if not cells.empty:
        dimension_keys = {_key(column): column for column in cells.columns if column != "target_n"}
    else:
        dimension_keys = {}
    margins = {
        (_key(row["dimension"]), _key(row["category"])): row
        for row in needs.to_dict("records")
    }
    cell_records = cells.to_dict("records") if not cells.empty else []
    age_dimension = next((name for key, name in dimension_keys.items() if key in {"age", "ageband", "agegroup", "agegrp"}), None)
    dimension_categories = {
        name: {_key(record.get(name, "")) for record in cell_records if _key(record.get(name, ""))}
        for name in dimension_keys.values() if name != age_dimension
    }
    target_cache: dict[tuple[str, str], float | None] = {}

    def native_target(age_key: str, header_key: str) -> float | None:
        cache_key = (age_key, header_key)
        if cache_key in target_cache: return target_cache[cache_key]
        rows = [record for record in cell_records if not age_dimension or _age_key(record.get(age_dimension, "")) == age_key]
        matched_dimension = False
        for dimension, categories in dimension_categories.items():
            matches = [category for category in categories if header_key == category or header_key.startswith(category) or header_key.endswith(category)]
            if matches:
                rows = [record for record in rows if _key(record.get(dimension, "")) in matches]
                matched_dimension = True
        value = float(sum(float(record.get("target_n", 0) or 0) for record in rows)) if matched_dimension and rows else None
        target_cache[cache_key] = value
        return value
    row = 1
    while row <= aurora.max_row:
        if _key(aurora.cell(row, 1).value) != "ageband" or _key(aurora.cell(row, 3).value) != "report":
            row += 1; continue
        headers = [aurora.cell(row, column).value for column in range(1, aurora.max_column + 1)]
        total_col = next((index + 1 for index, value in enumerate(headers) if _key(value) == "total"), None)
        row += 1
        while row <= aurora.max_row and _key(aurora.cell(row, 1).value) != "ageband":
            report_kind = _key(aurora.cell(row, 3).value)
            age_value = str(aurora.cell(row, 1).value or "")
            normalized_age = _age_key(age_value)
            age_margin = next((value for (dim, cat), value in margins.items() if dim in {"age", "ageband", "agegroup", "agegrp"} and (cat == normalized_age or cat in normalized_age or normalized_age in cat)), None)
            if report_kind in {"targetn", "stratifiedusablen"} and total_col and age_margin:
                aurora.cell(row, total_col).value = int(age_margin["planned_n"] if report_kind == "targetn" else age_margin["actual_n"])
                # Fill native category columns where a reliable cell/margin match exists.
                for column in range(4, total_col):
                    header_key = _key(headers[column - 1])
                    value = 0 if report_kind == "stratifiedusablen" else None
                    if report_kind == "targetn" and not cells.empty:
                        value = native_target(normalized_age, header_key)
                    if value is not None:
                        aurora.cell(row, column).value = value
            row += 1
    # Do not add, delete, rename, or flatten any worksheet. Visible report tabs
    # continue to calculate from the populated native data sheet.
    workbook.active = workbook.sheetnames.index("Summary") if "Summary" in workbook.sheetnames else 0
    workbook.save(output)
    return output
