"""Canonical SCOPE inputs shared by GUI validation and Python execution.

Assessment-specific source roles and field mappings live in Test Profiles.  This
module normalizes those sources without changing the validated optimizer math.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
import re
from typing import Any
import zipfile
from xml.etree import ElementTree as ET

import pandas as pd

from .configure import ProjectConfig
from .profiles import (
    TARGET_DIMENSION_ALIASES, InputRequirement, TestProfile,
    derive_marginal_target_rows, dimension_semantic, profile_optimization_rules,
    profile_sample_filters, target_row_criteria,
)


ROOT = Path(__file__).resolve().parents[1]
SCOPE_FIELDS = {
    "teid": "SCOPE_ID", "form": "SCOPE_FORM", "age": "SCOPE_AGE",
    "age_group": "SCOPE_AGE_GROUP", "sex": "SCOPE_SEX",
    "ethnicity": "SCOPE_ETHNICITY", "education": "SCOPE_EDUCATION",
    "region": "SCOPE_REGION", "language": "SCOPE_LANGUAGE",
    "study_type": "SCOPE_STUDY_TYPE",
}
FALLBACK_FIELDS = {
    "teid": ("teid", "scope_id", "participant_id", "id"),
    "form": ("vform", "form", "scope_form"),
    "age": ("agey", "age", "scope_age"),
    "age_group": ("agegrp", "age_group", "age band", "target age band", "scope_age_group"),
    "sex": ("sex", "gender", "scope_sex"),
    "ethnicity": ("eth", "ethnicity", "race_ethnicity", "scope_ethnicity"),
    "education": ("edl", "ped", "sed", "ped/sed", "education", "scope_education"),
    "region": ("region", "reg", "scope_region"),
    "language": ("lang", "testlang", "language", "scope_language"),
    "study_type": ("studytype", "study_type", "study", "scope_study_type"),
}
TARGET_COLUMNS = ("category", "group", "rule", "target", "unit", "tolerance", "priority", "required", "source", "criteria", "profile_value", "override", "override_reason")


@dataclass
class CanonicalInputs:
    case_universe: pd.DataFrame
    current_sample: pd.DataFrame
    available_pool: pd.DataFrame
    targets: pd.DataFrame
    universe_source: str
    universe_deferred: bool = False
    notes: list[str] = field(default_factory=list)


def _key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().casefold())


def _read_table(path: str | Path, preferred: tuple[str, ...] = ()) -> pd.DataFrame:
    source = Path(path)
    suffix = source.suffix.casefold()
    if suffix == ".csv":
        return pd.read_csv(source, dtype=str)
    if suffix in {".sas7bdat", ".xpt"}:
        return pd.read_sas(source, format="xport" if suffix == ".xpt" else "sas7bdat", encoding="latin1")
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported data file type '{suffix or '(none)'}': {source.name}")
    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else "xlrd"
    with pd.ExcelFile(source, engine=engine) as book:
        sheet = next((name for name in preferred if name and name in book.sheet_names), book.sheet_names[0])
        return pd.read_excel(book, sheet_name=sheet, dtype={"teid": str})


def _workbook_sheet(path: str | Path, names: tuple[str, ...]) -> pd.DataFrame | None:
    source = Path(path)
    if source.suffix.casefold() not in {".xlsx", ".xls", ".xlsm"}:
        return None
    with pd.ExcelFile(source, engine="openpyxl") as book:
        sheet = next((name for name in names if name in book.sheet_names), None)
        return pd.read_excel(book, sheet_name=sheet, dtype={"teid": str}) if sheet else None


def canonicalize_cases(
    frame: pd.DataFrame,
    profile: TestProfile,
    *,
    source_file: str = "",
    source_type: str = "",
    source_role: str = "",
    source_priority: int | str | None = None,
) -> pd.DataFrame:
    """Translate one assessment table into stable SCOPE_* concepts."""
    source = frame.copy()
    indexed = {_key(column): column for column in source.columns}
    result = pd.DataFrame(index=source.index)
    mapping = profile.schema_mapping or {}
    for semantic, output in SCOPE_FIELDS.items():
        configured = mapping.get(semantic, "")
        candidates = (configured, *FALLBACK_FIELDS[semantic])
        raw = next((indexed[_key(name)] for name in candidates if name and _key(name) in indexed), None)
        if raw is not None:
            result[output] = source[raw]
    for raw_flag in profile.target_flags:
        match = indexed.get(_key(raw_flag))
        if match is not None:
            canonical = "SCOPE_GROUP_" + re.sub(r"[^A-Z0-9]+", "_", raw_flag.upper()).strip("_")
            result[canonical] = pd.to_numeric(source[match], errors="coerce").fillna(0).astype(int)
    if "SCOPE_ID" in result:
        result["SCOPE_ID"] = result["SCOPE_ID"].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
        result = result.loc[result.SCOPE_ID.notna() & result.SCOPE_ID.ne("")].copy()
    if source_file:
        result["SCOPE_SOURCE_FILE"] = Path(source_file).name
        result["SCOPE_SOURCE_TYPE"] = str(source_type or "Unspecified")
        result["SCOPE_SOURCE_ROLE"] = str(source_role or "Candidate")
        result["SCOPE_SOURCE_PRIORITY"] = int(source_priority or 999)
    return result.reset_index(drop=True)


def configured_candidate_sources(config: ProjectConfig) -> list[dict[str, Any]]:
    """Return normalized candidate sources while retaining old .scope compatibility."""
    sources = [dict(item) for item in (config.candidate_datasets or []) if str(item.get("path", "")).strip()]
    # The editable Master / STDZ selector is authoritative for the first
    # source. Older GUI builds could retain its previous path in this list.
    if sources and str(config.dataset or "").strip():
        sources[0]["path"] = str(config.dataset).strip()
    if not sources and config.dataset:
        sources = [{"path": config.dataset, "source_type": config.study_type or "STDZ", "source_role": "Primary", "priority": 1}]
    for index, source in enumerate(sources, start=1):
        source["path"] = str(source.get("path", "")).strip()
        source["source_type"] = str(source.get("source_type") or "Unspecified")
        source["source_role"] = str(source.get("source_role") or ("Primary" if index == 1 else "Supplemental"))
        try: source["priority"] = int(source.get("priority", index))
        except (TypeError, ValueError): source["priority"] = index
    return sources


def _combine_candidate_sources(config: ProjectConfig, profile: TestProfile) -> tuple[pd.DataFrame, list[str]]:
    sources = configured_candidate_sources(config)
    if not sources:
        return pd.DataFrame(), []
    parts: list[pd.DataFrame] = []
    notes: list[str] = []
    demographic_semantics = {"eth": "ethnicity", "raceeth": "ethnicity", "edl": "education", "ped": "education"}
    required_semantics = ["teid"] if profile.uses_generic_optimizer else [
        "teid", "form", "age_group", *(demographic_semantics.get(value.casefold(), value.casefold()) for value in profile.demographics)
    ]
    for source in sources:
        path = Path(source["path"])
        if not path.is_file():
            raise ValueError(f"Candidate dataset was not found: {path}")
        canonical = canonicalize_cases(
            _read_table(path, ("case_universe", "master_prepped", "final_all", "selected_teids", "usable_pool", "available_pool")),
            profile, source_file=str(path), source_type=source["source_type"],
            source_role=source["source_role"], source_priority=source["priority"],
        )
        if "SCOPE_ID" not in canonical or canonical.empty:
            raise ValueError(f"{path.name} does not contain usable participant IDs using the selected Test Profile mapping.")
        missing = [SCOPE_FIELDS[name] for name in required_semantics if name in SCOPE_FIELDS and SCOPE_FIELDS[name] not in canonical]
        if missing:
            raise ValueError(f"Candidate dataset {path} is not coding-compatible with {profile.display_name}; missing mapped fields: {', '.join(missing)}. Replace this entry in Candidate case data / TEID list with a prepared dataset containing these fields.")
        parts.append(canonical)
        notes.append(f"{source['source_role']} dataset {path.name}: {canonical.SCOPE_ID.nunique():,} unique case(s)")
    combined = pd.concat(parts, ignore_index=True, sort=False)
    duplicates = combined.loc[combined.SCOPE_ID.duplicated(keep=False)].copy()
    if not duplicates.empty:
        rule = str((profile.sampling_specification or {}).get("duplicate_resolution", "source_priority")).casefold()
        if rule not in {"source_priority", "priority", "first"}:
            ids = duplicates.SCOPE_ID.drop_duplicates().head(10).tolist()
            raise ValueError(f"Duplicate participant IDs occur across candidate sources and the profile has no supported resolution rule: {', '.join(ids)}")
        combined = combined.sort_values(["SCOPE_SOURCE_PRIORITY", "SCOPE_SOURCE_FILE"], kind="stable").drop_duplicates("SCOPE_ID", keep="first")
        notes.append(f"Resolved {duplicates.SCOPE_ID.nunique():,} duplicate participant ID(s) using source priority.")
    else:
        notes.append("0 unresolved duplicate participant IDs")
    return combined.reset_index(drop=True), notes


def input_requirement_checks(config: ProjectConfig, profile: TestProfile) -> list[tuple[str, bool, str]]:
    checks: list[tuple[str, bool, str]] = []
    for requirement in profile.requirements_for(config.workflow):
        if requirement.role == "master_dataset":
            sources = configured_candidate_sources(config)
            value = sources[0]["path"] if sources else ""
        elif requirement.role == "target_specification":
            resolved = resolved_profile_input(config, profile, requirement.role)
            value = str(resolved or "")
        else:
            value = config.input_value(requirement.role)
        exists = Path(value).is_dir() if requirement.kind == "folder" and value else Path(value).is_file() if value else False
        label = f"{requirement.group.title()} — {requirement.label}"
        if value and not exists:
            checks.append((label, False, f"Selected {requirement.kind} was not found: {value}"))
        elif requirement.required:
            checks.append((label, exists, value if exists else "Required input is missing"))
        elif requirement.status.casefold() == "conditional":
            checks.append((label, True, value if exists else "Conditional; not supplied for this project"))
        else:
            checks.append((label, True, value if exists else "Optional; not supplied"))
    return checks


def resolved_profile_input(config: ProjectConfig, profile: TestProfile, role: str) -> Path | None:
    """Resolve a project override first, then an approved reusable profile resource."""
    project_value = config.input_value(role)
    if project_value:
        return Path(project_value)
    resource_names = {
        "target_specification": ("approved_targets", "census_targets"),
        "recruitment_targets": ("recruitment_targets",),
    }.get(role, ())
    return next((path for name in resource_names if (path := profile.resource_path(name))), None)


def _matches_project_value(series: pd.Series, selected: str, semantic: str) -> pd.Series:
    values = series.fillna("").astype(str).str.strip().str.casefold()
    wanted = str(selected or "").strip().casefold()
    if not wanted:
        return pd.Series(True, index=series.index)
    aliases = {
        "study_type": {
            "stdz": ("stdz", "standardization", "standardisation", "standard"),
            "clinical": ("clinical",),
            "validity": ("validity", "valid"),
        },
        "language": {
            "english": ("english", "eng", "e"),
            "spanish": ("spanish", "spa", "s"),
        },
    }
    candidates = aliases.get(semantic, {}).get(wanted, (wanted,))
    return values.apply(lambda value: any(candidate == value or candidate in value for candidate in candidates))


def _derive_initial_sample(universe: pd.DataFrame, config: ProjectConfig, profile: TestProfile) -> tuple[pd.DataFrame, list[str]]:
    """Derive first-run membership from profile-owned fields and target cells."""
    selected = universe.copy()
    notes: list[str] = []
    if "SCOPE_SOURCE_ROLE" in selected:
        primary = selected.loc[selected.SCOPE_SOURCE_ROLE.astype(str).str.casefold().eq("primary")].copy()
        if not primary.empty:
            selected = primary
            notes.append("Initial sample membership was derived from Primary sources; Supplemental sources remain eligible candidates.")
    if config.forms and "SCOPE_FORM" in selected:
        selected = selected.loc[selected.SCOPE_FORM.astype(str).str.upper().isin({value.upper() for value in config.forms})].copy()
    for column, value, semantic in (("SCOPE_LANGUAGE", config.language, "language"), ("SCOPE_STUDY_TYPE", config.study_type, "study_type")):
        if column in selected and selected[column].notna().any():
            matched = selected.loc[_matches_project_value(selected[column], value, semantic)].copy()
            if not matched.empty:
                selected = matched
            else:
                notes.append(f"No {column} values matched '{value}'; membership was not filtered by that field.")
    if config.fr_teid_file and Path(config.fr_teid_file).is_file():
        ids = canonicalize_cases(_read_table(config.fr_teid_file), profile)
        if "SCOPE_ID" not in ids:
            raise ValueError(f"The All Cases / TEID source does not contain the profile ID field: {Path(config.fr_teid_file).name}")
        selected = selected.loc[selected.SCOPE_ID.isin(set(ids.SCOPE_ID))].copy()
        notes.append(f"Initial membership was derived from {Path(config.fr_teid_file).name}.")
        return selected, notes
    if profile.source_config and {"SCOPE_FORM", "SCOPE_AGE_GROUP"}.issubset(selected.columns):
        import yaml
        config_path = ROOT / "scope" / "configs" / f"{profile.source_config}.yaml"
        payload = yaml.safe_load(config_path.read_text(encoding="utf-8")) if config_path.is_file() else {}
        target_cells = (payload or {}).get("target_cell_n") or []
        if target_cells and len(selected) > sum(int(item["n"]) for item in target_cells):
            parts = []
            for item in target_cells:
                mask = selected.SCOPE_FORM.astype(str).str.upper().eq(str(item["vform"]).upper())
                mask &= pd.to_numeric(selected.SCOPE_AGE_GROUP, errors="coerce").eq(int(item["agegrp"]))
                parts.append(selected.loc[mask].sort_values("SCOPE_ID").head(int(item["n"])))
            selected = pd.concat(parts, ignore_index=True).drop_duplicates("SCOPE_ID") if parts else selected
            notes.append("Initial membership was deterministically derived from the Test Profile's form/age target cells.")
    elif config.sample_size and len(selected) > config.sample_size:
        selected = selected.sort_values("SCOPE_ID").head(config.sample_size).copy()
        notes.append("Initial membership was deterministically limited to Target N.")
    else:
        notes.append("Initial membership was derived from the profile-filtered master dataset.")
    return selected.reset_index(drop=True), notes


def _read_scored_pool(config: ProjectConfig, profile: TestProfile) -> tuple[pd.DataFrame, list[str]]:
    if not config.scored_dir or not Path(config.scored_dir).is_dir() or not profile.source_config:
        return pd.DataFrame(), []
    import yaml
    config_path = ROOT / "scope" / "configs" / f"{profile.source_config}.yaml"
    payload = yaml.safe_load(config_path.read_text(encoding="utf-8")) if config_path.is_file() else {}
    rules = (payload or {}).get("pool_rules") or {}
    has_declared_source = any(
        (Path(config.scored_dir) / str(rule.get("file", ""))).is_file()
        or Path(str(Path(config.scored_dir) / str(rule.get("file", ""))) + ".parquet").is_file()
        for rule in rules.values()
    )
    if not has_declared_source:
        return pd.DataFrame(), ["No profile-declared scored files were found in the selected folder."]
    parts, notes = [], []
    for form, rule in rules.items():
        path = Path(config.scored_dir) / str(rule.get("file", ""))
        try:
            if path.is_file():
                raw = pd.read_sas(path, encoding="latin1")
            elif Path(str(path) + ".parquet").is_file():
                raw = pd.read_parquet(Path(str(path) + ".parquet"))
            else:
                raise FileNotFoundError(path.name)
        except Exception as exc:
            raise ValueError(f"Could not prepare scored-data source {path.name}: {exc}") from exc
        if not any(_key(column) == _key((profile.schema_mapping or {}).get("form", "vform")) for column in raw.columns):
            raw = raw.copy(); raw[(profile.schema_mapping or {}).get("form", "vform")] = form
        parts.append(canonicalize_cases(raw, profile))
        notes.append(f"{form}: {len(raw):,} scored record(s) loaded")
    if not parts:
        return pd.DataFrame(), notes
    pool = pd.concat(parts, ignore_index=True, sort=False)
    if "SCOPE_ID" in pool:
        pool = pool.drop_duplicates("SCOPE_ID")
    return pool.reset_index(drop=True), notes


def _filter_generic_candidate_universe(
    universe: pd.DataFrame,
    config: ProjectConfig,
    profile: TestProfile | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    """Apply project eligibility selections to the rows CP-SAT can select.

    These filters must constrain the complete solver universe, not only the
    optional starting sample.  A missing mapped field remains non-blocking for
    pre-filtered files, but any present field is enforced exactly as selected.
    Profile-owned sample rules (for example Study = STDZ for an international
    adaptation) are applied the same way and labelled as profile rules.
    """
    eligible = universe.copy()
    notes: list[str] = []
    filters: list[tuple[str, object, str, str, str]] = [
        ("SCOPE_FORM", config.forms, "form", "Forms", "include"),
        ("SCOPE_LANGUAGE", config.language, "language", "Language", "include"),
        ("SCOPE_STUDY_TYPE", config.study_type, "study_type", "Study Type", "include"),
    ]
    if profile is not None:
        for rule in profile_sample_filters(profile):
            column = SCOPE_FIELDS.get(str(rule.get("field", "")))
            if column:
                filters.append((column, rule.get("values", []), str(rule.get("field", "")), f"{rule.get('label', rule.get('field'))} (profile rule)", str(rule.get("operator", "include"))))
    for column, selected, semantic, label, mode in filters:
        values = list(selected) if isinstance(selected, (list, tuple, set)) else [selected]
        values = [str(value).strip() for value in values if str(value or "").strip()]
        if not values:
            continue
        if column not in eligible or not eligible[column].notna().any():
            mapped = (profile.schema_mapping or {}).get(semantic, "") if profile is not None else ""
            detail = f"no mapped {label.casefold()} field" + (f" (the profile maps {label.casefold()} to '{mapped}', which is not a column in the candidate data)" if mapped and mapped != column else "")
            notes.append(f"Eligibility filter — {label}: not applied because the candidate source has {detail}")
            continue
        before = int(eligible.SCOPE_ID.nunique()) if "SCOPE_ID" in eligible else len(eligible)
        mask = pd.Series(False, index=eligible.index)
        for value in values:
            mask |= _matches_project_value(eligible[column], value, semantic)
        matched = eligible.loc[mask].copy() if mode != "exclude" else eligible.loc[~mask].copy()
        if matched.empty:
            if label.endswith("(profile rule)"):
                notes.append(
                    f"Eligibility filter — {label} {'≠' if mode == 'exclude' else '='} {', '.join(values)}: not applied because it would remove every candidate case. "
                    "Review the profile's field mapping (for example Language must map to the dataset's language column, not the word 'English')."
                )
                continue
            raise ValueError(
                f"The {label} filter ({', '.join(values)}) matched 0 candidate cases. "
                f"Review the project selection or the Test Profile mapping for {column}."
            )
        eligible = matched
        after = int(eligible.SCOPE_ID.nunique()) if "SCOPE_ID" in eligible else len(eligible)
        verb = "≠" if mode == "exclude" else "="
        notes.append(
            f"Eligibility filter — {label} {verb} {', '.join(values)}: retained {after:,} of {before:,} unique case(s); excluded {before - after:,}"
        )
    return eligible.reset_index(drop=True), notes


def build_case_inputs(config: ProjectConfig, profile: TestProfile) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str, bool, list[str]]:
    current = pd.DataFrame()
    if config.master_file and Path(config.master_file).is_file():
        current = canonicalize_cases(_read_table(config.master_file, (config.master_sheet, "final_all", "master_prepped", "selected_teids")), profile)

    notes: list[str] = []
    deferred = False
    candidate_universe, candidate_notes = _combine_candidate_sources(config, profile)
    notes.extend(candidate_notes)
    if not candidate_universe.empty:
        universe = candidate_universe
        source_count = len(configured_candidate_sources(config))
        source = f"{source_count} candidate dataset(s)"
        if current.empty and not profile.uses_generic_optimizer:
            current, derivation_notes = _derive_initial_sample(universe, config, profile)
            notes.extend(derivation_notes)
    else:
        available = _workbook_sheet(config.master_file, ("available_pool", "usable_pool")) if config.master_file else None
        if available is not None:
            universe = pd.concat([current, canonicalize_cases(available, profile)], ignore_index=True).drop_duplicates("SCOPE_ID")
            source = "Current sample + workbook available pool"
        else:
            universe = current.copy()
            source = "Current sample"
            deferred = bool(config.scored_dir)
            if deferred:
                notes.append("Additional eligible candidates will be canonicalized from scored data during SCOPE preparation.")
    scored, scored_notes = _read_scored_pool(config, profile)
    notes.extend(scored_notes)
    if not scored.empty and "SCOPE_ID" in scored:
        selected_ids = set(current.get("SCOPE_ID", pd.Series(dtype=str)))
        available = scored.loc[~scored.SCOPE_ID.isin(selected_ids)].copy()
        universe = pd.concat([universe, scored], ignore_index=True, sort=False).drop_duplicates("SCOPE_ID")
        deferred = False
    if profile.uses_generic_optimizer and not universe.empty:
        universe, eligibility_notes = _filter_generic_candidate_universe(universe, config, profile)
        notes.extend(eligibility_notes)
        if not current.empty and "SCOPE_ID" in current:
            eligible_ids = set(universe.SCOPE_ID.astype(str))
            before = current.SCOPE_ID.nunique()
            current = current.loc[current.SCOPE_ID.astype(str).isin(eligible_ids)].copy()
            excluded = before - current.SCOPE_ID.nunique()
            if excluded:
                notes.append(f"Existing membership eligibility: excluded {excluded:,} case(s) outside the filtered candidate universe")
        if current.empty and not candidate_universe.empty:
            current, derivation_notes = _derive_initial_sample(universe, config, profile)
            notes.extend(derivation_notes)
    selected_ids = set(current.get("SCOPE_ID", pd.Series(dtype=str)).astype(str))
    available = universe.loc[~universe.get("SCOPE_ID", pd.Series(index=universe.index, dtype=str)).astype(str).isin(selected_ids)].copy() if "SCOPE_ID" in universe else pd.DataFrame()
    return universe.reset_index(drop=True), current.reset_index(drop=True), available.reset_index(drop=True), source, deferred, notes


def _first_column(frame: pd.DataFrame, *names: str) -> str | None:
    indexed = {_key(column): column for column in frame.columns}
    return next((indexed[_key(name)] for name in names if _key(name) in indexed), None)


def _generic_target_rows(frame: pd.DataFrame, source: str) -> list[dict[str, Any]]:
    category = _first_column(frame, "category", "dimension", "target category")
    group = _first_column(frame, "group", "level", "value")
    if category and not group:
        # Recruitment-style tables use Dimension + Category.
        dimension = _first_column(frame, "dimension")
        if dimension and category != dimension:
            group, category = category, dimension
    rule = _first_column(frame, "rule", "rule type", "constraint")
    target = _first_column(frame, "target", "target value")
    target_pct = _first_column(frame, "target_pct", "target percent", "target %")
    target_n = _first_column(frame, "target_n", "target n")
    explicit_unit = _first_column(frame, "unit", "target unit")
    minimum = _first_column(frame, "minimum", "min", "min_pct", "minimum percent")
    maximum = _first_column(frame, "maximum", "max", "max_pct", "maximum percent")
    tolerance = _first_column(frame, "tolerance", "tolerance_n", "allowable range")
    priority = _first_column(frame, "priority", "weight")
    required = _first_column(frame, "required", "is required")
    if not category or not group or not any((target, target_pct, target_n, minimum, maximum)):
        return []
    rows: list[dict[str, Any]] = []
    for record in frame.to_dict("records"):
        category_value, group_value = record.get(category), record.get(group)
        if pd.isna(category_value) or pd.isna(group_value) or not str(category_value).strip() or not str(group_value).strip():
            continue
        if minimum and maximum and pd.notna(record.get(minimum)) and pd.notna(record.get(maximum)):
            rule_value, target_value, unit = "Range", f"{record[minimum]}–{record[maximum]}", "%"
        elif maximum and pd.notna(record.get(maximum)):
            rule_value, target_value, unit = "Maximum", record[maximum], "%"
        elif minimum and pd.notna(record.get(minimum)):
            rule_value, target_value, unit = "Minimum", record[minimum], "%"
        elif target_n and pd.notna(record.get(target_n)):
            rule_value, target_value, unit = str(record.get(rule) or "Target"), record[target_n], "N"
        elif target_pct and pd.notna(record.get(target_pct)):
            rule_value, target_value, unit = str(record.get(rule) or "Target"), record[target_pct], "%"
        elif target and pd.notna(record.get(target)):
            rule_value, target_value, unit = str(record.get(rule) or "Target"), record[target], str(record.get(explicit_unit, "") or "") if explicit_unit else ""
        else:
            continue
        rows.append({
            "category": str(category_value).strip(), "group": str(group_value).strip(),
            "rule": rule_value, "target": target_value, "unit": unit,
            "tolerance": record.get(tolerance, "") if tolerance else "",
            "priority": record.get(priority, "") if priority else "",
            "required": str(record.get(required, "true")).strip().casefold() not in {"false", "0", "no", "optional"} if required else True, "source": source,
        })
    return rows


def _fully_crossed_target_rows(frame: pd.DataFrame, source: str) -> list[dict[str, Any]]:
    """Read one-row-per-cell demographic targets such as the WAIS Canada Target List.

    These workbooks describe a joint distribution directly (age band x sex x
    region x ethnicity x education) instead of repeating Category/Group rows.
    Preserve the complete cell definition as JSON so the generic optimizer can
    match each candidate against the same dimensions without a Research Plan.
    """
    indexed = {_key(column): column for column in frame.columns}
    dimensions: list[tuple[str, str]] = []
    for semantic, aliases in TARGET_DIMENSION_ALIASES.items():
        raw = next((indexed[_key(alias)] for alias in aliases if _key(alias) in indexed), None)
        if raw is not None:
            dimensions.append((semantic, raw))
    count = _first_column(frame, "count", "target count", "target_n", "target n")
    percent = _first_column(frame, "pcnt", "percent", "target_pct", "target percent", "target %")
    if len(dimensions) < 2 or not (count or percent):
        return []
    rows: list[dict[str, Any]] = []
    for record in frame.to_dict("records"):
        criteria = {
            semantic: str(record.get(column)).strip()
            for semantic, column in dimensions
            if pd.notna(record.get(column)) and str(record.get(column)).strip()
        }
        if len(criteria) != len(dimensions):
            continue
        if count and pd.notna(record.get(count)):
            target, unit = record[count], "N"
        elif percent and pd.notna(record.get(percent)):
            target, unit = record[percent], "%"
        else:
            continue
        labels = " | ".join(f"{name}={value}" for name, value in criteria.items())
        rows.append({
            "category": " × ".join(name.replace("_", " ").title() for name, _column in dimensions),
            "group": labels,
            "rule": "Target",
            "target": target,
            "unit": unit,
            "tolerance": "",
            "priority": "Demographic",
            "required": True,
            "source": source,
            "criteria": json.dumps(criteria, sort_keys=True),
        })
    return rows


def _profile_config_targets(profile: TestProfile) -> list[dict[str, Any]]:
    if not profile.source_config:
        return []
    path = ROOT / "scope" / "configs" / f"{profile.source_config}.yaml"
    if not path.is_file():
        return []
    import yaml
    payload = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    rows: list[dict[str, Any]] = []
    for item in payload.get("target_cell_n", []):
        rows.append({"category": "Sample Cell", "group": f"{item['vform']} / age group {item['agegrp']}", "rule": "Exact", "target": item["n"], "unit": "N", "tolerance": 0, "priority": "Required", "required": True, "source": path.name})
    for name, spec in (payload.get("osep_targets") or {}).items():
        if "target_pct" in spec:
            rule, target = "Target", spec["target_pct"]
        elif "min_pct" in spec and "max_pct" in spec:
            rule, target = "Range", f"{spec['min_pct']}–{spec['max_pct']}"
        elif "min_pct" in spec:
            rule, target = "Minimum", spec["min_pct"]
        elif "max_pct" in spec:
            rule, target = "Maximum", spec["max_pct"]
        else:
            continue
        rows.append({"category": "OSEP / Clinical", "group": name, "rule": rule, "target": target, "unit": "%", "tolerance": "", "priority": spec.get("weight", ""), "required": True, "source": path.name})
    return rows


def _basc_demographic_targets(path: Path, profile: TestProfile) -> list[dict[str, Any]]:
    if not profile.source_config or path.suffix.casefold() not in {".xlsx", ".xlsm"}:
        return []
    try:
        from openpyxl import load_workbook
        import yaml
        config_path = ROOT / "scope" / "configs" / f"{profile.source_config}.yaml"
        payload = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        sheet_mapping = payload.get("demo_target_sheets") or {}
        sheet_names = list(sheet_mapping.keys()) + [value for value in sheet_mapping.values() if value not in sheet_mapping]
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    rows: list[dict[str, Any]] = []
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        headers = [sheet.cell(6, column).value for column in range(1, sheet.max_column + 1)]
        current_form = None
        for row_number in range(7, sheet.max_row + 1):
            if sheet.cell(row_number, 1).value:
                current_form = str(sheet.cell(row_number, 1).value).strip().upper()
            if sheet.cell(row_number, 4).value != "US population" or sheet.cell(row_number, 2).value is None:
                continue
            age_group = int(sheet.cell(row_number, 2).value)
            for column, header in enumerate(headers, start=1):
                value = sheet.cell(row_number, column).value
                if column <= 4 or not header or value is None:
                    continue
                header_text = str(header).strip()
                if not header_text.startswith(("edl_", "eth_", "region_")):
                    continue
                category, group = header_text.split("_", 1)
                rows.append({"category": category.upper(), "group": f"{current_form} / age group {age_group} / {group}", "rule": "Target", "target": float(value), "unit": "%", "tolerance": "", "priority": "Demographic", "required": True, "source": path.name})
    return rows


def _docx_tables(path: Path) -> list[pd.DataFrame]:
    """Read structured Word tables without routing the DOCX ZIP through pandas."""
    try:
        with zipfile.ZipFile(path) as archive:
            document = archive.read("word/document.xml")
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise ValueError(f"Could not read Research Plan DOCX '{path.name}': {exc}") from exc
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    root = ET.fromstring(document)
    frames: list[pd.DataFrame] = []
    for table in root.findall(".//w:tbl", namespace):
        rows: list[list[str]] = []
        for row in table.findall("./w:tr", namespace):
            cells = []
            for cell in row.findall("./w:tc", namespace):
                text = " ".join(value.strip() for value in cell.itertext() if value.strip())
                cells.append(text)
            if any(cells): rows.append(cells)
        if len(rows) < 2: continue
        width = max(len(row) for row in rows)
        padded = [row + [""] * (width - len(row)) for row in rows]
        headers = [value or f"Column {index + 1}" for index, value in enumerate(padded[0])]
        frames.append(pd.DataFrame(padded[1:], columns=headers))
    return frames


def _target_rows_from_path(path: Path, profile: TestProfile) -> list[dict[str, Any]]:
    suffix = path.suffix.casefold()
    rows: list[dict[str, Any]] = []
    if suffix == ".csv":
        rows.extend(_generic_target_rows(pd.read_csv(path), path.name))
    elif suffix == ".json":
        rows.extend(_generic_target_rows(pd.read_json(path), path.name))
    elif suffix == ".docx":
        for index, frame in enumerate(_docx_tables(path), start=1):
            rows.extend(_generic_target_rows(frame, f"{path.name} · table {index}"))
    elif suffix in {".xlsx", ".xlsm"}:
        with pd.ExcelFile(path, engine="openpyxl") as book:
            for sheet in book.sheet_names:
                frame = pd.read_excel(book, sheet_name=sheet)
                rows.extend(_generic_target_rows(frame, f"{path.name} · {sheet}"))
                rows.extend(_fully_crossed_target_rows(frame, f"{path.name} · {sheet}"))
        rows.extend(_basc_demographic_targets(path, profile))
    elif suffix == ".xls":
        with pd.ExcelFile(path, engine="xlrd") as book:
            for sheet in book.sheet_names:
                frame = pd.read_excel(book, sheet_name=sheet)
                rows.extend(_generic_target_rows(frame, f"{path.name} · {sheet}"))
                rows.extend(_fully_crossed_target_rows(frame, f"{path.name} · {sheet}"))
    else:
        raise ValueError(
            f"Unsupported target source '{path.name}'. Choose an XLSX/XLSM SCOPE Target Template, "
            "a structured DOCX Research Plan, CSV, or JSON file."
        )
    if not rows:
        raise ValueError(
            f"No canonical targets were found in '{path.name}'. Choose a structured target source "
            "or load the SCOPE Target Template."
        )
    return rows


def _target_family(row: dict[str, Any]) -> str:
    category = str(row.get("category", "")).casefold()
    group = str(row.get("group", "")).casefold()
    if "clin" in group or ("clinical" in category and "osep" not in category):
        return "clinical"
    if group.startswith("tmp_") or "osep" in category:
        return "osep"
    if row.get("criteria") or category in {"sex", "eth", "ethnicity", "region", "edl", "education", "demographic", "census"} or "×" in str(row.get("category", "")):
        return "census"
    return "other"


def target_criteria(row: dict[str, Any] | pd.Series) -> dict[str, str]:
    """Return canonical dimension/value criteria for one target row."""
    raw = row.get("criteria", "")
    if isinstance(raw, dict):
        return {str(key): str(value) for key, value in raw.items()}
    if raw and not pd.isna(raw):
        try:
            parsed = json.loads(str(raw))
            if isinstance(parsed, dict):
                return {str(key): str(value) for key, value in parsed.items()}
        except (TypeError, ValueError, json.JSONDecodeError):
            pass
    category = _key(row.get("category", ""))
    semantic = next((name for name, aliases in TARGET_DIMENSION_ALIASES.items() if category in {_key(name), *(_key(alias) for alias in aliases)}), None)
    group = str(row.get("group", "")).strip()
    return {semantic: group} if semantic and group else {}


def target_required_semantics(targets: pd.DataFrame) -> set[str]:
    required: set[str] = set()
    for row in targets.to_dict("records"):
        required.update(target_criteria(row))
    return required


def _apply_tolerance_rule(row: dict[str, Any], rules: list[dict]) -> dict[str, Any]:
    """Rewrite a 'Target' marginal row into the profile's rule (±% range, or exact N)."""
    if str(row.get("rule", "Target")).strip().casefold() not in {"target", ""}:
        return row
    criteria = target_row_criteria(row)
    semantics = set(criteria) or ({dimension_semantic(row.get("category", ""))} if dimension_semantic(row.get("category", "")) else set())
    group = next(iter(criteria.values()), "") if criteria else str(row.get("group", ""))
    spec = next((rule for rule in rules if rule.get("semantic") in semantics and rule.get("applies_to", "all") == "group" and rule.get("group") == group), None)
    if spec is None:
        spec = next((rule for rule in rules if rule.get("semantic") in semantics and rule.get("applies_to", "all") == "all"), None)
    if not spec:
        return row
    unit = str(row.get("unit", "")).strip()
    numbers = re.findall(r"-?\d+(?:\.\d+)?", str(row.get("target", "")))
    if not numbers:
        return row
    value = float(numbers[0])
    target_n = spec.get("target_n")
    tolerance_percent = spec.get("tolerance_percent")
    tolerance_n = spec.get("tolerance_n")
    updated = dict(row)
    if target_n not in (None, ""):
        updated.update({"rule": "Exact", "target": target_n, "unit": "N"})
    elif unit == "%" and tolerance_percent not in (None, ""):
        fraction = value / 100 if value > 1 else value
        pct = fraction * 100
        low, high = max(0.0, pct - float(tolerance_percent)), min(100.0, pct + float(tolerance_percent))
        updated.update({"rule": "Range", "target": f"{low:g}–{high:g}"})
    elif unit.upper() == "N" and tolerance_n not in (None, ""):
        width = float(tolerance_n)
        updated.update({"rule": "Range", "target": f"{max(0.0, value - width):g}–{value + width:g}"})
    else:
        return row
    if spec.get("priority"):
        updated["priority"] = spec["priority"]
    return updated


def _apply_profile_target_interpretation(rows: list[dict[str, Any]], profile: TestProfile, source_name: str = "") -> list[dict[str, Any]]:
    """Separate the source target format from the optimization/reporting format.

    Fully crossed cell-by-cell sources (for example a WAIS-5 Canada census
    workbook) are imported as-is, then collapsed into the marginal/group targets
    declared by the profile.  Raw cells only remain CP-SAT objective rows when
    the profile explicitly keeps them (cells_as = "objective").
    """
    interpretation = profile.target_interpretation or {}
    rules = profile_optimization_rules(profile)
    if not interpretation and not rules:
        return rows
    crossed_index = {index for index, row in enumerate(rows) if len(target_row_criteria(row)) >= 2}
    derived: list[dict[str, Any]] = []
    if interpretation and crossed_index:
        dimensions = [str(value) for value in interpretation.get("derive_marginals") or []]
        if dimensions:
            label = str(interpretation.get("derived_source_label", "")).strip() or (f"Derived from {source_name}" if source_name else "Derived marginal targets")
            derived = derive_marginal_target_rows([rows[index] for index in sorted(crossed_index)], dimensions, profile.target_tolerances or {}, source_label=label)
        cells_as = str(interpretation.get("cells_as", "report_only")).strip().casefold() or "report_only"
        if derived and cells_as not in {"objective", "include", "keep"}:
            rows = [row for index, row in enumerate(rows) if index not in crossed_index]
    rows = [*rows, *derived]
    if rules:
        rows = [_apply_tolerance_rule(row, rules) for row in rows]
    return rows


def build_target_specification(config: ProjectConfig, profile: TestProfile) -> pd.DataFrame:
    """Normalize supported target sources into the table shown to users and engines."""
    rows: list[dict[str, Any]] = []
    target_path = resolved_profile_input(config, profile, "target_specification")
    if config.workflow == "sample_optimization":
        if not target_path:
            raise ValueError("This Test Profile has no approved target resource. Contact the SCOPE administrator.")
        if not target_path.is_file():
            raise ValueError(f"The approved target resource is missing: {target_path}. Contact the SCOPE administrator.")
    if target_path and target_path.is_file():
        rows.extend(_target_rows_from_path(target_path, profile))
    rows.extend(_profile_config_targets(profile))
    rows = _apply_profile_target_interpretation(rows, profile, target_path.name if target_path else "")
    if config.sample_size:
        rows.append({"category": "Sample", "group": "Overall", "rule": "Exact", "target": config.sample_size, "unit": "N", "tolerance": config.target_tolerance, "priority": "Required", "required": True, "source": "Project Configuration"})
    for group, target in config.target_n_by_agegrp.items():
        rows.append({"category": "Age Group", "group": str(group), "rule": "Exact", "target": int(target), "unit": "N", "tolerance": config.target_tolerance, "priority": "Required", "required": True, "source": "Project Configuration"})
    selected_families = {
        "census": config.include_census, "osep": config.include_osep, "clinical": config.include_clinical,
    }
    rows = [row for row in rows if _target_family(row) == "other" or selected_families.get(_target_family(row), True)]
    overrides = {
        (_key(item.get("category", "")), _key(item.get("group", "")), _key(item.get("rule", ""))): item
        for item in (config.target_overrides or [])
    }
    for row in rows:
        identity = (_key(row.get("category", "")), _key(row.get("group", "")), _key(row.get("rule", "")))
        override = overrides.get(identity)
        if not override:
            continue
        row["profile_value"] = row.get("target", "")
        row["target"] = override.get("target", row.get("target", ""))
        if override.get("tolerance") not in (None, ""):
            row["tolerance"] = override["tolerance"]
        if override.get("priority") not in (None, ""):
            row["priority"] = override["priority"]
        row["source"] = "Project Override"
        row["override"] = True
        row["override_reason"] = str(override.get("reason", ""))
    unique: list[dict[str, Any]] = []
    seen = set()
    for row in rows:
        normalized = {column: row.get(column, "") for column in TARGET_COLUMNS}
        identity = tuple(str(normalized[column]) for column in ("category", "group", "rule", "target", "unit", "source"))
        if identity not in seen:
            unique.append(normalized); seen.add(identity)
    return pd.DataFrame(unique, columns=TARGET_COLUMNS)


def prepare_canonical_inputs(config: ProjectConfig, profile: TestProfile) -> CanonicalInputs:
    universe, current, available, source, deferred, notes = build_case_inputs(config, profile)
    targets = build_target_specification(config, profile)
    return CanonicalInputs(universe, current, available, targets, source, deferred, notes)


def materialize_initial_sample(config: ProjectConfig, profile: TestProfile) -> Path:
    """Create the legacy Part A master workbook for a first-run project."""
    if config.master_file and Path(config.master_file).is_file():
        return Path(config.master_file)
    sources = configured_candidate_sources(config)
    if not sources:
        raise ValueError("First-run optimization needs at least one Master / Candidate dataset.")
    prepared = prepare_canonical_inputs(config, profile)
    if prepared.current_sample.empty or "SCOPE_ID" not in prepared.current_sample:
        raise ValueError("SCOPE could not derive an initial sample from the selected case sources.")
    selected_ids = set(prepared.current_sample.SCOPE_ID.astype(str))
    raw_parts = []
    for source in sources:
        raw = _read_table(source["path"], ("case_universe", "master_prepped", "final_all", "teid_list"))
        indexed = {_key(column): column for column in raw.columns}
        configured = (profile.schema_mapping or {}).get("teid", "")
        id_column = next((indexed[_key(name)] for name in (configured, *FALLBACK_FIELDS["teid"]) if name and _key(name) in indexed), None)
        if id_column is None:
            raise ValueError(f"{Path(source['path']).name} does not contain the Test Profile's participant ID field.")
        raw = raw.copy()
        raw["SCOPE_SOURCE_FILE"] = Path(source["path"]).name
        raw["SCOPE_SOURCE_TYPE"] = source["source_type"]
        raw["SCOPE_SOURCE_ROLE"] = source["source_role"]
        raw["SCOPE_SOURCE_PRIORITY"] = source["priority"]
        raw["__scope_id"] = raw[id_column].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
        raw_parts.append(raw.loc[raw["__scope_id"].isin(selected_ids)].copy())
    selected_raw = pd.concat(raw_parts, ignore_index=True, sort=False).sort_values("SCOPE_SOURCE_PRIORITY", kind="stable").drop_duplicates("__scope_id").drop(columns="__scope_id")
    output = Path(config.project_dir) / "scope_inputs" / "SCOPE_Derived_Initial_Sample.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    selected_raw.to_excel(output, sheet_name="final_all", index=False)
    return output


def canonical_target_records(targets: pd.DataFrame) -> list[dict]:
    """Return stable JSON-native records for project persistence and equality QA."""
    if targets.empty:
        return []
    clean = targets.where(pd.notna(targets), "")
    return json.loads(clean.to_json(orient="records"))


def canonical_validation_checks(prepared: CanonicalInputs, profile: TestProfile, config: ProjectConfig | None = None) -> list[tuple[str, bool, str]]:
    universe, current, available, targets = prepared.case_universe, prepared.current_sample, prepared.available_pool, prepared.targets
    checks: list[tuple[str, bool, str]] = []
    profile_ok = profile.status == "Validated"
    checks.append(("Test Profile — Validated profile", profile_ok, f"{profile.display_name} v{profile.version} · {profile.status}"))
    id_ready = "SCOPE_ID" in universe and not universe.empty
    unique = id_ready and not universe.SCOPE_ID.duplicated().any()
    detail = f"N = {universe.SCOPE_ID.nunique():,} from {prepared.universe_source}" if id_ready else "No canonical participant IDs were prepared"
    if prepared.universe_deferred:
        detail += "; scored-data candidates are added during preparation"
    checks.extend([
        ("Case Universe — Cases prepared", id_ready, detail),
        ("Case Universe — Unique participant IDs", unique, "All SCOPE_ID values are unique" if unique else "Duplicate or missing SCOPE_ID values"),
    ])
    for note in prepared.notes:
        if note.startswith("Eligibility filter —"):
            label, _, result = note.partition(":")
            checks.append((label, True, result.strip()))
    if profile.uses_generic_optimizer:
        required_semantics = target_required_semantics(targets)
        for semantic in sorted(required_semantics):
            column = SCOPE_FIELDS.get(semantic)
            ready = bool(column and column in universe and universe[column].notna().any())
            checks.append((
                f"Case Universe — {semantic.replace('_', ' ').title()} target field",
                ready,
                column if ready else f"The candidate TEID list needs a {semantic.replace('_', ' ')} column because the target workbook uses that dimension",
            ))
    else:
        checks.extend([
            ("Case Universe — Forms recognized", bool("SCOPE_FORM" in universe and universe.SCOPE_FORM.notna().any()), ", ".join(sorted(universe.get("SCOPE_FORM", pd.Series(dtype=str)).dropna().astype(str).unique())[:12]) or "No form values mapped"),
            ("Case Universe — Age groups recognized", bool("SCOPE_AGE_GROUP" in universe and universe.SCOPE_AGE_GROUP.notna().any()), ", ".join(sorted(universe.get("SCOPE_AGE_GROUP", pd.Series(dtype=str)).dropna().astype(str).unique())[:15]) or "No age-group values mapped"),
        ])
    if {"SCOPE_SOURCE_FILE", "SCOPE_SOURCE_ROLE"}.issubset(universe.columns):
        source_count = int(universe.SCOPE_SOURCE_FILE.nunique())
        checks.append(("CASE DATA — Source datasets loaded", source_count > 0, f"{source_count} source dataset(s) loaded"))
        for (role, filename), rows in universe.groupby(["SCOPE_SOURCE_ROLE", "SCOPE_SOURCE_FILE"], sort=False):
            checks.append((f"CASE DATA — {role}: {filename}", True, f"{rows.SCOPE_ID.nunique():,} unique case(s) retained"))
        duplicate_note = next((note for note in prepared.notes if "duplicate participant" in note.casefold()), "0 unresolved duplicate participant IDs")
        checks.append(("CASE DATA — Duplicates", "unresolved" in duplicate_note.casefold() and duplicate_note.lstrip().startswith("0") or "resolved" in duplicate_note.casefold(), duplicate_note))
        checks.append(("CASE DATA — Candidate Universe", id_ready, f"{universe.SCOPE_ID.nunique():,} unique cases" if id_ready else "No candidate universe prepared"))
    expected_names = sorted(target_required_semantics(targets) & {"sex", "ethnicity", "education", "region"}) if profile.uses_generic_optimizer else ["sex", "ethnicity", "education", "region"]
    expected_demo = [SCOPE_FIELDS[name] for name in expected_names]
    demo_columns = [name for name in expected_demo if name in universe]
    checks.append(("Case Universe — Demographics mapped", len(demo_columns) == len(expected_demo), ", ".join(demo_columns) if len(demo_columns) == len(expected_demo) else f"Found {len(demo_columns)} of {len(expected_demo)}: {', '.join(demo_columns) or 'none'}"))
    group_columns = [column for column in universe if column.startswith("SCOPE_GROUP_")]
    groups_ready = len(group_columns) == len(profile.target_flags) if profile.target_flags else True
    checks.append(("Case Universe — OSEP/clinical fields", groups_ready, f"{len(group_columns)} of {len(profile.target_flags)} mapped" if profile.target_flags else "No profile group fields required"))
    current_ready = "SCOPE_ID" in current and not current.empty
    current_label = "Existing selected sample" if config and (config.use_existing_sample or config.master_file) else "Initial sample derived"
    checks.append((f"Case Universe — {current_label}", current_ready, f"N = {current.SCOPE_ID.nunique():,}" if current_ready else "No initial/current sample could be derived"))
    available_ready = "SCOPE_ID" in available and not available.empty
    target_n = int((config.sample_size if config else None) or profile.default_target_n or 0)
    sufficient_universe = id_ready and (not target_n or universe.SCOPE_ID.nunique() >= target_n)
    checks.append((
        "Case Universe — Available / replacement cases",
        available_ready or (profile.uses_generic_optimizer and sufficient_universe),
        f"N = {available.SCOPE_ID.nunique():,}" if available_ready else "The candidate universe exactly supplies the requested sample; no unselected cases will remain" if profile.uses_generic_optimizer and sufficient_universe else (f"Target N is {target_n:,} but only {universe.SCOPE_ID.nunique():,} candidate case(s) remain after the profile's sample rules; lower Target N or relax the sample rules" if profile.uses_generic_optimizer and target_n else "No available replacement cases were derived"),
    ))
    if current_ready and id_ready and not prepared.universe_deferred:
        missing = set(current.SCOPE_ID) - set(universe.SCOPE_ID)
        checks.append(("Case Universe — Initial membership", not missing, "Every initial/current TEID is in the case universe" if not missing else f"{len(missing)} current TEID(s) are absent from the case universe"))
    else:
        checks.append(("Case Universe — Initial membership", current_ready and id_ready, "Membership will be completed with scored-data candidates during preparation" if prepared.universe_deferred else "Case universe/current sample unavailable"))
    checks.append(("Targets — Canonical specification", not targets.empty, f"{len(targets)} target rule(s) prepared" if not targets.empty else "No target rules could be interpreted"))
    if not targets.empty:
        grouped = targets.groupby(["category", "group", "rule"], dropna=False).target.nunique(dropna=False)
        conflicts = grouped[grouped > 1]
        checks.append(("Targets — Conflicting rules", conflicts.empty, "No conflicting target values" if conflicts.empty else f"{len(conflicts)} category/group/rule combination(s) have conflicting values"))
    supported = {str(value).strip().casefold() for value in (profile.target_families or ["Census", "OSEP", "Clinical"])}
    requested = {
        "demographic": (True if config is None else config.include_census) and bool(supported & {"census", "demographic", "demographics"}),
        "osep": (True if config is None else config.include_osep) and "osep" in supported,
        "clinical": (True if config is None else config.include_clinical) and "clinical" in supported,
    }
    for category in ("demographic", "osep", "clinical"):
        if not targets.empty:
            families = targets.apply(lambda row: _target_family(row.to_dict()), axis=1)
            family = "census" if category == "demographic" else category
            mask = families.eq(family)
        else:
            mask = pd.Series(dtype=bool)
        if requested[category]:
            checks.append((f"Targets — {category.title()} targets", bool(mask.any()) if len(mask) else False, f"{int(mask.sum())} rule(s)" if len(mask) else "None found"))
    return checks
