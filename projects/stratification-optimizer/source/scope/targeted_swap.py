"""Profile-driven input preparation for targeted case swaps.

Targeted-swap users may provide either participant-level tables or the
TEID-list workbooks produced by SCOPE reporting.  List-only workbooks are
hydrated from an explicitly selected case universe before validation and
reporting.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

import pandas as pd

from .canonical import FALLBACK_FIELDS, SCOPE_FIELDS
from .profiles import TestProfile


EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
SAMPLE_SHEETS = ("final_all", "selected_cases", "selected_teids", "final_sample_1500", "teid_list")
AVAILABLE_SHEETS = ("available_pool", "available_cases", "usable_pool", "available_swapin", "teid_list")
UNIVERSE_SHEETS = ("case_universe", "master_prepped", "all_usable", "all_cases", "final_all")
REFRESHED_SELECTED_SHEETS = {"selected", "selected_cases", "selected_teids", "final_all"}
REFRESHED_AVAILABLE_SHEETS = {"available", "available_pool", "available_cases", "usable_pool"}


@dataclass
class TargetedSwapInputs:
    sample: pd.DataFrame
    available: pd.DataFrame
    locks: list[str]
    sample_source: str
    available_source: str
    universe_source: str = ""
    sample_data_source: str = ""
    available_data_source: str = ""
    overlap_count: int = 0


def _key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().casefold())


def _clean_ids(values: pd.Series) -> pd.Series:
    return values.astype("string").str.strip().str.replace(r"\.0$", "", regex=True)


def _indexed_columns(frame: pd.DataFrame) -> dict[str, object]:
    return {_key(column): column for column in frame.columns}


def _semantic_for_name(name: str, profile: TestProfile) -> str | None:
    wanted = _key(name)
    for semantic, aliases in FALLBACK_FIELDS.items():
        configured = str((profile.schema_mapping or {}).get(semantic, ""))
        candidates = (semantic, configured, SCOPE_FIELDS.get(semantic, ""), *aliases)
        if wanted in {_key(candidate) for candidate in candidates if candidate}:
            return semantic
    return None


def _configured_field(profile: TestProfile, semantic: str) -> str:
    return str((profile.schema_mapping or {}).get(semantic) or FALLBACK_FIELDS[semantic][0])


def default_lock_fields(profile: TestProfile) -> list[str]:
    declared = (profile.sampling_specification or {}).get("default_locks") or []
    if declared:
        locks = []
        for value in declared:
            semantic = _semantic_for_name(str(value), profile)
            locks.append(_configured_field(profile, semantic) if semantic else str(value))
        return locks
    return [_configured_field(profile, semantic) for semantic in ("form", "age_group", "sex")]


def normalize_lock_fields(profile: TestProfile, requested: list[str] | None) -> list[str]:
    values = requested if requested is not None else default_lock_fields(profile)
    normalized: list[str] = []
    for value in values:
        semantic = _semantic_for_name(str(value), profile)
        field = _configured_field(profile, semantic) if semantic else str(value)
        if field not in normalized:
            normalized.append(field)
    return normalized


def stage_swap_selection(
    add_teid: object,
    remove_teid: object,
    reason: object = "",
    existing: list[dict] | None = None,
) -> dict | None:
    """Normalize one UI swap selection and reject incomplete or duplicate pairs.

    Returning ``None`` for an entirely blank entry lets the GUI safely call this
    when users continue without having started another pair.
    """
    add = str(add_teid or "").strip()
    remove = str(remove_teid or "").strip()
    note = str(reason or "").strip()
    if not add and not remove and not note:
        return None
    if not add or not remove:
        raise ValueError("Enter both the Add TEID and Remove TEID for this swap selection.")
    for row in existing or []:
        if str(row.get("remove", "")).strip() == remove or str(row.get("add", "")).strip() == add:
            raise ValueError("Each Remove TEID and Add TEID may appear only once.")
    return {"remove": remove, "add": add, "reason": note}


def _column_score(columns: list[object], role: str, profile: TestProfile) -> int:
    keys = {_key(column) for column in columns}
    exact_ids = {_key("teid"), _key("scope_id"), _key((profile.schema_mapping or {}).get("teid", ""))}
    score = 100 if keys & exact_ids else 0
    if role == "sample" and any(key.startswith("selectedteids") or key.startswith("selectedcases") for key in keys):
        score = max(score, 80)
    if role == "available" and any("available" in key and ("teid" in key or "pool" in key) for key in keys):
        score = max(score, 80)
    if role == "universe" and not keys & exact_ids:
        return 0
    for semantic in ("form", "age", "age_group", "sex", "ethnicity", "education", "region"):
        candidates = ((profile.schema_mapping or {}).get(semantic, ""), *FALLBACK_FIELDS[semantic])
        if any(_key(value) in keys for value in candidates if value):
            score += 5
    return score


def _read_excel_role(
    path: Path,
    role: str,
    profile: TestProfile,
    combine_refreshed_membership: bool = False,
) -> tuple[pd.DataFrame, str]:
    preferred = {_key(name): index for index, name in enumerate(
        SAMPLE_SHEETS if role == "sample" else AVAILABLE_SHEETS if role == "available" else UNIVERSE_SHEETS
    )}
    candidates: list[tuple[int, int, int, str]] = []
    engine = "openpyxl" if path.suffix.casefold() in {".xlsx", ".xlsm"} else "xlrd"
    with pd.ExcelFile(path, engine=engine) as book:
        for sheet_index, sheet in enumerate(book.sheet_names):
            for header in range(6):
                try:
                    columns = list(pd.read_excel(book, sheet_name=sheet, header=header, nrows=0).columns)
                except Exception:
                    continue
                score = _column_score(columns, role, profile)
                if score:
                    preferred_bonus = max(0, 10 - preferred.get(_key(sheet), 10))
                    candidates.append((score + preferred_bonus, -header, -sheet_index, sheet))
        if not candidates:
            description = "participant-level data or a recognizable TEID list" if role != "universe" else "participant-level data with a TEID column"
            raise ValueError(f"{path.name} does not contain {description} in the first six rows of any worksheet.")
        if role == "available" and combine_refreshed_membership:
            selected_candidates = [item for item in candidates if _key(item[3]) in {_key(name) for name in REFRESHED_SELECTED_SHEETS}]
            available_candidates = [item for item in candidates if _key(item[3]) in {_key(name) for name in REFRESHED_AVAILABLE_SHEETS}]
            if selected_candidates and available_candidates:
                selected_choice = max(selected_candidates)
                available_choice = max(available_candidates)
                selected_frame = pd.read_excel(book, sheet_name=selected_choice[3], header=-selected_choice[1])
                available_frame = pd.read_excel(book, sheet_name=available_choice[3], header=-available_choice[1])
                frame = pd.concat([selected_frame, available_frame], ignore_index=True, sort=False)
                return frame, f"{path.name} [{selected_choice[3]} + {available_choice[3]}]"
        _score, negative_header, _negative_sheet, sheet = max(candidates)
        header = -negative_header
        frame = pd.read_excel(book, sheet_name=sheet, header=header)
    return frame, f"{path.name} [{sheet}]"


def _read_role(
    path: str | Path,
    role: str,
    profile: TestProfile,
    combine_refreshed_membership: bool = False,
) -> tuple[pd.DataFrame, str]:
    source = Path(path)
    if not source.is_file():
        raise ValueError(f"The {role.replace('_', ' ')} file was not found: {source}")
    suffix = source.suffix.casefold()
    if suffix in EXCEL_SUFFIXES:
        frame, label = _read_excel_role(source, role, profile, combine_refreshed_membership)
    elif suffix == ".csv":
        frame, label = pd.read_csv(source), source.name
    elif suffix in {".sas7bdat", ".xpt"}:
        frame = pd.read_sas(source, format="xport" if suffix == ".xpt" else "sas7bdat", encoding="latin1")
        label = source.name
    else:
        raise ValueError(f"Unsupported targeted-swap input type '{suffix or '(none)'}': {source.name}")
    return frame, label


def _role_id_column(frame: pd.DataFrame, role: str, profile: TestProfile) -> object | None:
    indexed = _indexed_columns(frame)
    configured = (profile.schema_mapping or {}).get("teid", "")
    for candidate in (configured, *FALLBACK_FIELDS["teid"], SCOPE_FIELDS["teid"]):
        if candidate and _key(candidate) in indexed:
            return indexed[_key(candidate)]
    if role == "sample":
        return next((column for column in frame.columns if _key(column).startswith("selectedteids") or _key(column).startswith("selectedcases")), None)
    if role == "available":
        return next((column for column in frame.columns if "available" in _key(column) and ("teid" in _key(column) or "pool" in _key(column))), None)
    return None


def _derive_age_groups(frame: pd.DataFrame, profile: TestProfile) -> None:
    target = _configured_field(profile, "age_group")
    if target in frame and frame[target].notna().any():
        return
    indexed = _indexed_columns(frame)
    age_column = next((indexed[_key(name)] for name in ((profile.schema_mapping or {}).get("age", ""), *FALLBACK_FIELDS["age"]) if name and _key(name) in indexed), None)
    month_column = next((column for key, column in indexed.items() if key in {"agem", "agemonth", "agemonths"}), None)
    if age_column is None:
        return
    years = pd.to_numeric(frame[age_column], errors="coerce")
    months = pd.to_numeric(frame[month_column], errors="coerce").fillna(0) if month_column is not None else 0
    total_months = years * 12 + months
    derived = pd.Series(pd.NA, index=frame.index, dtype="string")
    for label in profile.age_structure or []:
        match = re.fullmatch(r"\s*(\d+)\s*:\s*(\d+)\s*[-\u2013]\s*(\d+)\s*:\s*(\d+)\s*", str(label))
        if not match:
            continue
        y1, m1, y2, m2 = map(int, match.groups())
        derived.loc[total_months.between(y1 * 12 + m1, y2 * 12 + m2)] = str(label).replace("\u2013", "-")
    if derived.notna().any():
        frame[target] = derived


def _derive_forms(frame: pd.DataFrame, profile: TestProfile) -> None:
    target = _configured_field(profile, "form")
    if target in frame and frame[target].notna().any():
        return
    rules = ((profile.sampling_specification or {}).get("targeted_swap_derivations") or {}).get("form_age_ranges") or []
    if not rules:
        return
    indexed = _indexed_columns(frame)
    age_column = next((indexed[_key(name)] for name in ((profile.schema_mapping or {}).get("age", ""), *FALLBACK_FIELDS["age"]) if name and _key(name) in indexed), None)
    if age_column is None:
        return
    age = pd.to_numeric(frame[age_column], errors="coerce")
    derived = pd.Series(pd.NA, index=frame.index, dtype="string")
    for rule in rules:
        try:
            minimum, maximum = float(rule["min_age"]), float(rule["max_age"])
        except (KeyError, TypeError, ValueError):
            continue
        derived.loc[age.between(minimum, maximum)] = str(rule.get("value", ""))
    if derived.notna().any():
        frame[target] = derived


def _materialize_target_flags(frame: pd.DataFrame, profile: TestProfile) -> None:
    for target in profile.target_flags:
        if target in frame:
            continue
        pattern = re.compile(rf"^{re.escape(target)}(?:_RF\d+)?$", re.IGNORECASE)
        columns = [column for column in frame.columns if pattern.fullmatch(str(column))]
        if columns:
            frame[target] = pd.concat([pd.to_numeric(frame[column], errors="coerce").eq(1) for column in columns], axis=1).any(axis=1).astype(int)


def normalize_case_frame(frame: pd.DataFrame, role: str, profile: TestProfile) -> pd.DataFrame:
    source = frame.copy()
    id_column = _role_id_column(source, role, profile)
    if id_column is None:
        raise ValueError(f"The {role.replace('_', ' ')} input does not contain a recognizable TEID column.")
    exact_id_keys = {_key(value) for value in ((profile.schema_mapping or {}).get("teid", ""), *FALLBACK_FIELDS["teid"], SCOPE_FIELDS["teid"])}
    if _key(id_column) not in exact_id_keys:
        source = pd.DataFrame({"teid": _clean_ids(source[id_column])})
    else:
        source["teid"] = _clean_ids(source[id_column])
    source = source.loc[source.teid.notna() & source.teid.ne("")].copy()
    indexed = _indexed_columns(source)
    for semantic in FALLBACK_FIELDS:
        if semantic == "teid":
            continue
        target = _configured_field(profile, semantic)
        candidates = ((profile.schema_mapping or {}).get(semantic, ""), *FALLBACK_FIELDS[semantic], SCOPE_FIELDS.get(semantic, ""))
        raw = next((indexed[_key(name)] for name in candidates if name and _key(name) in indexed), None)
        if raw is not None and target not in source:
            source[target] = source[raw]
    _derive_age_groups(source, profile)
    _derive_forms(source, profile)
    _materialize_target_flags(source, profile)
    return source.reset_index(drop=True)


def _hydrate(
    frame: pd.DataFrame,
    universe: pd.DataFrame,
    label: str,
    *,
    alternate_universe: pd.DataFrame | None = None,
) -> pd.DataFrame:
    if universe.teid.duplicated().any():
        duplicates = universe.loc[universe.teid.duplicated(keep=False), "teid"].drop_duplicates().head(10).tolist()
        raise ValueError("The targeted-swap case universe contains duplicate TEIDs: " + ", ".join(duplicates))
    missing = sorted(set(frame.teid) - set(universe.teid))
    if missing:
        preview = ", ".join(missing[:10]) + ("..." if len(missing) > 10 else "")
        if alternate_universe is not None and set(missing).issubset(set(alternate_universe.teid)):
            raise ValueError(
                f"{len(missing)} TEID(s) from the {label} were not found in its matching participant dataset "
                f"but were found in the other participant dataset: {preview}. "
                "The membership sources appear to be assigned to the wrong stage. "
                "Current selected TEID list must be the pre-swap list that matches the Original dataset; "
                "the add-candidate/refreshed source must match the Updated dataset."
            )
        raise ValueError(f"{len(missing)} TEID(s) from the {label} were not found in the case universe: {preview}")
    lookup = universe.set_index("teid", drop=False)
    result = frame.set_index("teid", drop=False).copy()
    aligned = lookup.loc[result.index]
    missing_columns = [column for column in aligned.columns if column != "teid" and column not in result]
    if missing_columns:
        result = pd.concat([result, aligned[missing_columns]], axis=1)
    for column in (column for column in aligned.columns if column != "teid" and column in frame.columns):
        blank = result[column].isna() | result[column].astype("string").str.strip().eq("")
        result.loc[blank, column] = aligned.loc[blank, column]
    return result.reset_index(drop=True)


def load_targeted_swap_inputs(
    sample_file: str | Path,
    available_file: str | Path,
    profile: TestProfile,
    required_locks: list[str] | None = None,
    case_universe_file: str | Path = "",
    sample_data_file: str | Path = "",
    available_data_file: str | Path = "",
) -> TargetedSwapInputs:
    sample_path = Path(sample_file).resolve()
    available_path = Path(available_file).resolve()
    raw_sample, sample_source = _read_role(sample_file, "sample", profile)
    raw_available, available_source = _read_role(
        available_file, "available", profile,
        combine_refreshed_membership=sample_path != available_path,
    )
    sample = normalize_case_frame(raw_sample, "sample", profile)
    available = normalize_case_frame(raw_available, "available", profile)
    locks = normalize_lock_fields(profile, required_locks)

    universe_source = ""
    sample_data_source = ""
    available_data_source = ""
    sample_data: pd.DataFrame | None = None
    available_data: pd.DataFrame | None = None
    if str(sample_data_file).strip():
        raw_sample_data, sample_data_source = _read_role(sample_data_file, "universe", profile)
        sample_data = normalize_case_frame(raw_sample_data, "universe", profile)
    if str(available_data_file).strip():
        raw_available_data, available_data_source = _read_role(available_data_file, "universe", profile)
        available_data = normalize_case_frame(raw_available_data, "universe", profile)
    if sample_data is not None:
        sample = _hydrate(
            sample, sample_data, "current selected TEID list",
            alternate_universe=available_data,
        )
    if available_data is not None:
        available = _hydrate(
            available, available_data, "available/candidate TEID list",
            alternate_universe=sample_data,
        )
    if str(case_universe_file).strip() and (not sample_data_source or not available_data_source):
        raw_universe, universe_source = _read_role(case_universe_file, "universe", profile)
        universe = normalize_case_frame(raw_universe, "universe", profile)
        if not sample_data_source:
            sample = _hydrate(sample, universe, "current selected TEID list")
        if not available_data_source:
            available = _hydrate(available, universe, "available/candidate TEID list")

    missing_fields = [field for field in locks if field not in sample.columns or field not in available.columns]
    if missing_fields:
        detail = ", ".join(missing_fields)
        raise ValueError(
            "The selected swap workbooks contain TEID lists but not all participant-level lock fields "
            f"({detail}). Select a complete Case universe on Project so SCOPE can hydrate those TEIDs."
        )
    if sample.teid.duplicated().any():
        raise ValueError("Targeted swaps require unique TEIDs in the current sample.")
    if available.teid.duplicated().any():
        raise ValueError("Targeted swaps require unique TEIDs in the available pool.")
    # A refreshed candidate dataset commonly contains the currently selected
    # participants as well as new candidates. Current membership is authoritative:
    # overlapping TEIDs remain selected and are removed from the effective add pool.
    # This is not a data error and must not block entry of an otherwise valid swap.
    selected_ids = set(sample.teid)
    overlap_count = int(available.teid.isin(selected_ids).sum())
    if overlap_count:
        available = available.loc[~available.teid.isin(selected_ids)].copy().reset_index(drop=True)
    return TargetedSwapInputs(
        sample, available, locks, sample_source, available_source, universe_source,
        sample_data_source, available_data_source, overlap_count,
    )


def comparable_lock_value(value: object, field: str, profile: TestProfile) -> str:
    text = str(value).strip().casefold().replace("\u2013", "-")
    semantic = _semantic_for_name(field, profile)
    if semantic == "form":
        match = re.search(r"(?:record\s*)?form[_\s-]*(\d+)|\brf[_\s-]*(\d+)", text)
        if match:
            return next(group for group in match.groups() if group is not None)
    # SAS and Excel often represent the same categorical code differently
    # (for example 3.0 versus 3). Compare the value of a numeric code instead
    # of its source-specific display formatting.
    try:
        number = Decimal(text)
        if number.is_finite():
            return format(number.normalize(), "f")
    except InvalidOperation:
        pass
    return re.sub(r"\s+", "", text)
