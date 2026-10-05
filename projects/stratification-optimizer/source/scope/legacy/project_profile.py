"""Small, user-created project settings shared by the GUI and both engines."""
from __future__ import annotations

import json
import os
from pathlib import Path

import pandas as pd


def load_project_profile() -> dict:
    value = os.environ.get("OPT_ENGINE_PROJECT_PROFILE", "").strip()
    if not value:
        return {}
    path = Path(value)
    if not path.is_file():
        raise FileNotFoundError(f"Project setup profile was not found: {path}")
    with path.open("r", encoding="utf-8") as handle:
        profile = json.load(handle)
    if not isinstance(profile, dict):
        raise ValueError("Project setup profile must contain a JSON object.")
    return profile


def matching_column(df: pd.DataFrame, requested: str) -> str | None:
    if requested in df.columns:
        return requested
    lookup = {str(column).lower(): column for column in df.columns}
    return lookup.get(str(requested).lower())


def apply_field_map(df: pd.DataFrame, profile: dict) -> pd.DataFrame:
    """Expose user-selected source labels under the engine's canonical names."""
    mapping = profile.get("field_map") or {}
    if not mapping:
        return df
    out = df.copy()
    for canonical in ("teid", "vform", "agey", "agegrp"):
        source = matching_column(out, str(mapping.get(canonical, "")))
        if source and source != canonical:
            out[canonical] = out[source]
    return out


def apply_scope_filter(df: pd.DataFrame, profile: dict) -> pd.DataFrame:
    """Apply legacy scope plus visual-builder filters to a normalized dataset."""
    rule = profile.get("scope_filter") or {}
    column = matching_column(df, str(rule.get("column", "")))
    values = [str(value).strip().upper() for value in rule.get("values", []) if str(value).strip()]
    out = df
    if column and values:
        out = out.loc[out[column].astype(str).str.strip().str.upper().isin(values)].copy()
    # Version 3 visual profiles may have multiple eligible-sample filters,
    # each selected from detected labels in the representative dataset.
    for requested, allowed in (profile.get("filters") or {}).items():
        column = matching_column(out, str(requested))
        allowed_values = [str(value).strip().upper() for value in allowed if str(value).strip()]
        if column and allowed_values:
            out = out.loc[out[column].astype(str).str.strip().str.upper().isin(allowed_values)].copy()
    return out


def apply_age_bands(df: pd.DataFrame, age: pd.Series | None, profile: dict) -> pd.DataFrame:
    bands = profile.get("age_bands") or []
    if age is None or not bands:
        return df
    numeric_age = pd.to_numeric(age, errors="coerce")
    group = pd.Series(pd.NA, index=df.index, dtype="Int64")
    for band in bands:
        try:
            group_id = int(band["group"])
            if "min_months" in band and "max_months" in band:
                # New GUI profiles use months so half-year groups are exact,
                # even where the source holds fractional ages.
                low = float(band["min_months"])
                high = float(band["max_months"])
                # Dataset ages are commonly stored as decimals (for example
                # 5.99 for 5 years, 11 months), so compare their completed
                # age-in-month rather than requiring an exact integer month.
                values = (numeric_age * 12).floordiv(1)
            else:
                # Preserve support for the first GUI profile format.
                low = float(band["min"])
                high = float(band["max"])
                values = numeric_age
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError("Each custom age band needs group plus min/max years or min_months/max_months.") from exc
        group = group.mask(values.between(low, high), group_id)
    out = df.copy()
    out["agegrp"] = group
    return out


def apply_clinical_groups(df: pd.DataFrame, profile: dict) -> pd.DataFrame:
    groups = profile.get("clinical_groups") or []
    if not groups:
        return df
    out = df.copy()
    for group in groups:
        if not group.get("enabled"):
            continue
        column = matching_column(out, str(group.get("source_column", "")))
        target_flag = str(group.get("target_flag", "")).strip()
        values = [str(value).strip().upper() for value in group.get("values", []) if str(value).strip()]
        if not column or not target_flag or not values:
            continue
        matched = out[column].astype(str).str.strip().str.upper().isin(values)
        existing_source = out[target_flag] if target_flag in out.columns else pd.Series(0, index=out.index)
        existing = pd.to_numeric(existing_source, errors="coerce").fillna(0).astype(int)
        out[target_flag] = (existing.astype(bool) | matched).astype(int)
    return out


def selected_demographics(defaults: list[str], profile: dict) -> list[str]:
    selected = profile.get("demographics")
    if not selected:
        return defaults
    wanted = [str(value).strip() for value in selected if str(value).strip()]
    # Preserve configuration order.  Visual profiles can introduce a new
    # supported demographic field; normalization later quietly skips a field
    # that is absent in a particular source file.
    return wanted or defaults
