"""Profile and QA summaries that never alter sample membership."""
from __future__ import annotations

import re
from typing import Iterable

import pandas as pd


DISPLAY_FIELDS = {
    "age": ("agegrp", "age_group", "agey", "age"),
    "age band": ("agegrp", "age_group", "agey", "age"),
    "sex": ("sex", "gender"),
    "ethnicity": ("eth", "ethnicity"),
    "raceeth": ("eth", "raceeth", "ethnicity"),
    "education": ("edl", "ped", "education"),
    "parent education": ("ped", "edl", "education"),
    "ped": ("ped", "edl", "education"),
    "region": ("region",),
    "language": ("testlang", "language"),
    "form": ("vform", "form"),
}


def _key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).casefold())


def _field(sample: pd.DataFrame, display: str) -> str | None:
    indexed = {_key(column): str(column) for column in sample.columns}
    for candidate in DISPLAY_FIELDS.get(display.strip().casefold(), (display,)):
        if _key(candidate) in indexed:
            return indexed[_key(candidate)]
    return None


def _target_lookup(targets: Iterable[dict] | None) -> dict[tuple[str, str], float]:
    lookup: dict[tuple[str, str], float] = {}
    for row in targets or []:
        category = row.get("category", row.get("Category", ""))
        group = row.get("group", row.get("Group", ""))
        unit = str(row.get("unit", row.get("Unit", ""))).strip().casefold()
        value = row.get("target", row.get("Target"))
        try:
            number = float(value)
        except (TypeError, ValueError):
            continue
        if unit in {"%", "percent", "percentage"}:
            normalized = number / 100 if number > 1 else number
            lookup[(_key(category), _key(group))] = normalized
            lookup.setdefault(("", _key(group)), normalized)
    return lookup


def _count_table(sample: pd.DataFrame, fields: list[str], title: str, targets: dict[tuple[str, str], float]) -> pd.DataFrame:
    total = max(1, len(sample))
    table = sample.groupby(fields, dropna=False).size().reset_index(name="Selected N")
    table["Selected %"] = table["Selected N"] / total
    target_values = []
    for row in table.itertuples(index=False, name=None):
        labels = row[:len(fields)]
        joined = " | ".join(str(value) for value in labels)
        target = targets.get((_key(title), _key(joined)))
        if target is None and len(fields) == 1:
            target = targets.get((_key(fields[0]), _key(labels[0])))
            if target is None:
                target = targets.get((_key(title), _key(labels[0])))
        target_values.append(target)
    table["Target %"] = target_values
    table["Diff"] = table["Selected %"] - table["Target %"]
    return table


def _sheet_key(prefix: str, label: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9]+", "_", label).strip("_")
    return f"{prefix}_{safe}"[:31]


def _target_comparison_tables(sample: pd.DataFrame, targets: list[dict] | None) -> dict[str, pd.DataFrame]:
    """Materialize canonical target rows into explicit Actual/Target/Diff tables."""
    grouped: dict[str, list[dict]] = {}
    for target in targets or []:
        category = str(target.get("category", target.get("Category", ""))).strip()
        group = str(target.get("group", target.get("Group", ""))).strip()
        unit = str(target.get("unit", target.get("Unit", ""))).strip()
        raw_target = target.get("target", target.get("Target"))
        numbers = re.findall(r"-?\d+(?:\.\d+)?", str(raw_target))
        if not category or not group or not numbers:
            continue
        target_number = sum(float(value) for value in numbers) / len(numbers)
        if unit == "%" and target_number > 1:
            target_number /= 100
        subset = sample
        pieces = [piece.strip() for piece in group.split("/")]
        form = pieces[0] if len(pieces) >= 2 and _field(sample, "Form") else ""
        age_match = next((re.search(r"age\s*(?:group|band)?\s*(.+)$", piece, re.I) for piece in pieces if re.search(r"age\s*(?:group|band)?\s*(.+)$", piece, re.I)), None)
        age_value = age_match.group(1).strip() if age_match else ""
        form_field, age_field = _field(sample, "Form"), _field(sample, "Age")
        if form and form_field:
            subset = subset.loc[subset[form_field].astype(str).str.strip().eq(form)]
        if age_value and age_field:
            subset = subset.loc[subset[age_field].astype(str).str.strip().eq(age_value)]
        category_field = _field(sample, category)
        selected_n: int
        denominator = len(subset)
        group_value = pieces[-1] if len(pieces) >= 3 else group
        if category.casefold() == "sample cell":
            selected_n = denominator
        elif group in sample.columns:
            selected_n = int(pd.to_numeric(subset[group], errors="coerce").fillna(0).astype(bool).sum())
        elif category_field:
            values = subset[category_field].astype(str).str.strip()
            candidates = [group_value, re.sub(r"^\d+", "", group_value)]
            match_value = next((candidate for candidate in candidates if candidate and values.eq(candidate).any()), candidates[-1])
            selected_n = int(values.eq(match_value).sum())
            group_value = match_value
        else:
            continue
        actual_pct = selected_n / denominator if denominator else 0.0
        row = {"Form": form, "Age Group": age_value, "Group": group_value, "Selected N": selected_n}
        if unit == "%":
            row.update({"Target %": target_number, "Selected %": actual_pct, "Diff": actual_pct - target_number})
        else:
            row.update({"Target N": target_number, "Diff N": selected_n - target_number})
        grouped.setdefault(category, []).append(row)
    return {_sheet_key("Target", category): pd.DataFrame(rows) for category, rows in grouped.items() if rows}


def profile_sample(
    sample: pd.DataFrame,
    demographics: list[str],
    flags: list[str],
    required_demo_tables: list[str] | None = None,
    targets: list[dict] | None = None,
) -> dict[str, pd.DataFrame]:
    """Build profile-owned marginal and cross-tab reports.

    Percentage comparisons use ``Diff = Actual - Target``. If no percentage
    target is available, Target and Diff remain blank.
    """
    outputs: dict[str, pd.DataFrame] = {}
    target_lookup = _target_lookup(targets)
    if {"vform", "agegrp"}.issubset(sample.columns):
        outputs["form_age_counts"] = _count_table(sample, ["vform", "agegrp"], "Form × Age", target_lookup)
    fields_seen: set[str] = set()
    for display in ["Sex", *demographics]:
        field = _field(sample, display)
        if not field or field in fields_seen:
            continue
        fields_seen.add(field)
        outputs[_sheet_key("Demo", display)] = _count_table(sample, [field], display, target_lookup)
    for requested in required_demo_tables or []:
        parts = [part.strip() for part in re.split(r"\s*(?:×|\b[xX]\b)\s*", requested) if part.strip()]
        if len(parts) != 2:
            continue
        fields = [_field(sample, part) for part in parts]
        if any(field is None for field in fields):
            continue
        outputs[_sheet_key("Cross", requested)] = _count_table(sample, [str(field) for field in fields], requested, target_lookup)
    outputs.update(_target_comparison_tables(sample, targets))
    total = max(1, len(sample))
    for flag in flags:
        if flag not in sample.columns:
            continue
        count = int(pd.to_numeric(sample[flag], errors="coerce").fillna(0).astype(bool).sum())
        actual = count / total
        family = "Clinical" if "clin" in flag.casefold() or flag.casefold().startswith("c_") else "OSEP"
        target = target_lookup.get((_key(family), _key(flag)), target_lookup.get(("", _key(flag))))
        outputs[_sheet_key(family, flag)] = pd.DataFrame({
            "Group": [flag], "Selected N": [count], "Selected %": [actual],
            "Target %": [target], "Diff": [actual - target if target is not None else None],
        })
    return outputs
