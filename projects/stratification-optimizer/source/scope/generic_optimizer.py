"""Profile-driven CP-SAT optimization for non-legacy SCOPE assessments."""
from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Callable, TYPE_CHECKING

import pandas as pd
from ortools.sat.python import cp_model

from .canonical import CanonicalInputs, SCOPE_FIELDS, target_criteria

if TYPE_CHECKING:
    from .profiles import TestProfile


OUTPUT_FIELDS = {
    "SCOPE_ID": "teid",
    "SCOPE_FORM": "vform",
    "SCOPE_AGE": "agey",
    "SCOPE_AGE_GROUP": "agegrp",
    "SCOPE_SEX": "sex",
    "SCOPE_ETHNICITY": "eth",
    "SCOPE_EDUCATION": "edl",
    "SCOPE_REGION": "region",
    "SCOPE_LANGUAGE": "language",
    "SCOPE_STUDY_TYPE": "study_type",
}


@dataclass
class GenericOptimization:
    selected: pd.DataFrame
    available: pd.DataFrame
    target_adherence: pd.DataFrame
    metrics: dict[str, object]


def _normalize(value: object, semantic: str) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if semantic == "age_group":
        numbers = [int(value) for value in re.findall(r"\d+", text)]
        if len(numbers) >= 6:
            return f"{numbers[0]}:{numbers[1]}-{numbers[3]}:{numbers[4]}"
        if len(numbers) >= 4:
            return f"{numbers[0]}:{numbers[1]}-{numbers[2]}:{numbers[3]}"
    if re.fullmatch(r"-?\d+(?:\.0+)?", text):
        return str(int(float(text)))
    aliases = {
        "female": "f", "woman": "f", "girl": "f",
        "male": "m", "man": "m", "boy": "m",
    }
    key = re.sub(r"[^a-z0-9]+", "", text.casefold())
    return aliases.get(key, key)


def _target_number(row: dict, target_n: int) -> tuple[float, str]:
    raw = row.get("target", 0)
    numbers = [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", str(raw))]
    if not numbers:
        raise ValueError(f"Target value is not numeric: {raw!r}")
    unit = str(row.get("unit", "N")).strip().casefold()
    if unit in {"%", "percent", "percentage"}:
        pct = numbers[0] / 100 if numbers[0] > 1 else numbers[0]
        return pct * target_n, "percent"
    if str(row.get("rule", "")).strip().casefold() == "range" and len(numbers) >= 2:
        return (numbers[0] + numbers[1]) / 2, "range"
    return numbers[0], "count"


def _profile_level_aliases(profile: TestProfile | None) -> dict[str, dict[str, str]]:
    """Map source codes and display labels to the same profile-owned level."""
    aliases: dict[str, dict[str, str]] = {}
    if profile is None:
        return aliases
    for name, raw in (profile.reporting_variables or {}).items():
        specification = raw if isinstance(raw, dict) else {"canonical": str(raw)}
        semantic = str(specification.get("canonical") or name)
        semantic_aliases = aliases.setdefault(semantic, {})
        for level in specification.get("levels") or []:
            if isinstance(level, dict):
                value = level.get("value", level.get("label", ""))
                label = level.get("label", value)
            else:
                value = label = level
            canonical = _normalize(value, semantic) or _normalize(label, semantic)
            for candidate in (value, label):
                normalized = _normalize(candidate, semantic)
                if normalized:
                    semantic_aliases[normalized] = canonical
    return aliases


def _matching_indices(
    universe: pd.DataFrame,
    criteria: dict[str, str],
    level_aliases: dict[str, dict[str, str]] | None = None,
) -> list[int]:
    mask = pd.Series(True, index=universe.index)
    for semantic, expected in criteria.items():
        column = SCOPE_FIELDS.get(semantic)
        if not column or column not in universe:
            raise ValueError(
                f"The target workbook uses '{semantic.replace('_', ' ')}', but the candidate TEID list has no mapped field for it. "
                "Add that demographic column to the TEID list or map it in the Test Profile."
            )
        aliases = (level_aliases or {}).get(semantic, {})
        normalize = lambda value: aliases.get(_normalize(value, semantic), _normalize(value, semantic))
        normalized = universe[column].map(normalize)
        mask &= normalized.eq(normalize(expected))
    return universe.index[mask].tolist()


def _output_frame(frame: pd.DataFrame) -> pd.DataFrame:
    output = frame.rename(columns=OUTPUT_FIELDS).copy()
    first = [name for name in OUTPUT_FIELDS.values() if name in output]
    remaining = [name for name in output if name not in first]
    return output[first + remaining].reset_index(drop=True)


def _rule_bounds(row: dict, target_n: int) -> tuple[float | None, float | None]:
    """Return (minimum, maximum) acceptable counts for one target row."""
    target_value, _kind = _target_number(row, target_n)
    rule = str(row.get("rule", "Target")).strip().casefold()
    tolerance = row.get("tolerance", "")
    try:
        tolerance_n = float(tolerance) if tolerance not in (None, "") and not pd.isna(tolerance) else 0.0
    except (TypeError, ValueError):
        tolerance_n = 0.0
    if rule == "exact" and math.isclose(target_value, round(target_value)):
        exact = int(round(target_value)); return exact, exact
    if rule in {"minimum", "min"}:
        return math.ceil(target_value - tolerance_n), None
    if rule in {"maximum", "max"}:
        return None, math.floor(target_value + tolerance_n)
    if rule == "range":
        values = [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", str(row.get("target", "")))]
        if len(values) >= 2:
            if str(row.get("unit", "")).strip() == "%":
                values = [(value / 100 if value > 1 else value) * target_n for value in values]
            return math.ceil(min(values)), math.floor(max(values))
    return None, None  # soft Target rows use the objective, not hard bounds


def _infeasible_diagnosis(universe: pd.DataFrame, target_n: int, row_specs: list[tuple[dict, dict[str, str], list[int], float]]) -> str:
    """Explain which hard rules cannot be met against the eligible candidate pool."""
    problems: list[str] = []
    if len(universe) < target_n:
        problems.append(f"only {len(universe):,} eligible candidates exist for a target N of {target_n:,}")
    for row, criteria, indices, _target in row_specs:
        minimum, maximum = _rule_bounds(row, target_n)
        if minimum is None and maximum is None:
            continue
        available = len(indices)
        label = " / ".join(f"{key}={value}" for key, value in criteria.items()) or str(row.get("group", row.get("category", "?")))
        if maximum is not None and maximum < 0:
            problems.append(f"{label}: maximum {maximum} is impossible")
        elif minimum is not None and minimum > available:
            problems.append(f"{label}: needs at least {minimum} matching cases but only {available} eligible candidate(s) match")
        elif minimum is not None and maximum is not None and minimum > maximum:
            problems.append(f"{label}: rule {row.get('rule')} {row.get('target')}{row.get('unit', '')} resolves to an empty range ({minimum}–{maximum})")
        elif available == 0:
            problems.append(f"{label}: no eligible candidates match — check the dataset's level labels against the target workbook")
    detail = "; ".join(problems[:6]) if problems else "the combined exact/minimum/maximum targets conflict; review the tolerances and derived marginals together"
    return f"CP-SAT could not find a feasible sample. {detail}."


def optimize(
    prepared: CanonicalInputs,
    target_n: int,
    max_seconds: int = 300,
    progress: Callable[[str], None] | None = None,
    profile: TestProfile | None = None,
) -> GenericOptimization:
    """Select exactly ``target_n`` candidates with minimum target deviation."""
    universe = prepared.case_universe.reset_index(drop=True).copy()
    if "SCOPE_ID" not in universe or universe.empty:
        raise ValueError("No candidate TEIDs were prepared for optimization.")
    if universe.SCOPE_ID.duplicated().any():
        raise ValueError("Candidate TEIDs must be unique before optimization.")
    if target_n < 1:
        raise ValueError("Target N must be greater than zero.")
    if target_n > len(universe):
        raise ValueError(f"Target N is {target_n:,}, but the candidate TEID list contains only {len(universe):,} unique cases.")

    target_rows = []
    level_aliases = _profile_level_aliases(profile)
    for row in prepared.targets.to_dict("records"):
        criteria = target_criteria(row)
        if criteria:
            target_rows.append((row, criteria, _matching_indices(universe, criteria, level_aliases)))
    if not target_rows:
        raise ValueError("The target workbook does not contain demographic target rows that can be matched to candidate cases.")

    if progress:
        progress(f"Building a generic CP-SAT model for {len(universe):,} candidates and {len(target_rows):,} target cells.")
    model = cp_model.CpModel()
    selected = [model.NewBoolVar(f"selected_{index}") for index in universe.index]
    model.Add(sum(selected) == int(target_n))
    scale = 1000
    deviations = []
    row_specs: list[tuple[dict, dict[str, str], list[int], float]] = []
    relaxed: list[str] = []
    for number, (row, criteria, indices) in enumerate(target_rows):
        target_value, _kind = _target_number(row, target_n)
        rule = str(row.get("rule", "Target")).strip().casefold()
        tolerance = row.get("tolerance", "")
        try:
            tolerance_n = float(tolerance) if tolerance not in (None, "") and not pd.isna(tolerance) else 0.0
        except (TypeError, ValueError):
            tolerance_n = 0.0
        count = sum(selected[index] for index in indices)
        minimum, maximum = _rule_bounds(row, target_n)
        # Hard constraints are only safe when enough matching candidates exist.
        # When a rule is infeasible (for example a region with fewer eligible cases
        # than its rounded minimum), relax it to a soft objective so CP-SAT gets
        # as close to the target as possible instead of failing outright.
        hard_minimum = minimum if minimum is not None and minimum <= len(indices) else None
        hard_maximum = maximum if maximum is not None else None
        if rule == "exact" and math.isclose(target_value, round(target_value)) and len(indices) >= int(round(target_value)):
            model.Add(count == int(round(target_value)))
        elif rule in {"minimum", "min"} and hard_minimum is not None:
            model.Add(count >= int(hard_minimum))
        elif rule in {"maximum", "max"} and hard_maximum is not None:
            model.Add(count <= int(hard_maximum))
        elif rule == "range" and minimum is not None and maximum is not None and minimum <= len(indices):
            model.Add(count >= int(math.ceil(minimum)))
            model.Add(count <= int(math.floor(maximum)))
        else:
            if rule in {"exact", "minimum", "min", "maximum", "max", "range"} and (minimum is not None or maximum is not None):
                relaxed.append(str(row.get("group", row.get("category", "?"))))
            deviation = model.NewIntVar(0, max(target_n * scale, int(abs(target_value) * scale) + target_n * scale), f"deviation_{number}")
            model.AddAbsEquality(deviation, count * scale - int(round(target_value * scale)))
            deviations.append(deviation)
        row_specs.append((row, criteria, indices, target_value))
    if relaxed and progress:
        progress(f"Relaxed infeasible demographic rule(s) to soft targets: {', '.join(relaxed[:6])}{' …' if len(relaxed) > 6 else ''}")

    current_ids = set(prepared.current_sample.get("SCOPE_ID", pd.Series(dtype=str)).astype(str))
    changes = []
    if current_ids:
        for index, teid in enumerate(universe.SCOPE_ID.astype(str)):
            changes.append(1 - selected[index] if teid in current_ids else selected[index])
    target_error = sum(deviations) if deviations else 0
    model.Minimize(target_error * (len(universe) + 1) + sum(changes))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(1, int(max_seconds))
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise ValueError(_infeasible_diagnosis(universe, target_n, row_specs))
    selected_indices = [index for index, variable in enumerate(selected) if solver.Value(variable)]
    selected_frame = universe.iloc[selected_indices].copy()
    available_frame = universe.drop(index=selected_indices).copy()

    adherence_rows = []
    selected_index_set = set(selected_indices)
    total_absolute_difference = 0.0
    for row, criteria, indices, target_value in row_specs:
        actual = sum(index in selected_index_set for index in indices)
        difference = actual - target_value
        total_absolute_difference += abs(difference)
        adherence_rows.append({
            **{name.replace("_", " ").title(): value for name, value in criteria.items()},
            "Target N": target_value,
            "Selected N": actual,
            "Difference N": difference,
            "Target %": target_value / target_n if target_n else 0,
            "Selected %": actual / target_n if target_n else 0,
            "Difference %": difference / target_n if target_n else 0,
            "Available Candidates": len(indices),
            "Source": row.get("source", ""),
            "Note": "Relaxed to soft target" if str(row.get("group", row.get("category", ""))) in relaxed else "",
        })
    status_name = "Optimal" if status == cp_model.OPTIMAL else "Feasible (time limit)"
    metrics = {
        "final_sample_n": len(selected_frame),
        "available_pool_n": len(available_frame),
        "target_cells": len(row_specs),
        "total_absolute_target_difference_n": round(total_absolute_difference, 3),
        "hard_constraints": "Passed" if not relaxed else f"{len(relaxed)} rule(s) relaxed to soft targets",
        "relaxed_targets": relaxed,
        "solver_status": status_name,
        "optimizer": "Generic CP-SAT",
    }
    if progress:
        progress(f"CP-SAT completed with status {status_name}; selected N = {len(selected_frame):,}.")
    return GenericOptimization(
        selected=_output_frame(selected_frame),
        available=_output_frame(available_frame),
        target_adherence=pd.DataFrame(adherence_rows),
        metrics=metrics,
    )
