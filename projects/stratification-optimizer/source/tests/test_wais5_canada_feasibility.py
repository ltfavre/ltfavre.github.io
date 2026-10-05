"""WAIS-5 Canada real-world infeasibility: filters must not zero the universe,
and CP-SAT must fail with an actionable reason instead of a generic one."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from scope.canonical import prepare_canonical_inputs
from scope.configure import ProjectConfig
from scope.profiles import validate_profile_data, workflow_readiness
from scope import engine as engine_module


ROOT = Path(__file__).resolve().parents[1]
AGE_BANDS = [f"{age}:0-{age}:11" for age in range(16, 29)]


def _profile(language_column: str) -> object:
    """WAIS-5 Canada as installed; language_column mimics the user's mapping."""
    payload = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    payload.update({
        "id": "WAIS5_CA", "display_name": "WAIS-5 Canada", "market": "Canada",
        "profile_kind": "international_adaptation", "status": "Validated",
        "schema_mapping": {
            "teid": "TEID", "age_group": "Target Age Band", "sex": "Sex",
            "ethnicity": "Ethnicity", "education": "PED/SED", "region": "Region",
            "language": language_column, "study_type": "Study",
        },
        "demographics": ["sex", "eth", "edl", "region"],
        "target_families": ["Census"], "languages": ["English"], "study_types": ["STDZ"],
        "capabilities": {"sample_optimization": True}, "source_config": "generic_cp_sat",
        "sampling_specification": {
            "engine_config": "generic_cp_sat",
            "sample_filters": [
                {"field": "study_type", "values": ["STDZ"], "operator": "include"},
                {"field": "language", "values": ["English"], "operator": "include"},
            ],
            "target_tolerances": {
                "Sex": {"tolerance_percent": 2},
                "Region": {"tolerance_percent": 10},
                "Ethnicity": {"tolerance_percent": 5},
            },
            "target_interpretation": {"source_format": "crossed_cells", "cells_as": "report_only", "derive_marginals": ["sex", "region", "ethnicity"]},
        },
    })
    profile, _checks = validate_profile_data(payload)
    assert profile is not None
    return profile


def _project(tmp_path: Path, rows: list[dict]) -> ProjectConfig:
    candidates = tmp_path / "candidates.xlsx"
    pd.DataFrame(rows).to_excel(candidates, sheet_name="Candidates", index=False)
    targets = tmp_path / "targets.xlsx"
    pd.DataFrame([
        {"Target Age Band": "16:0-16:11", "Sex": "F", "Region": "W", "Ethnicity": "WH", "PED/SED": 1, "Count": 1.0, "Pcnt": 0.5},
        {"Target Age Band": "17:0-17:11", "Sex": "M", "Region": "CE", "Ethnicity": "BL", "PED/SED": 3, "Count": 1.0, "Pcnt": 0.5},
    ]).to_excel(targets, sheet_name="Target List", index=False)
    return ProjectConfig(
        profile_id="WAIS5_CA", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", dataset=str(candidates),
        candidate_datasets=[{"path": str(candidates), "source_type": "STDZ", "source_role": "Primary", "priority": 1}],
        study_type="", language="", forms=[], target_file=str(targets), sample_size=2,
        include_census=True, include_osep=False, include_clinical=False, max_solver_seconds=10,
    )


def _rows() -> list[dict]:
    return [
        {"TEID": "1", "Target Age Band": "16:0-16:11", "Sex": "F", "Region": "W", "Ethnicity": "WH", "PED/SED": 1, "Study": "STDZ", "TestLang": "English"},
        {"TEID": "2", "Target Age Band": "16:0-16:11", "Sex": "F", "Region": "W", "Ethnicity": "WH", "PED/SED": 1, "Study": "STDZ", "TestLang": "English"},
        {"TEID": "3", "Target Age Band": "17:0-17:11", "Sex": "M", "Region": "CE", "Ethnicity": "BL", "PED/SED": 3, "Study": "Clinical", "TestLang": "English"},
    ]


def test_profile_language_rule_with_missing_column_does_not_zero_universe(tmp_path):
    # The installed profile mapped language to the literal value "English"; the candidate set has no language column.
    rows = [{key: value for key, value in row.items() if key != "TestLang"} for row in _rows()]
    profile = _profile(language_column="English")
    config = _project(tmp_path, rows)
    prepared = prepare_canonical_inputs(config, profile)
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2"]  # STDZ kept; the broken language rule is skipped with a note
    assert any("Language" in note and "not applied" in note and "English" in note for note in prepared.notes)


def test_profile_language_rule_names_the_misconfigured_column(tmp_path):
    rows = [{**row, "TestLang": None} for row in _rows()]  # language column exists but is empty
    profile = _profile(language_column="TestLang")
    config = _project(tmp_path, rows)
    prepared = prepare_canonical_inputs(config, profile)
    assert any("Language" in note and "TestLang" in note for note in prepared.notes)


def test_profile_rule_that_would_remove_every_case_falls_back_with_a_note(tmp_path):
    # Correct column, but a value that matches nothing.
    profile = _profile(language_column="TestLang")
    profile.sample_filters[1]["values"] = ["Klingon"]
    config = _project(tmp_path, _rows())
    prepared = prepare_canonical_inputs(config, profile)
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2"]
    assert any("would remove every candidate case" in note for note in prepared.notes)


def test_generic_optimization_preflight_names_the_real_blocker(tmp_path, monkeypatch):
    profile = _profile(language_column="TestLang")
    config = _project(tmp_path, _rows())
    config.sample_size = 50  # larger than the eligible pool
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    with pytest.raises(ValueError, match="Target N is 50.*only 2 candidate case"):
        engine_module.run("optimize", config)


def test_solver_diagnosis_names_the_unsatisfiable_rule(tmp_path, monkeypatch):
    """An infeasible demographic rule is relaxed to a soft target, not an error."""
    rows = [
        {"TEID": "1", "Target Age Band": "16:0-16:11", "Sex": "F", "Region": "W", "Ethnicity": "WH", "PED/SED": 1, "Study": "STDZ", "TestLang": "English"},
        {"TEID": "2", "Target Age Band": "16:0-16:11", "Sex": "F", "Region": "W", "Ethnicity": "WH", "PED/SED": 1, "Study": "STDZ", "TestLang": "English"},
        {"TEID": "3", "Target Age Band": "17:0-17:11", "Sex": "M", "Region": "CE", "Ethnicity": "BL", "PED/SED": 3, "Study": "STDZ", "TestLang": "English"},
        {"TEID": "4", "Target Age Band": "17:0-17:11", "Sex": "M", "Region": "CE", "Ethnicity": "BL", "PED/SED": 3, "Study": "STDZ", "TestLang": "English"},
    ]
    profile = _profile(language_column="TestLang")
    # Force an impossible per-group exact N: CE has only 2 candidates but demands 5.
    profile.target_tolerances["Region | CE"] = {"applies_to": "group", "group": "CE", "target_n": 5}
    profile.target_interpretation["derive_marginals"] = ["region"]
    config = _project(tmp_path, rows)
    config.sample_size = 4
    prepared = prepare_canonical_inputs(config, profile)
    config.canonical_targets = [dict(row) for row in prepared.targets.to_dict("records")]
    config.target_review_confirmed = True
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    result = engine_module.run("optimize", config)
    selected = pd.read_excel(result.outputs["optimized_sample"], sheet_name="final_all", dtype={"teid": str})
    assert len(selected) == 4  # solved as close as possible instead of erroring
    assert result.metrics["relaxed_targets"]  # the impossible CE rule was relaxed
    assert "relaxed" in str(result.metrics["hard_constraints"])


def test_generic_profile_with_report_tables_but_no_template_is_flagged(tmp_path):
    """The reporting gap must be visible so a custom template is not silently ignored."""
    profile = _profile(language_column="TestLang")
    payload = {
        "id": "WAIS5_CA", "display_name": "WAIS-5 Canada", "version": "0.1", "status": "Validated",
        "source_config": "generic_cp_sat", "capabilities": {"sample_optimization": True},
        "schema_mapping": profile.schema_mapping,
        "reporting": {"required_tables": {"demographics": ["Sex"], "demo_cross_tabs": [], "demo_three_way": []}},
        "sampling_specification": dict(profile.sampling_specification or {}),
    }
    profile_with_tables, _checks = validate_profile_data(payload)
    assert profile_with_tables is not None
    reporting = workflow_readiness(profile_with_tables)["reporting"]
    assert reporting["status"] == "NEEDS REVIEW"
    assert any("demographic output template" in item for item in reporting["missing"])
