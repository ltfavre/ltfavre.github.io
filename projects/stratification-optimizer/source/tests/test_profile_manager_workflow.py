"""Progressive-disclosure profile authoring: international adaptations, sample
filters, target interpretation, and optimization tolerances (WAIS-5 Canada)."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

import scope.profiles as profiles_module
from gui.profile_manager import build_profile_draft, normalized_workflow_inputs
from scope.canonical import build_target_specification, canonical_target_records, prepare_canonical_inputs
from scope.configure import ProjectConfig
from scope.profiles import (
    derive_marginal_target_rows, profile_sample_filters, profile_optimization_rules,
    profile_validation_summary, validate_profile_data, workflow_readiness,
)


ROOT = Path(__file__).resolve().parents[1]


def _international_wais_ca() -> dict:
    """WAIS-5 Canada configured from a dataset, a crossed census workbook, the
    WAIS-5 US reference profile, and manual RD decisions — no Research Plan."""
    parent = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    draft = build_profile_draft("WAIS-5", "5", "0.1", "Canada", "", "international", "WAIS5", parent, "WAIS5.json")
    draft.update({
        "schema_mapping": {
            "teid": "TEID", "age_group": "Target Age Band", "sex": "Sex",
            "region": "Region", "ethnicity": "Ethnicity", "education": "PED/SED", "study_type": "Study",
        },
        "demographics": ["sex", "eth", "edl", "region"],
        "target_families": ["Census"],
        "capabilities": {"sample_optimization": True},
        "source_config": "generic_cp_sat",
        "workflow_inputs": normalized_workflow_inputs({}, {"sample_optimization": True}),
    })
    draft["sampling_specification"] = {
        "engine_config": "generic_cp_sat",
        "sample_filters": [
            {"field": "study_type", "label": "Study", "operator": "include", "values": ["STDZ"], "origin": "manual", "note": "RD decision: standardization cases only"},
        ],
        "target_tolerances": {
            "Region": {"tolerance_percent": 10, "priority": "Demographic", "origin": "manual", "note": "RD: Region ±10%"},
            "Sex": {"tolerance_percent": 5, "priority": "Demographic", "origin": "inherited"},
        },
        "target_interpretation": {
            "source_format": "crossed_cells",
            "cells_as": "report_only",
            "derive_marginals": ["region", "sex", "ethnicity", "education"],
            "notes": "WAIS 5 EN CDN Targets.xlsx is fully crossed; optimize derived marginals.",
        },
        "target_sources": [
            {"path": "WAIS 5 EN CDN Targets.xlsx", "format": "crossed_cells", "family": "Census", "notes": "Fully crossed census cells", "origin": "assessment"},
        ],
    }
    return draft


def test_international_draft_marks_inherited_origins_and_needs_no_research_plan():
    draft = _international_wais_ca()
    assert draft["profile_kind"] == "international_adaptation"
    assert draft["id"] == "WAIS5_CA"
    assert draft["status"] == "Needs Review"
    # The WAIS-5 US parent supplies demographics, so that section is inherited.
    assert draft["provenance"]["field_origins"]["demographics"] == "inherited"
    profile, _checks = validate_profile_data(draft)
    assert profile is not None
    summary = profile_validation_summary(profile)
    assert summary["core_valid"]
    # Sources are optional evidence, not missing setup.
    sources = next(item for item in summary["sections"] if item["key"] == "sources")
    assert sources["status"] in {"NOT CONFIGURED", "COMPLETE"}
    # Inherited values are flagged for confirmation until the author reviews them.
    pending = workflow_readiness(profile)["sample_optimization"]
    assert pending["status"] == "NEEDS REVIEW"
    assert any("confirmation of inherited" in item for item in pending["missing"])
    # Confirming the inherited sections (the Profile Manager's Confirm action) clears review.
    draft["adaptation"]["review"] = {key: "confirmed" for key in draft["adaptation"]["review"]}
    confirmed, _checks = validate_profile_data(draft)
    readiness = workflow_readiness(confirmed)
    assert readiness["sample_optimization"]["status"] == "READY", readiness["sample_optimization"]


def test_international_draft_without_parent_is_a_valid_skeleton():
    draft = build_profile_draft("New Assessment", "1", "0.1", "Australia", "", "international")
    assert draft["profile_kind"] == "international_adaptation"
    profile, _checks = validate_profile_data(draft)
    assert profile is not None
    assert profile_validation_summary(profile)["core_valid"]
    assert all(item["status"] == "NOT CONFIGURED" for item in profile_validation_summary(profile)["workflows"].values())


def test_sample_filters_satisfy_the_eligibility_section():
    draft = _international_wais_ca()
    draft["adaptation"]["review"] = {key: "confirmed" for key in draft["adaptation"]["review"]}
    profile, _checks = validate_profile_data(draft)
    summary = profile_validation_summary(profile)
    eligibility = next(item for item in summary["sections"] if item["key"] == "eligibility")
    assert eligibility["status"] == "COMPLETE"
    filters = profile_sample_filters(profile)
    assert filters[0]["field"] == "study_type"
    assert filters[0]["values"] == ["STDZ"]


def test_target_tolerances_complete_the_targets_section():
    draft = _international_wais_ca()
    draft["adaptation"]["review"] = {key: "confirmed" for key in draft["adaptation"]["review"]}
    profile, _checks = validate_profile_data(draft)
    targets = next(item for item in profile_validation_summary(profile)["sections"] if item["key"] == "targets")
    assert targets["status"] == "COMPLETE"
    rules = {rule["semantic"]: rule for rule in profile_optimization_rules(profile)}
    assert rules["region"]["tolerance_percent"] == 10
    assert rules["sex"]["origin"] == "inherited"


def test_derive_marginal_target_rows_from_crossed_cells():
    crossed = [
        {"criteria": {"age_group": "16:0-17:11", "sex": "F", "region": "W"}, "target": 25.0, "unit": "%"},
        {"criteria": {"age_group": "16:0-17:11", "sex": "M", "region": "W"}, "target": 25.0, "unit": "%"},
        {"criteria": {"age_group": "18:0-19:11", "sex": "F", "region": "CE"}, "target": 30.0, "unit": "%"},
        {"criteria": {"age_group": "18:0-19:11", "sex": "M", "region": "CE"}, "target": 20.0, "unit": "%"},
    ]
    derived = derive_marginal_target_rows(crossed, ["region", "sex"], {"Region": {"tolerance_percent": 10}}, source_label="Derived from WAIS 5 EN CDN Targets.xlsx")
    by_category = {}
    for row in derived:
        by_category.setdefault(row["category"], {})[row["group"]] = row
    assert by_category["Region"]["W"]["rule"] == "Range"
    assert by_category["Region"]["W"]["target"] == "40–60"  # 50% ± 10%
    assert by_category["Region"]["CE"]["target"] == "40–60"
    assert by_category["Sex"]["F"]["rule"] == "Target"
    assert float(by_category["Sex"]["F"]["target"]) == pytest.approx(55.0)
    assert all(json.loads(row["criteria"]) for row in derived)


def test_derive_marginal_target_rows_honor_exact_target_n_all_and_per_group():
    crossed = [
        {"criteria": {"age_group": "16:0-17:11", "sex": "F"}, "target": 50.0, "unit": "%"},
        {"criteria": {"age_group": "18:0-19:11", "sex": "F"}, "target": 50.0, "unit": "%"},
    ]
    derived = derive_marginal_target_rows(crossed, ["age_group"], {
        "Age Group": {"target_n": 100, "priority": "Required"},
        "Age Group | 16:0-17:11": {"applies_to": "group", "group": "16:0-17:11", "target_n": 120},
    })
    by_group = {row["group"]: row for row in derived}
    assert by_group["16:0-17:11"]["rule"] == "Exact" and by_group["16:0-17:11"]["target"] == "120" and by_group["16:0-17:11"]["unit"] == "N"
    assert by_group["18:0-19:11"]["rule"] == "Exact" and by_group["18:0-19:11"]["target"] == "100"


def test_build_target_specification_applies_exact_n_to_marginals(tmp_path):
    draft = _international_wais_ca()
    draft["sampling_specification"]["target_tolerances"] = {
        "Region": {"target_n": 1, "priority": "Required"},
    }
    profile, _checks = validate_profile_data(draft)
    workbook = tmp_path / "WAIS 5 EN CDN Targets.xlsx"
    pd.DataFrame([
        {"Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1, "Count": 1.0, "Pcnt": 0.5},
        {"Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4, "Count": 1.0, "Pcnt": 0.5},
    ]).to_excel(workbook, sheet_name="Target List", index=False)
    config = ProjectConfig(
        profile_id="WAIS5_CA", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", target_file=str(workbook), sample_size=2, study_type="",
    )
    targets = build_target_specification(config, profile)
    region = targets[targets["category"] == "Region"]
    assert region["rule"].tolist() == ["Exact", "Exact"]
    assert region["target"].astype(str).tolist() == ["1", "1"]


def test_build_target_specification_derives_marginals_and_drops_cells(tmp_path):
    profile, _checks = validate_profile_data(_international_wais_ca())
    workbook = tmp_path / "WAIS 5 EN CDN Targets.xlsx"
    pd.DataFrame([
        {"Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1, "Count": 1.0, "Pcnt": 0.5},
        {"Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4, "Count": 1.0, "Pcnt": 0.5},
    ]).to_excel(workbook, sheet_name="Target List", index=False)
    config = ProjectConfig(
        profile_id="WAIS5_CA", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", target_file=str(workbook), sample_size=2, study_type="",
    )
    targets = build_target_specification(config, profile)
    categories = targets["category"].tolist()
    # Raw crossed cells are report-only: they must not become CP-SAT objective rows.
    assert not any("×" in str(category) for category in categories)
    assert "Region" in categories and "Sex" in categories and "Ethnicity" in categories and "Education" in categories
    region = targets[targets["category"] == "Region"].iloc[0]
    assert region["rule"] == "Range"
    low, high = (float(value) for value in str(region["target"]).replace("%", "").split("–"))
    assert low == pytest.approx(40.0) and high == pytest.approx(60.0)
    sample = targets[targets["category"] == "Sample"].iloc[0]
    assert sample["rule"] == "Exact" and int(sample["target"]) == 2


def test_cells_remain_objective_rows_when_profile_keeps_them(tmp_path):
    draft = _international_wais_ca()
    draft["sampling_specification"]["target_interpretation"]["cells_as"] = "objective"
    profile, _checks = validate_profile_data(draft)
    workbook = tmp_path / "targets.xlsx"
    pd.DataFrame([
        {"Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1, "Pcnt": 0.5},
        {"Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4, "Pcnt": 0.5},
    ]).to_excel(workbook, sheet_name="Target List", index=False)
    config = ProjectConfig(
        profile_id="WAIS5_CA", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", target_file=str(workbook), sample_size=2, study_type="",
    )
    targets = build_target_specification(config, profile)
    categories = [str(value) for value in targets["category"].tolist()]
    assert any("×" in category for category in categories)  # crossed cells kept
    assert "Region" in categories  # marginals also derived


def _wais_ca_project(tmp_path: Path) -> ProjectConfig:
    candidates = tmp_path / "candidates.xlsx"
    pd.DataFrame([
        {"TEID": "1", "Study": "STDZ", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "2", "Study": "STDZ", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "3", "Study": "GT", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "4", "Study": "V_WAIS4B", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
        {"TEID": "5", "Study": "STDZ", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
    ]).to_excel(candidates, sheet_name="Candidates", index=False)
    targets = tmp_path / "WAIS 5 EN CDN Targets.xlsx"
    pd.DataFrame([
        {"Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1, "Count": 1.0, "Pcnt": 0.5},
        {"Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4, "Count": 1.0, "Pcnt": 0.5},
    ]).to_excel(targets, sheet_name="Target List", index=False)
    return ProjectConfig(
        profile_id="WAIS5_CA", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", dataset=str(candidates),
        candidate_datasets=[{"path": str(candidates), "source_type": "STDZ", "source_role": "Primary", "priority": 1}],
        study_type="", language="", forms=[],  # no project-level filter: the profile rule decides
        target_file=str(targets), sample_size=2,
        include_census=True, include_osep=False, include_clinical=False, max_solver_seconds=10,
    )


def test_profile_sample_rule_filters_the_optimization_universe(tmp_path):
    profile, _checks = validate_profile_data(_international_wais_ca())
    config = _wais_ca_project(tmp_path)
    prepared = prepare_canonical_inputs(config, profile)
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2", "5"]
    assert any("Study (profile rule) = STDZ" in note and "excluded 2" in note for note in prepared.notes)


def test_wais5_canada_end_to_end_optimization_uses_derived_marginals(tmp_path, monkeypatch):
    import scope.engine as engine_module
    from scope.engine import run

    draft = _international_wais_ca()
    # Simulate the author confirming inherited sections and marking Validated.
    draft["adaptation"]["review"] = {key: "confirmed" for key in draft["adaptation"]["review"]}
    draft["status"] = "Validated"
    profile, _checks = validate_profile_data(draft)
    config = _wais_ca_project(tmp_path)
    prepared = prepare_canonical_inputs(config, profile)
    config.canonical_targets = canonical_target_records(prepared.targets)
    config.target_review_confirmed = True
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    result = run("optimize", config)
    selected = pd.read_excel(result.outputs["optimized_sample"], sheet_name="final_all", dtype={"teid": str})
    # The profile rule (Study = STDZ) constrains eligibility even with no project filter.
    assert set(selected["teid"]) == {"5", "1"} or set(selected["teid"]) == {"5", "2"}
    adherence = pd.read_excel(result.outputs["qa_summary"], sheet_name="Target Adherence")
    region_rows = adherence[adherence["Region"].notna()]
    assert not region_rows.empty
    assert region_rows["Selected N"].tolist() == [1, 1]  # CE and W each inside their 40–60% band
    assert all("Derived from" in str(value) for value in region_rows["Source"])  # optimization ran on derived marginals
