from pathlib import Path
from dataclasses import replace

import pandas as pd

import scope.engine as engine_module
from scope.canonical import build_target_specification, canonical_target_records, prepare_canonical_inputs
from scope.configure import ProjectConfig
from scope.engine import run, validate
from scope.profiles import validate_profile_data, workflow_readiness


def _profile():
    payload = {
        "id": "WAIS_CA_TEST",
        "display_name": "WAIS Canada Test",
        "version": "1.0",
        "status": "Validated",
        "runnable": True,
        "source_config": "generic_cp_sat",
        "capabilities": {"sample_optimization": True},
        "schema_mapping": {
            "teid": "TEID", "age_group": "Target Age Band", "sex": "Sex",
            "region": "Region", "ethnicity": "Ethnicity", "education": "PED/SED",
        },
        "demographics": ["sex", "eth", "edl", "region"],
        "target_families": ["Census"],
        "default_target_n": 2,
        "workflow_inputs": {
            "sample_optimization": [
                {"role": "master_dataset", "label": "Candidate case data / TEID list", "group": "CASE DATA", "status": "required", "kind": "file", "accepted_types": [".xlsx"], "adapter": "canonical_table"},
                {"role": "target_specification", "label": "Demographic target workbook", "group": "TARGETS", "status": "required", "kind": "file", "accepted_types": [".xlsx"], "adapter": "canonical_table"},
            ]
        },
        "sampling_specification": {"engine_config": "generic_cp_sat", "default_target_n": 2, "target_families": ["Census"]},
        "reporting_specification": {"required_tables": {}},
        "output_specification": {"required_resources": []},
    }
    profile, _checks = validate_profile_data(payload)
    assert profile is not None
    return profile


def _project(tmp_path: Path) -> ProjectConfig:
    candidate = tmp_path / "candidates.xlsx"
    targets = tmp_path / "targets.xlsx"
    pd.DataFrame([
        {"TEID": "1", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "2", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "3", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
        {"TEID": "4", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
    ]).to_excel(candidate, sheet_name="Candidates", index=False)
    pd.DataFrame([
        {"Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1, "Count": 1.0, "Pcnt": 0.5},
        {"Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4, "Count": 1.0, "Pcnt": 0.5},
    ]).to_excel(targets, sheet_name="Target List", index=False)
    return ProjectConfig(
        profile_id="WAIS_CA_TEST", project_dir=str(tmp_path), project_name="WAIS Canada",
        workflow="sample_optimization", output_dir=str(tmp_path / "outputs"), output_selection="both",
        dataset=str(candidate), candidate_datasets=[{"path": str(candidate), "source_type": "STDZ", "source_role": "Primary", "priority": 1}],
        target_file=str(targets), sample_size=2, include_census=True, include_osep=False, include_clinical=False,
        max_solver_seconds=10,
    )


def test_generic_profile_needs_no_research_plan_forms_age_structure_or_exclusions():
    result = workflow_readiness(_profile())["sample_optimization"]
    assert result["status"] == "READY"
    assert not result["missing"]


def test_fully_crossed_target_workbook_and_generic_cp_sat_run(tmp_path, monkeypatch):
    profile = _profile()
    config = _project(tmp_path)
    targets = build_target_specification(config, profile)
    assert len(targets) == 3  # two demographic cells plus the project Target N
    assert targets.loc[0, "criteria"]
    config.canonical_targets = canonical_target_records(targets)
    config.target_review_confirmed = True
    assert not [(name, detail) for name, passed, detail in validate(config, profile) if not passed and name != "READY TO RUN SCOPE"]
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    result = run("optimize", config)
    selected = pd.read_excel(result.outputs["optimized_sample"], sheet_name="final_all", dtype={"teid": str})
    adherence = pd.read_excel(result.outputs["qa_summary"], sheet_name="Target Adherence")
    assert len(selected) == 2
    assert set(selected["teid"]).issubset({"1", "2", "3", "4"})
    assert result.metrics["solver_status"] == "Optimal"
    assert adherence["Selected N"].tolist() == [1, 1]
    assert result.outputs["demographic_tables"].is_file()


def test_teid_only_candidates_get_actionable_demographic_error(tmp_path):
    profile = _profile()
    config = _project(tmp_path)
    pd.DataFrame({"TEID": ["1", "2", "3"]}).to_excel(config.dataset, index=False)
    failures = [(name, detail) for name, passed, detail in validate(config, profile) if not passed]
    assert any("target field" in name.casefold() and "candidate TEID list needs" in detail for name, detail in failures)


def test_numeric_age_group_codes_match_profile_age_band_labels(tmp_path, monkeypatch):
    profile = replace(_profile(), reporting_variables={
        "Age": {
            "canonical": "age_group", "field": "agegrp", "display": "Age Band",
            "levels": [
                {"value": "1", "label": "16:0–17:11"},
                {"value": "2", "label": "18:0–19:11"},
            ],
        }
    })
    config = _project(tmp_path)
    pd.DataFrame([
        {"TEID": "1", "Target Age Band": 1, "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "2", "Target Age Band": 1, "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "3", "Target Age Band": 2, "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
        {"TEID": "4", "Target Age Band": 2, "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
    ]).to_excel(config.dataset, sheet_name="Candidates", index=False)
    targets = build_target_specification(config, profile)
    config.canonical_targets = canonical_target_records(targets)
    config.target_review_confirmed = True
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    result = run("optimize", config)
    adherence = pd.read_excel(result.outputs["qa_summary"], sheet_name="Target Adherence")
    assert adherence["Selected N"].tolist() == [1, 1]


def test_study_type_filters_the_actual_generic_cp_sat_universe(tmp_path, monkeypatch):
    profile = _profile()
    config = _project(tmp_path)
    pd.DataFrame([
        {"TEID": "1", "Study": "STDZ", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "2", "Study": "STDZ", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
        {"TEID": "3", "Study": "GT", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
        {"TEID": "4", "Study": "V_WAIS4B", "Target Age Band": "18.00.00-19.11.30", "Sex": "M", "Region": "W", "Ethnicity": "WH", "PED/SED": 4},
        {"TEID": "5", "Study": "ID", "Target Age Band": "16.00.00-17.11.30", "Sex": "F", "Region": "CE", "Ethnicity": "AS", "PED/SED": 1},
    ]).to_excel(config.dataset, sheet_name="Candidates", index=False)
    prepared = prepare_canonical_inputs(config, profile)
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2"]
    assert any("Study Type = STDZ" in note and "excluded 3" in note for note in prepared.notes)

    config.canonical_targets = canonical_target_records(prepared.targets)
    config.target_review_confirmed = True
    monkeypatch.setattr(engine_module, "load_profile", lambda _profile_id: profile)
    result = run("optimize", config)
    selected = pd.read_excel(result.outputs["optimized_sample"], sheet_name="final_all", dtype={"teid": str})
    assert set(selected["teid"]) == {"1", "2"}
    assert set(selected["study_type"]) == {"STDZ"}
