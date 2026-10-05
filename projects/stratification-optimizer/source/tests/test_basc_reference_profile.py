from dataclasses import replace
from pathlib import Path

import pandas as pd

from scope.canonical import resolved_profile_input
from scope.configure import ProjectConfig
from scope.engine import _output_package_checks, run, validate
from scope.profiles import load_profile, validate_profile_package


def test_basc_profile_is_self_contained_and_developer_validated():
    profile = load_profile("BASC4_NB")
    assert profile.version == "1.3"
    assert profile.status == "Validated"
    assert profile.sampling_specification["current_sample_required_for_first_run"] is False
    assert all(passed for _name, passed, _detail in validate_profile_package(profile))
    assert all(path and path.is_file() for _label, path, required in profile.resource_status() if required)


def test_missing_profile_resource_fails_package_validation(tmp_path):
    profile = load_profile("BASC4_NB")
    broken = replace(profile, resources={**profile.resources, "demographic_template": "missing.xlsx"}, package_dir=str(tmp_path))
    checks = validate_profile_package(broken)
    assert any(not passed and "Resources" in name for name, passed, _detail in checks)
    assert checks[-1][0] == "PROFILE STATUS"
    assert checks[-1][1] is False


def test_basc_approved_targets_resolve_without_project_override(tmp_path):
    profile = load_profile("BASC4_NB")
    config = ProjectConfig(profile_id=profile.id, project_dir=str(tmp_path), target_file="")
    assert resolved_profile_input(config, profile, "target_specification") == profile.resource_path("approved_targets")
    config.target_file = str(tmp_path / "project_targets.xlsx")
    assert resolved_profile_input(config, profile, "target_specification") == Path(config.target_file)


def test_basc_validation_previews_seventeen_optimization_outputs():
    checks = _output_package_checks(load_profile("BASC4_NB"), "sample_optimization")
    outputs = [item for item in checks if item[0] != "OUTPUT PACKAGE — Summary"]
    assert len(outputs) == 17
    assert all(passed for _name, passed, _detail in outputs)
    assert checks[-1][2] == "17/17 outputs ready"


def test_recruitment_uses_profile_targets_without_project_target_file(tmp_path):
    profile = load_profile("BASC4_NB")
    config = ProjectConfig(
        profile_id=profile.id, profile_version=profile.version,
        project_dir=str(tmp_path), project_name="BASC Recruitment",
        workflow="recruitment_planning", sample_size=2400,
    )
    checks = validate(config, profile)
    target_check = next(item for item in checks if item[0] == "TARGETS — Approved recruitment specification")
    assert target_check[1] is True
    assert str(profile.resource_path("recruitment_targets")) == target_check[2]


def test_targeted_swap_creates_project_run_package_and_history(tmp_path):
    selected_file, available_file = tmp_path / "selected.xlsx", tmp_path / "available.xlsx"
    selected = pd.DataFrame({"teid": ["1"], "vform": ["TRS_C"], "agegrp": ["6"], "sex": ["F"]})
    available = pd.DataFrame({"teid": ["2"], "vform": ["TRS_C"], "agegrp": ["6"], "sex": ["F"]})
    selected.to_excel(selected_file, sheet_name="final_all", index=False)
    available.to_excel(available_file, sheet_name="available_pool", index=False)
    config = ProjectConfig(
        profile_id="BASC4_NB", profile_version="1.3", project_dir=str(tmp_path),
        output_dir=str(tmp_path / "outputs"), project_name="BASC4NB Mood",
        workflow="targeted_swap", swap_sample_file=str(selected_file),
        swap_available_file=str(available_file), locks=["vform", "agegrp", "sex"],
        targeted_swaps=[{"remove": "1", "add": "2", "reason": "approved"}],
    )
    result = run("targeted_swap", config)
    assert result.run_dir.parent.name == "runs"
    assert result.run_dir.parent.parent.name == "BASC4NB_Mood"
    assert (result.run_dir / "Cases" / "Selected_TEIDs.xlsx").is_file()
    assert (result.run_dir / "Cases" / "Available_Pool.xlsx").is_file()
    assert (result.run_dir / "QA" / "Swap_Log.xlsx").is_file()
    assert (result.run_dir / "Stratification" / "Updated_Stratification.xlsx").is_file()
    assert (result.run_dir / "run_manifest.json").is_file()
    assert config.run_history[-1]["profile"] == {"id": "BASC4_NB", "version": "1.3"}
    assert config.path.is_file()
