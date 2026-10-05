from __future__ import annotations

import json
from pathlib import Path

import scope.profiles as profiles_module
from gui.profile_manager import build_profile_draft, normalized_workflow_inputs
from scope.profiles import (
    discover_profile_catalog, install_profile_package, load_profile,
    profile_validation_summary, validate_profile_data, workflow_readiness,
)


ROOT = Path(__file__).resolve().parents[1]


def test_legacy_targeted_inputs_are_migrated_to_scope_v2_sources():
    legacy = {
        "targeted_swap": [
            {"role": "swap_sample", "label": "Existing sample", "status": "required", "kind": "file"},
            {"role": "swap_available_pool", "label": "Available pool", "status": "required", "kind": "file"},
        ]
    }
    migrated = normalized_workflow_inputs(legacy, {"targeted_swap": True, "targeted_adjustment": True})
    by_role = {item["role"]: item for item in migrated["targeted_swap"]}
    assert by_role["swap_case_universe"]["accepted_types"] == [".xlsx", ".sas7bdat", ".csv"]
    assert by_role["swap_sample"]["label"] == "Current selected TEID list"
    assert by_role["swap_available_pool"]["status"] == "required"
    assert by_role["swap_sample_data"]["status"] == "optional"
    assert by_role["swap_available_data"]["status"] == "optional"
    assert by_role["swap_case_universe"]["status"] == "optional"


def test_new_cp_sat_profile_gets_default_task_inputs():
    migrated = normalized_workflow_inputs({}, {"sample_optimization": True})
    roles = {item["role"] for item in migrated["sample_optimization"]}
    assert roles == {"master_dataset", "current_sample", "target_specification"}


def test_existing_wisc6_is_discovered_valid_and_ready_without_recreation():
    catalog = discover_profile_catalog([ROOT / "profiles"])
    wisc = next(profile for profile in catalog.profiles if profile.id == "WISC6")
    summary = profile_validation_summary(wisc)
    assert summary["overall"] == "VALID"
    assert wisc.market == "United States"
    assert summary["workflows"]["targeted_swap"]["status"] == "READY"
    assert summary["workflows"]["recruitment_planning"]["status"] == "READY"
    assert summary["workflows"]["sample_optimization"]["status"] == "NOT CONFIGURED"


def test_wais5_canada_adaptation_does_not_require_a_research_plan():
    parent = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    draft = build_profile_draft("WAIS-5", "5", "0.1", "Canada", "", "adapt", "WAIS5", parent, "WAIS5.json")
    profile, _checks = validate_profile_data(draft)
    assert profile is not None
    assert draft["id"] == "WAIS5_CA"
    assert draft["status"] == "Needs Review"
    assert draft["source_documents"][0]["type"] == "Existing Profile"
    assert "research_specification" not in (draft.get("output_specification") or {}).get("required_resources", [])
    assert profile_validation_summary(profile)["core_valid"]


def test_manual_profile_without_documents_is_a_valid_draft():
    draft = build_profile_draft("Future Assessment", "1", "0.1", "Canada", "", "manual")
    profile, _checks = validate_profile_data(draft)
    summary = profile_validation_summary(profile)
    assert summary["core_valid"]
    assert not draft["source_documents"]
    assert all(item["status"] == "NOT CONFIGURED" for item in summary["workflows"].values())


def test_partial_profile_can_be_swap_ready_without_optimization():
    data = build_profile_draft("Partial Assessment", "1", "0.1", "Canada", "English", "manual")
    data.update({
        "forms": ["Form A"],
        "schema_mapping": {"teid": "teid", "form": "form", "age_group": "age_group", "sex": "sex"},
        "sampling_specification": {"default_locks": ["form", "age_group", "sex"]},
        "capabilities": {"targeted_swap": True, "sample_optimization": False},
        "workflow_inputs": normalized_workflow_inputs({}, {"targeted_swap": True}),
    })
    profile, _checks = validate_profile_data(data)
    readiness = workflow_readiness(profile)
    assert profile_validation_summary(profile)["overall"] == "VALID"
    assert readiness["targeted_swap"]["status"] == "READY"
    assert readiness["sample_optimization"]["status"] == "NOT CONFIGURED"


def test_saved_wais5_canada_reopens_from_user_catalog(tmp_path, monkeypatch):
    storage = tmp_path / "profiles"
    monkeypatch.setattr(profiles_module, "USER_PROFILES_DIR", storage)
    parent = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    draft = build_profile_draft("WAIS-5", "5", "0.1", "Canada", "", "adapt", "WAIS5", parent, "")
    install_profile_package(draft)
    reopened = discover_profile_catalog([storage])
    saved = next(profile for profile in reopened.profiles if profile.id == "WAIS5_CA")
    assert saved.display_name == "WAIS-5 Canada"
    assert saved.market == "Canada"
    assert saved.adaptation["parent_profile_id"] == "WAIS5"
