from __future__ import annotations

import pandas as pd

from scope.configure import ProjectConfig
from scope.engine import run, validate, validate_targeted_swaps
from scope.profiles import load_profile
from scope.targeted_swap import comparable_lock_value, default_lock_fields, stage_swap_selection


def _teid_report(path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"SCOPE output": ["Participant membership lists"]}).to_excel(
            writer, sheet_name="README", index=False
        )
        pd.DataFrame({
            "Selected TEIDs (N=1)": ["100001"],
            "Available Pool TEIDs (N=1)": ["100002"],
        }).to_excel(writer, sheet_name="TEID_List", index=False)
    return path


def _celf6_universe(path):
    cases = pd.DataFrame([
        {"teid": "100001", "agey": 7, "agem": 4, "sex": "F", "eth": "WH", "ped": 3, "region": "MW", "C_LD_RF1": 1},
        {"teid": "100002", "agey": 7, "agem": 7, "sex": "F", "eth": "BL", "ped": 4, "region": "SO", "C_LD_RF1": 0},
    ])
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame([["Complete eligible CELF-6 case universe"]]).to_excel(
            writer, sheet_name="All_Usable", index=False, header=False
        )
        cases.to_excel(writer, sheet_name="All_Usable", index=False, startrow=2)
    return path


def test_celf6_teid_lists_are_hydrated_and_swapped(tmp_path):
    report = _teid_report(tmp_path / "demo_tables.xlsx")
    universe = _celf6_universe(tmp_path / "case_universe.xlsx")
    config = ProjectConfig(
        profile_id="CELF6",
        project_dir=str(tmp_path),
        workflow="targeted_swap",
        swap_sample_file=str(report),
        swap_available_file=str(report),
        swap_case_universe_file=str(universe),
        targeted_swaps=[{"remove": "100001", "add": "100002", "reason": "approved"}],
        # Old saved CELF-6 projects used BASC field names. They must migrate.
        locks=["vform", "agegrp", "sex"],
    )
    assert not [name for name, passed, _detail in validate(config) if not passed]
    result = run("targeted_swap", config)
    updated = pd.read_excel(result.outputs["sample"], sheet_name="final_all", dtype={"teid": str})
    assert updated.teid.tolist() == ["100002"]
    assert updated.columns.tolist() == ["teid"]
    assert not any("TEIDs (N=" in str(column) for column in updated.columns)
    audit = pd.read_excel(result.outputs["audit_log"], sheet_name="Swap Log")
    assert {"form", "age_group", "sex"}.issubset(audit.columns)


def test_overlapping_refreshed_source_uses_selected_membership_as_authority(tmp_path):
    selected = tmp_path / "selected.xlsx"
    available = tmp_path / "available.xlsx"
    original_data = tmp_path / "original.xlsx"
    updated_data = tmp_path / "updated.xlsx"
    pd.DataFrame({"teid": ["100001"]}).to_excel(selected, index=False)
    pd.DataFrame({"teid": ["100001", "100002"]}).to_excel(available, index=False)
    pd.DataFrame([
        {"teid": "100001", "agey": 7, "agem": 4, "sex": "F", "C_LD_RF1": 1},
    ]).to_excel(original_data, index=False)
    pd.DataFrame([
        {"teid": "100001", "agey": 7, "agem": 4, "sex": "F", "C_LD_RF1": 1},
        {"teid": "100002", "agey": 7, "agem": 7, "sex": "F", "C_LD_RF1": 0},
    ]).to_excel(updated_data, index=False)
    config = ProjectConfig(
        profile_id="CELF6", project_dir=str(tmp_path), workflow="targeted_swap",
        swap_sample_file=str(selected), swap_available_file=str(available),
        swap_sample_data_file=str(original_data), swap_available_data_file=str(updated_data),
        targeted_swaps=[{"remove": "100001", "add": "100002", "reason": "approved"}],
        locks=["form", "age_group", "sex"], output_selection="teid_lists",
    )

    checks = validate(config)
    failures = [(name, detail) for name, passed, detail in checks if not passed]
    assert not failures
    overlap = next(detail for name, passed, detail in checks if name == "MEMBERSHIP — Cross-source overlap" and passed)
    assert "1 currently selected TEID" in overlap

    result = run("targeted_swap", config)
    refreshed_selected = pd.read_excel(result.outputs["selected_cases"], sheet_name="final_all", dtype={"teid": str})
    refreshed_available = pd.read_excel(result.outputs["available_pool"], sheet_name=0, dtype={"teid": str})
    assert refreshed_selected.teid.tolist() == ["100002"]
    assert refreshed_available.teid.tolist() == ["100001"]


def test_multisheet_membership_workbook_uses_role_specific_sheet(tmp_path):
    membership = tmp_path / "membership.xlsx"
    with pd.ExcelWriter(membership, engine="openpyxl") as writer:
        pd.DataFrame([
            {"teid": "100001", "agey": 7, "agem": 4, "sex": "F", "C_LD_RF1": 1},
        ]).to_excel(writer, sheet_name="Selected Cases", index=False)
        pd.DataFrame([
            {"teid": "100002", "agey": 7, "agem": 7, "sex": "F", "C_LD_RF1": 0},
        ]).to_excel(writer, sheet_name="Available Pool", index=False)
    config = ProjectConfig(
        profile_id="CELF6", project_dir=str(tmp_path), workflow="targeted_swap",
        swap_sample_file=str(membership), swap_available_file=str(membership),
        targeted_swaps=[{"remove": "100001", "add": "100002"}],
        locks=["form", "age_group", "sex"], output_selection="teid_lists",
    )

    checks = validate(config)
    assert not [(name, detail) for name, passed, detail in checks if not passed]
    assert any(
        name == "MEMBERSHIP — Cross-source overlap" and "No TEIDs" in detail
        for name, passed, detail in checks if passed
    )


def test_separate_refreshed_inventory_combines_selected_and_available_sheets(tmp_path):
    current = tmp_path / "current.xlsx"
    refreshed = tmp_path / "refreshed.xlsx"
    pd.DataFrame([
        {"teid": "100001", "form": 1.0, "age_group": 3.0, "sex": "F"},
        {"teid": "100002", "form": 1.0, "age_group": 3.0, "sex": "F"},
    ]).to_excel(current, sheet_name="Selected", index=False)
    with pd.ExcelWriter(refreshed, engine="openpyxl") as writer:
        pd.DataFrame([
            {"teid": "100001", "form": "form_1", "age_group": "3", "sex": "F"},
            {"teid": "100003", "form": "form_1", "age_group": "3", "sex": "F"},
        ]).to_excel(writer, sheet_name="Selected", index=False)
        pd.DataFrame([
            {"teid": "100002", "form": "form_1", "age_group": "3", "sex": "F"},
            {"teid": "100004", "form": "form_1", "age_group": "3", "sex": "F"},
        ]).to_excel(writer, sheet_name="Available", index=False)

    config = ProjectConfig(
        profile_id="CELF6", project_dir=str(tmp_path), workflow="targeted_swap",
        swap_sample_file=str(current), swap_available_file=str(refreshed),
        targeted_swaps=[{"remove": "100002", "add": "100003"}],
        locks=["form", "age_group", "sex"], output_selection="teid_lists",
    )
    checks = validate(config)
    assert not [(name, detail) for name, passed, detail in checks if not passed]
    overlap = next(detail for name, passed, detail in checks if name == "MEMBERSHIP — Cross-source overlap" and passed)
    assert "2 currently selected TEID" in overlap

    result = run("targeted_swap", config)
    selected = pd.read_excel(result.outputs["selected_cases"], sheet_name="final_all", dtype={"teid": str})
    available = pd.read_excel(result.outputs["available_pool"], sheet_name=0, dtype={"teid": str})
    assert set(selected.teid) == {"100001", "100003"}
    assert set(available.teid) == {"100002", "100004"}


def test_teid_lists_request_case_universe_when_lock_fields_are_absent(tmp_path):
    report = _teid_report(tmp_path / "demo_tables.xlsx")
    profile = load_profile("CELF6")
    try:
        validate_targeted_swaps(
            report, report, [{"remove": "100001", "add": "100002"}],
            profile=profile,
        )
    except ValueError as exc:
        assert "Case universe" in str(exc)
    else:
        raise AssertionError("TEID-only inputs should require a case universe")


def test_reversed_membership_stage_has_actionable_validation_message(tmp_path):
    current = tmp_path / "refreshed_selected.xlsx"
    available = tmp_path / "available.xlsx"
    original_data = tmp_path / "original.xlsx"
    updated_data = tmp_path / "updated.xlsx"
    pd.DataFrame({"teid": ["200"]}).to_excel(current, index=False)
    pd.DataFrame({"teid": ["200"]}).to_excel(available, index=False)
    pd.DataFrame([{"teid": "100", "form": 1, "age_group": 3, "sex": "F"}]).to_excel(original_data, index=False)
    pd.DataFrame([{"teid": "200", "form": "form_1", "age_group": "3", "sex": "F"}]).to_excel(updated_data, index=False)

    try:
        validate_targeted_swaps(
            current, available, [{"remove": "100", "add": "200"}],
            ["form", "age_group", "sex"], profile=load_profile("CELF6"),
            sample_data_file=original_data, available_data_file=updated_data,
        )
    except ValueError as exc:
        message = str(exc)
        assert "pre-swap list" in message
        assert "Original dataset" in message
        assert "Updated dataset" in message
    else:
        raise AssertionError("Reversed membership stages should produce actionable guidance")


def test_celf6_profile_owns_celf6_lock_names():
    assert default_lock_fields(load_profile("CELF6")) == ["form", "age_group", "sex"]


def test_celf6_lock_comparison_normalizes_sas_and_updated_dataset_codes():
    profile = load_profile("CELF6")
    assert comparable_lock_value(1.0, "form", profile) == comparable_lock_value("form_1", "form", profile)
    assert comparable_lock_value(3.0, "age_group", profile) == comparable_lock_value("3", "age_group", profile)


def test_targeted_swap_accepts_equivalent_sas_and_updated_lock_formats(tmp_path):
    selected = tmp_path / "selected.xlsx"
    available = tmp_path / "available.xlsx"
    original_data = tmp_path / "original.xlsx"
    updated_data = tmp_path / "updated.xlsx"
    pd.DataFrame({"teid": ["625588"]}).to_excel(selected, index=False)
    pd.DataFrame({"teid": ["625588", "639970"]}).to_excel(available, index=False)
    pd.DataFrame([
        {"teid": "625588", "form": 1.0, "age_group": 3.0, "sex": "F"},
    ]).to_excel(original_data, index=False)
    pd.DataFrame([
        {"teid": "625588", "form": "form_1", "age_group": "3", "sex": "F"},
        {"teid": "639970", "form": "form_1", "age_group": "3", "sex": "F"},
    ]).to_excel(updated_data, index=False)

    checks = validate_targeted_swaps(
        selected, available, [{"remove": "625588", "add": "639970"}],
        ["form", "age_group", "sex"], profile=load_profile("CELF6"),
        sample_data_file=original_data, available_data_file=updated_data,
    )
    assert not [(name, detail) for name, passed, detail in checks if not passed]


def test_swap_selection_staging_normalizes_a_complete_pair():
    row = stage_swap_selection(" 100002 ", " 100001 ", " approved ", [])
    assert row == {"remove": "100001", "add": "100002", "reason": "approved"}


def test_swap_selection_staging_ignores_blank_fields_but_rejects_partial_pairs():
    assert stage_swap_selection("", "", "", []) is None
    try:
        stage_swap_selection("100002", "", "", [])
    except ValueError as exc:
        assert "both the Add TEID and Remove TEID" in str(exc)
    else:
        raise AssertionError("A partial swap selection must not be silently discarded")


def test_swap_selection_staging_rejects_reused_add_or_remove_teids():
    existing = [{"remove": "100001", "add": "100002", "reason": "approved"}]
    for add, remove in (("100003", "100001"), ("100002", "100004")):
        try:
            stage_swap_selection(add, remove, "", existing)
        except ValueError as exc:
            assert "may appear only once" in str(exc)
        else:
            raise AssertionError("A reused TEID must be rejected")


def test_targeted_adjustment_can_publish_teid_lists_without_tables(tmp_path):
    report = _teid_report(tmp_path / "membership.xlsx")
    universe = _celf6_universe(tmp_path / "case_data.xlsx")
    config = ProjectConfig(
        profile_id="CELF6", project_dir=str(tmp_path), workflow="targeted_adjustment",
        swap_sample_file=str(report), swap_available_file=str(report), swap_case_universe_file=str(universe),
        targeted_swaps=[{"remove": "100001", "add": "100002", "reason": "OSEP adjustment"}],
        locks=["form", "age_group", "sex"], output_selection="teid_lists",
    )
    assert not [name for name, passed, _detail in validate(config) if not passed]
    result = run("targeted_adjustment", config)
    assert result.outputs["selected_cases"].is_file()
    assert result.outputs["available_pool"].is_file()
    assert "updated_stratification" not in result.outputs
    assert result.outputs["swap_log"].is_file()
