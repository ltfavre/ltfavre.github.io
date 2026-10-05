from pathlib import Path
import zipfile

import pandas as pd

from scope.canonical import build_target_specification, canonical_target_records, prepare_canonical_inputs
from scope.configure import ProjectConfig
from scope.engine import run, validate
from scope.profiles import load_profile


def _case_rows():
    return [
        {"teid": "1", "vform": "TRS_C", "agegrp": 3, "agey": 8, "sex": "F", "eth": "WH", "edl": "3", "region": "MW", "tmp_ADHD": 1, "tmp_EBD": 0, "tmp_ASD": 0, "tmp_DD": 0, "tmp_SLD": 0, "tmp_SLI": 0, "tmp_clin": 1},
        {"teid": "2", "vform": "TRS_C", "agegrp": 3, "agey": 8, "sex": "M", "eth": "BL", "edl": "2", "region": "SO", "tmp_ADHD": 0, "tmp_EBD": 0, "tmp_ASD": 0, "tmp_DD": 0, "tmp_SLD": 1, "tmp_SLI": 0, "tmp_clin": 1},
    ]


def _project(tmp_path: Path) -> ProjectConfig:
    master = tmp_path / "master.xlsx"; current = tmp_path / "current.xlsx"; targets = tmp_path / "targets.xlsx"; scored = tmp_path / "scored"; scored.mkdir()
    pd.DataFrame(_case_rows()).to_excel(master, sheet_name="case_universe", index=False)
    pd.DataFrame(_case_rows()[:1]).to_excel(current, sheet_name="final_all", index=False)
    pd.DataFrame([
        {"Category": "Sex", "Group": "Female", "Rule": "Exact", "Target": 50, "Unit": "%", "Priority": "Required"},
        {"Category": "Sex", "Group": "Male", "Rule": "Exact", "Target": 50, "Unit": "%", "Priority": "Required"},
    ]).to_excel(targets, sheet_name="SCOPE Targets", index=False)
    return ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), dataset=str(master), master_file=str(current), scored_dir=str(scored), target_file=str(targets), forms=["TRS_C"], demographics=["sex", "eth", "edl", "region"])


def test_profile_sources_prepare_canonical_case_universe_and_current_sample(tmp_path):
    config = _project(tmp_path); profile = load_profile("BASC4_NB")
    prepared = prepare_canonical_inputs(config, profile)
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2"]
    assert prepared.current_sample.SCOPE_ID.tolist() == ["1"]
    assert {"SCOPE_FORM", "SCOPE_AGE_GROUP", "SCOPE_SEX", "SCOPE_GROUP_TMP_ADHD"}.issubset(prepared.case_universe.columns)


def test_changed_master_selector_replaces_stale_candidate_path(tmp_path):
    config = _project(tmp_path)
    stale = tmp_path / "old_raw.csv"
    pd.DataFrame({"teid": ["9"]}).to_csv(stale, index=False)
    config.candidate_datasets = [{"path": str(stale), "source_role": "Primary", "priority": 1}]
    prepared = prepare_canonical_inputs(config, load_profile("BASC4_NB"))
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1", "2"]


def test_selected_teids_sheet_takes_precedence_over_available_pool(tmp_path):
    config = _project(tmp_path)
    with pd.ExcelWriter(config.dataset) as book:
        pd.DataFrame(_case_rows()[:1]).to_excel(book, sheet_name="selected_teids", index=False)
        pd.DataFrame(_case_rows()[1:]).to_excel(book, sheet_name="available_pool", index=False)
    prepared = prepare_canonical_inputs(config, load_profile("BASC4_NB"))
    assert prepared.case_universe.SCOPE_ID.tolist() == ["1"]


def test_target_review_round_trips_and_validation_detects_source_changes(tmp_path):
    config = _project(tmp_path); profile = load_profile("BASC4_NB")
    rows = canonical_target_records(build_target_specification(config, profile))
    config.canonical_targets = rows; config.target_review_confirmed = True
    assert not [name for name, passed, _detail in validate(config) if not passed]
    loaded = ProjectConfig.load(config.save())
    assert loaded.canonical_targets == rows
    assert loaded.target_review_confirmed
    changed = pd.read_excel(config.target_file, sheet_name="SCOPE Targets")
    changed.loc[0, "Target"] = 51
    changed.to_excel(config.target_file, sheet_name="SCOPE Targets", index=False)
    failures = [name for name, passed, _detail in validate(loaded) if not passed]
    assert "Targets — Review confirmation" in failures


def test_optional_and_conditional_profile_inputs_do_not_block(tmp_path):
    config = _project(tmp_path); profile = load_profile("BASC4_NB")
    rows = canonical_target_records(build_target_specification(config, profile))
    config.canonical_targets = rows; config.target_review_confirmed = True
    assert not config.fr_teid_file
    assert not config.case_swap_dir
    failures = [name for name, passed, _detail in validate(config) if not passed]
    assert not failures


def test_saved_gui_and_python_configs_prepare_identical_canonical_inputs(tmp_path):
    gui_config = _project(tmp_path); profile = load_profile("BASC4_NB")
    gui_config.canonical_targets = canonical_target_records(build_target_specification(gui_config, profile))
    gui_config.target_review_confirmed = True
    python_config = ProjectConfig.load(gui_config.save())
    gui_inputs = prepare_canonical_inputs(gui_config, profile)
    python_inputs = prepare_canonical_inputs(python_config, profile)
    pd.testing.assert_frame_equal(gui_inputs.case_universe, python_inputs.case_universe)
    pd.testing.assert_frame_equal(gui_inputs.current_sample, python_inputs.current_sample)
    assert canonical_target_records(gui_inputs.targets) == canonical_target_records(python_inputs.targets)


def test_python_target_review_uses_shared_canonical_builder(tmp_path):
    config = _project(tmp_path)
    result = run("review_targets", config, confirm=True)
    assert result.outputs["canonical_targets"].is_file()
    loaded = ProjectConfig.load(config.path)
    assert loaded.target_review_confirmed
    assert loaded.canonical_targets == canonical_target_records(build_target_specification(loaded, load_profile("BASC4_NB")))


def _docx_table(path: Path, rows: list[list[str]]) -> Path:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xml_rows = "".join(
        "<w:tr>" + "".join(f"<w:tc><w:p><w:r><w:t>{value}</w:t></w:r></w:p></w:tc>" for value in row) + "</w:tr>"
        for row in rows
    )
    document = f'<w:document xmlns:w="{ns}"><w:body><w:tbl>{xml_rows}</w:tbl></w:body></w:document>'
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)
    return path


def test_structured_docx_target_source_is_parsed_without_excel_routing(tmp_path):
    source = _docx_table(tmp_path / "Research Plan.docx", [
        ["Category", "Group", "Rule", "Target", "Unit", "Priority"],
        ["Sex", "Female", "Exact", "50", "%", "High"],
        ["Sex", "Male", "Exact", "50", "%", "High"],
    ])
    config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), target_file=str(source), include_osep=False, include_clinical=False)
    rows = canonical_target_records(build_target_specification(config, load_profile("BASC4_NB")))
    assert any(row["category"] == "Sex" and row["group"] == "Female" for row in rows)


def test_unstructured_docx_target_source_fails_with_clear_warning(tmp_path):
    source = _docx_table(tmp_path / "Research Plan.docx", [["Narrative"], ["No structured target table here"]])
    config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), target_file=str(source))
    try:
        build_target_specification(config, load_profile("BASC4_NB"))
    except ValueError as exc:
        assert "No canonical targets were found" in str(exc)
        assert "SCOPE Target Template" in str(exc)
    else:
        raise AssertionError("An unstructured DOCX must not silently return profile defaults")
