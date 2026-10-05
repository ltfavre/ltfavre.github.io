"""Small end-to-end workflow fixtures and GUI hand-off contracts."""
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from scope.canonical import build_target_specification, canonical_target_records
from scope.configure import ProjectConfig
from scope.engine import ProgressWriter, RunResult, run, validate
from scope.profiles import load_profile


ROOT = Path(__file__).resolve().parents[1]


def _targets(path: Path) -> Path:
    pd.DataFrame([
        {"Dimension": "Age Band", "Category": "6", "Target_N": 2},
        {"Dimension": "Age Band", "Category": "7", "Target_N": 2},
        {"Dimension": "Sex", "Category": "Female", "Target_N": 2},
        {"Dimension": "Sex", "Category": "Male", "Target_N": 2},
    ]).to_excel(path, sheet_name="Recruitment Targets", index=False)
    return path


def _cases() -> list[dict]:
    return [
        {"teid": "1", "vform": "TRS_C", "agegrp": 3, "agey": 8, "sex": "F", "eth": "WH", "edl": "3", "region": "MW", "tmp_ADHD": 1, "tmp_EBD": 0, "tmp_ASD": 0, "tmp_DD": 0, "tmp_SLD": 0, "tmp_SLI": 0, "tmp_clin": 1},
        {"teid": "2", "vform": "TRS_C", "agegrp": 3, "agey": 8, "sex": "M", "eth": "BL", "edl": "2", "region": "SO", "tmp_ADHD": 0, "tmp_EBD": 0, "tmp_ASD": 0, "tmp_DD": 0, "tmp_SLD": 1, "tmp_SLI": 0, "tmp_clin": 1},
    ]


def test_recruitment_profile_project_validate_run_results(tmp_path):
    config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), workflow="recruitment_planning", sample_size=4, recruitment_plan_file=str(_targets(tmp_path / "targets.xlsx")))
    assert not [name for name, passed, _detail in validate(config) if not passed]
    result = run("recruitment_plan", config)
    assert result.run_dir.is_dir()
    assert all(path.exists() for path in result.outputs.values())
    assert result.metrics["target_n"] == 4
    assert {"recruitment_report", "target_tables"}.issubset(result.outputs)


def test_sample_optimization_profile_project_validate_run_results(tmp_path):
    universe, current, targets, scored = tmp_path / "universe.xlsx", tmp_path / "current.xlsx", tmp_path / "canonical_targets.xlsx", tmp_path / "scored"
    scored.mkdir(); pd.DataFrame(_cases()).to_excel(universe, sheet_name="case_universe", index=False); pd.DataFrame(_cases()[:1]).to_excel(current, sheet_name="final_all", index=False)
    pd.DataFrame([
        {"Category": "Sex", "Group": "Female", "Rule": "Exact", "Target": 50, "Unit": "%", "Priority": "Required"},
        {"Category": "Sex", "Group": "Male", "Rule": "Exact", "Target": 50, "Unit": "%", "Priority": "Required"},
    ]).to_excel(targets, sheet_name="SCOPE Targets", index=False)
    config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), workflow="sample_optimization", dataset=str(universe), master_file=str(current), scored_dir=str(scored), target_file=str(targets), forms=["TRS_C"], demographics=["sex", "eth", "edl", "region"])
    profile = load_profile(config.profile_id); config.canonical_targets = canonical_target_records(build_target_specification(config, profile)); config.target_review_confirmed = True
    assert not [name for name, passed, _detail in validate(config) if not passed]

    prepare_file = tmp_path / "prepared.xlsx"; prepare_file.touch()
    run_dir = tmp_path / "optimized"; run_dir.mkdir(); sample = run_dir / "optimized_sample.xlsx"; report = run_dir / "stratification.xlsx"; teids = run_dir / "teids.xlsx"
    pd.DataFrame(_cases()).to_excel(sample, sheet_name="final_all", index=False); pd.DataFrame({"status": ["passed"]}).to_excel(report, index=False); pd.DataFrame({"teid": ["1", "2"]}).to_excel(teids, index=False)
    prepared = RunResult("prepare", tmp_path, {"primary": prepare_file}, "prepared")
    optimized = RunResult("optimize", run_dir, {"optimized_sample": sample, "stratification_report": report, "teid_list": teids}, "optimized", {"final_sample_n": 2, "hard_constraints": "Passed"})
    with patch("scope.prepare.run", return_value=prepared), patch("scope.optimize.run", return_value=optimized):
        result = run("optimize", config)
    assert result.run_dir.is_dir()
    assert all(path.exists() for path in result.outputs.values())
    assert result.metrics["hard_constraints"] == "Passed"


def test_targeted_swap_profile_project_validate_run_results(tmp_path):
    current, available = tmp_path / "current.xlsx", tmp_path / "available.xlsx"
    pd.DataFrame([{"teid": "1", "vform": "TRS_C", "agegrp": 3, "sex": "F", "eth": "WH"}]).to_excel(current, sheet_name="final_all", index=False)
    pd.DataFrame([{"teid": "2", "vform": "TRS_C", "agegrp": 3, "sex": "F", "eth": "BL"}]).to_excel(available, sheet_name="available_pool", index=False)
    config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), workflow="targeted_swap", swap_sample_file=str(current), swap_available_file=str(available), targeted_swaps=[{"remove": "1", "add": "2", "reason": "fixture"}], locks=["vform", "agegrp", "sex"])
    assert not [name for name, passed, _detail in validate(config) if not passed]
    result = run("targeted_swap", config)
    assert all(path.exists() for path in result.outputs.values())
    assert result.metrics["optimizer"] == "CP-SAT NOT RERUN"


def test_gui_runs_in_background_and_routes_failures_to_validation():
    source = (ROOT / "gui" / "app.py").read_text(encoding="utf-8")
    assert "threading.Thread(target=worker, daemon=True).start()" in source
    assert 'self.events.put(("error"' in source
    assert "self._show_step(3)" in source
    assert "verify_outputs" in (ROOT / "scope" / "engine.py").read_text(encoding="utf-8")


def test_cli_print_progress_callback_does_not_recurse():
    import contextlib

    writer = ProgressWriter(print)
    with contextlib.redirect_stdout(writer):
        print("shared progress callback")
    writer.flush()
