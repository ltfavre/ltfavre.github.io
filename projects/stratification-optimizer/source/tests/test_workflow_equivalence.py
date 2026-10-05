"""Regression guards: GUI and CLI remain thin clients of the shared engine."""
from pathlib import Path

import pandas as pd

from scope.configure import ProjectConfig
from scope.engine import run


ROOT = Path(__file__).resolve().parents[1]


def test_gui_contains_no_solver_implementation():
    source = (ROOT / "gui" / "app.py").read_text(encoding="utf-8")
    assert "from scope.engine import run, validate" in source
    assert "CpModel(" not in source
    assert "apply_targeted_swaps(" not in source
    assert "Complete! Everything is ready for review." in source
    assert "Cute animal party" not in source
    for animal in ("kitty", "bunny", "puppy", "guinea_pig", "success_party", "attention_party"):
        assert animal in source
        assert (ROOT / "gui" / "assets" / f"{animal}.png").is_file()


def test_cli_uses_the_same_engine_entry_point():
    source = (ROOT / "scripts" / "run_scope.py").read_text(encoding="utf-8")
    assert "from scope.engine import run" in source
    assert "result = run(" in source


def test_saved_project_and_python_swap_paths_are_equivalent(tmp_path):
    sample = tmp_path / "sample.xlsx"; available = tmp_path / "available.xlsx"
    pd.DataFrame([
        {"teid": "1", "vform": "A", "agegrp": "6", "sex": "F", "eth": "A"},
        {"teid": "3", "vform": "A", "agegrp": "6", "sex": "M", "eth": "B"},
    ]).to_excel(sample, sheet_name="final_all", index=False)
    pd.DataFrame([
        {"teid": "2", "vform": "A", "agegrp": "6", "sex": "F", "eth": "B"},
    ]).to_excel(available, sheet_name="available_pool", index=False)
    swaps = [{"remove": "1", "add": "2", "reason": "Retest"}]
    gui_dir = tmp_path / "gui"; script_dir = tmp_path / "script"; gui_dir.mkdir(); script_dir.mkdir()
    gui_config = ProjectConfig(profile_id="BASC4_NB", project_dir=str(tmp_path), workflow="targeted_swap", output_dir=str(gui_dir),
        swap_sample_file=str(sample), swap_available_file=str(available), targeted_swaps=swaps, locks=["vform", "agegrp", "sex"])
    loaded = ProjectConfig.load(gui_config.save(tmp_path / "scope_project.json")); loaded.output_dir = str(gui_dir)
    script_config = ProjectConfig(**{**loaded.__dict__, "output_dir": str(script_dir)})
    gui_result = run("targeted_swap", loaded)
    script_result = run("targeted_swap", script_config, sample_file=sample, available_file=available, swaps=swaps)
    gui_sample = pd.read_excel(gui_result.outputs["sample"], sheet_name="final_all", dtype={"teid": str}).sort_values("teid").reset_index(drop=True)
    script_sample = pd.read_excel(script_result.outputs["sample"], sheet_name="final_all", dtype={"teid": str}).sort_values("teid").reset_index(drop=True)
    pd.testing.assert_frame_equal(gui_sample, script_sample)
    gui_qa = pd.read_excel(gui_result.outputs["qa"], sheet_name=None)
    script_qa = pd.read_excel(script_result.outputs["qa"], sheet_name=None)
    assert gui_qa.keys() == script_qa.keys()
    for sheet in gui_qa:
        left, right = gui_qa[sheet], script_qa[sheet]
        if "timestamp" in left.columns:
            left, right = left.drop(columns="timestamp"), right.drop(columns="timestamp")
        pd.testing.assert_frame_equal(left, right)
