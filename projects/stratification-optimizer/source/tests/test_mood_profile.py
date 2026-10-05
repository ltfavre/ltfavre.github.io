from pathlib import Path

from scope.legacy.test_config import load_test_config
from scope.profiles import load_profile, validate_profile_package


def test_mood_execution_resolves_selected_teids_sheet(tmp_path):
    import pandas as pd
    from scope.configure import ProjectConfig
    from scope.engine import environment
    workbook = tmp_path / "mood.xlsx"
    with pd.ExcelWriter(workbook) as book:
        pd.DataFrame({"teid": ["1"]}).to_excel(book, sheet_name="selected_teids", index=False)
        pd.DataFrame({"teid": ["2"]}).to_excel(book, sheet_name="available_pool", index=False)
    config = ProjectConfig(profile_id="BASC4_NB_MOOD", project_dir=str(tmp_path), master_file=str(workbook))
    env = environment(config, load_profile(config.profile_id))
    assert env["OPT_ENGINE_MASTER_SHEET"] == "selected_teids"
    assert pd.read_excel(env["OPT_ENGINE_MASTER_INPUT"], sheet_name=env["OPT_ENGINE_MASTER_SHEET"]).teid.tolist() == [1]


def test_mood_profile_routes_to_official_mood_scored_sources():
    profile = load_profile("BASC4_NB_MOOD")
    assert profile.source_config == "basc4_mood"
    assert all(passed for _, passed, _ in validate_profile_package(profile))
    config = load_test_config(profile.source_config)
    assert config.project == "MOOD"
    assert config.target_flag == "tmp_EBD"
    assert config.final_total_n == 2400
    for form, rule in config.pool_rules.items():
        assert rule.file == f"basc4nb_mood_scored_{form.replace('_', '').lower()}_be.sas7bdat"
    folder = config.default_paths.scored_data_dir.replace("\\", "/")
    assert folder.lower().endswith(config.expected_scored_source)
    assert load_profile("BASC4_NB").source_config == "basc4_adhd"
