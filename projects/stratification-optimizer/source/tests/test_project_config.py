import json

from scope.configure import ProjectConfig


def test_project_config_uses_grouped_storage_and_round_trips(tmp_path):
    config = ProjectConfig(
        profile_id="BASC4_NB", project_dir=str(tmp_path), workflow="targeted_swap",
        swap_sample_file="sample.xlsx", swap_available_file="pool.xlsx", swap_case_universe_file="universe.xlsx",
        swap_sample_data_file="original.sas7bdat", swap_available_data_file="updated.sas7bdat",
        targeted_swaps=[{"remove": "1", "add": "2", "reason": "Retest"}],
        forms=["TRS_C"], sample_size=100, output_selection="teid_lists",
        demographic_tables_1way=["Age", "Sex"],
        demographic_tables_2way=["Age × Sex"],
        demographic_tables_3way=["Age × Sex × Region"],
        demographic_tables_configured=True,
    )
    path = config.save()
    raw = json.loads(path.read_text(encoding="utf-8"))
    assert set(raw) == {"workflow", "test_profile", "project_configuration", "inputs", "canonical_targets", "output", "runs", "engine_version"}
    assert set(raw["inputs"]) == {"case_universe_sources", "current_sample", "target_sources", "recruitment", "targeted_swap"}
    loaded = ProjectConfig.load(path)
    assert loaded.profile_id == config.profile_id
    assert loaded.targeted_swaps == config.targeted_swaps
    assert loaded.swap_sample_file == "sample.xlsx"
    assert loaded.swap_case_universe_file == "universe.xlsx"
    assert loaded.swap_sample_data_file == "original.sas7bdat"
    assert loaded.swap_available_data_file == "updated.sas7bdat"
    assert loaded.output_selection == "teid_lists"
    assert loaded.demographic_tables_1way == ["Age", "Sex"]
    assert loaded.demographic_tables_2way == ["Age × Sex"]
    assert loaded.demographic_tables_3way == ["Age × Sex × Region"]
    assert loaded.demographic_tables_configured is True
