from __future__ import annotations

"""Generalized test/form configuration loader for the Stratification Optimizer.

Every value that used to be hardcoded per clinical form inside
"Optimization Engine Part A.py" / "Part B.py" (pool rules, clinical flag
columns, per-cell target Ns, OSEP targets, demo-target-template layout,
output file names, etc.) now lives in a YAML file under ``configs/``. This
module loads and validates that YAML into a single ``TestConfig`` object
that Part A and Part B both consume.

Select which config to use with the ``OPT_ENGINE_TEST_CONFIG`` environment
variable (a config id matching a file name under ``configs/``, e.g.
``basc4_adhd`` -> ``configs/basc4_adhd.yaml``). Defaults to ``basc4_adhd``.
"""

from dataclasses import dataclass, field
from pathlib import Path
import os

import yaml

CONFIGS_DIR = Path(__file__).resolve().parent.parent / "configs"


@dataclass
class PoolRule:
    file: str
    standard_testid: set
    exclude_testid: set
    agegrp: set


@dataclass
class OsepTarget:
    weight: float
    target_pct: float | None = None
    min_pct: float | None = None
    max_pct: float | None = None

    def as_dict(self) -> dict:
        out = {"weight": self.weight}
        if self.target_pct is not None:
            out["target_pct"] = self.target_pct
        if self.min_pct is not None:
            out["min_pct"] = self.min_pct
        if self.max_pct is not None:
            out["max_pct"] = self.max_pct
        return out


@dataclass
class DemoWriteLayout:
    agegrp_rows: dict
    srp_agegrp_rows: dict
    osep_form_to_col: dict
    osep_flag_to_row: dict
    byage_min_age: int
    byage_max_age: int


@dataclass
class DefaultPaths:
    fr_teid_file: str
    fr_admin_sheets: list
    default_master_input: str
    default_master_sheet: str
    case_swap_dir: str
    target_workbook: str
    scored_data_dir: str
    part_a_output_root: str
    part_b_output_root: str
    case_swap_files: list
    default_project_dir: str | None = None


@dataclass
class OutputNames:
    updated_teid: str
    demo_osep: str
    swap_table: str
    # Optional extra Part B deliverables (only written when populated) that a
    # few forms (e.g. Mood) add beyond the original three-workbook baseline.
    selected_teid: str | None = None
    available_pool: str | None = None


@dataclass
class EngineParams:
    max_candidates_per_remove: int = 20
    num_workers: int = 8
    scale: int = 1000
    demo_weight: int = 1
    swap_weight: int = 10
    swap_lock_cols: list = field(default_factory=lambda: ["vform", "agegrp"])
    max_solver_seconds: int = 300
    # Optional separate time budget for Part B's phase-2 (demographic
    # refinement) solve. When set, a phase-2 timeout falls back to the
    # phase-1 OSEP-optimal solution instead of raising. When None, phase 2
    # shares max_solver_seconds and a non-optimal/feasible result still
    # raises (the original ADHD-line behavior).
    phase2_max_solver_seconds: int | None = None


@dataclass
class TestConfig:
    test_id: str
    display_name: str
    project: str
    target_flag: str
    target_pct: float
    sex_col: str
    female_value: str
    male_value: str
    demo_cols: list
    flag_cols: list
    age_source_cols: list
    key_cols: list
    valid_edl_values: set
    pool_rules: dict  # vform -> PoolRule
    target_cell_n: dict  # (vform, agegrp) -> int
    osep_targets: dict  # flag -> OsepTarget
    vform_order: list
    form_families: list
    demo_target_sheets: dict
    demo_write_layout: DemoWriteLayout
    default_paths: DefaultPaths
    output_names: OutputNames
    engine_params: EngineParams
    export_cols: list
    export_priority_cols: list
    final_total_n: int | None = None

    # --- Optional behavioral feature toggles ---------------------------
    # These default to the original ADHD-line behavior (all off/None) and
    # are only turned on for forms whose legacy scripts added the feature
    # (e.g. ASD's project column + capacity check, Mood's flag/agegrp
    # derivation, phase-2 solver fallback, and extra output files).

    # Set a "project" column unconditionally in normalize_sample(). ADHD's
    # original script never set this (EXPORT_COLS silently drops it), so it
    # stays off by default to preserve ADHD's byte-identical output.
    set_project_column: bool = False
    # Run a Part B pre-check that the combined master+available universe can
    # fill every TARGET_CELL_N cell, raising a clear error otherwise.
    validate_target_capacity: bool = False
    # Rebuild Part B's final "remaining available" pool from the full source
    # universe (master + available) minus the final sample, so cases removed
    # by an OSEP swap are correctly returned to the available pool.
    rebuild_available_pool_from_universe: bool = False
    # Part B: if set, verify Part A's run_log records a scored_data_dir
    # ending with this string before proceeding (Mood-only guard against
    # running Part B against another form's Part A output).
    expected_scored_source: str | None = None
    # Derive "agegrp" from age-in-years when the source data doesn't already
    # have it (Mood's scored data lacks a precomputed agegrp column).
    derive_agegrp_if_missing: bool = False
    # Part B: drop invalid-EDL (0) rows from the starting master sample, not
    # just from the available/replacement pool. ADHD's original script
    # always did this; ASD's does not (it only cleans the available pool).
    filter_master_edl: bool = True
    # Derive missing OSEP flag columns from a "study" text column (regex per
    # flag) and/or CID indicator columns, for source data without
    # precomputed flag columns (Mood-only; empty for other forms).
    flag_study_patterns: dict = field(default_factory=dict)
    cid_groups: dict = field(default_factory=dict)


def _load_yaml(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def load_test_config(test_id: str | None = None, configs_dir: Path | None = None) -> TestConfig:
    test_id = test_id or os.environ.get("OPT_ENGINE_TEST_CONFIG", "basc4_adhd")
    configs_dir = configs_dir or CONFIGS_DIR
    path = configs_dir / f"{test_id}.yaml"
    if not path.exists():
        available = sorted(p.stem for p in configs_dir.glob("*.yaml"))
        raise FileNotFoundError(
            f"No test config named '{test_id}' found at {path}. "
            f"Available configs: {available}"
        )
    raw = _load_yaml(path)

    pool_rules = {
        vform: PoolRule(
            file=rule["file"],
            standard_testid=set(rule.get("standard_testid", [])),
            exclude_testid=set(rule.get("exclude_testid", [])),
            agegrp=set(rule.get("agegrp", [])),
        )
        for vform, rule in raw["pool_rules"].items()
    }

    target_cell_n = {
        (row["vform"], int(row["agegrp"])): int(row["n"])
        for row in raw["target_cell_n"]
    }

    osep_targets = {
        flag: OsepTarget(
            weight=rule["weight"],
            target_pct=rule.get("target_pct"),
            min_pct=rule.get("min_pct"),
            max_pct=rule.get("max_pct"),
        )
        for flag, rule in raw["osep_targets"].items()
    }

    layout_raw = raw["demo_write_layout"]
    demo_write_layout = DemoWriteLayout(
        agegrp_rows={int(k): int(v) for k, v in layout_raw["agegrp_rows"].items()},
        srp_agegrp_rows={int(k): int(v) for k, v in layout_raw["srp_agegrp_rows"].items()},
        osep_form_to_col=dict(layout_raw["osep_form_to_col"]),
        osep_flag_to_row={k: int(v) for k, v in layout_raw["osep_flag_to_row"].items()},
        byage_min_age=int(layout_raw["byage_min_age"]),
        byage_max_age=int(layout_raw["byage_max_age"]),
    )

    paths_raw = raw["default_paths"]
    default_paths = DefaultPaths(
        fr_teid_file=paths_raw["fr_teid_file"],
        fr_admin_sheets=list(paths_raw["fr_admin_sheets"]),
        default_master_input=paths_raw["default_master_input"],
        default_master_sheet=paths_raw["default_master_sheet"],
        case_swap_dir=paths_raw["case_swap_dir"],
        target_workbook=paths_raw["target_workbook"],
        scored_data_dir=paths_raw["scored_data_dir"],
        part_a_output_root=paths_raw["part_a_output_root"],
        part_b_output_root=paths_raw["part_b_output_root"],
        case_swap_files=list(paths_raw.get("case_swap_files", [])),
        default_project_dir=paths_raw.get("default_project_dir"),
    )

    names_raw = raw["output_names"]
    output_names = OutputNames(
        updated_teid=names_raw["updated_teid"],
        demo_osep=names_raw["demo_osep"],
        swap_table=names_raw["swap_table"],
        selected_teid=names_raw.get("selected_teid"),
        available_pool=names_raw.get("available_pool"),
    )

    params_raw = raw.get("engine_params", {})
    phase2_seconds = params_raw.get("phase2_max_solver_seconds")
    engine_params = EngineParams(
        max_candidates_per_remove=int(params_raw.get("max_candidates_per_remove", 20)),
        num_workers=int(params_raw.get("num_workers", 8)),
        scale=int(params_raw.get("scale", 1000)),
        demo_weight=int(params_raw.get("demo_weight", 1)),
        swap_weight=int(params_raw.get("swap_weight", 10)),
        swap_lock_cols=list(params_raw.get("swap_lock_cols", ["vform", "agegrp"])),
        max_solver_seconds=int(params_raw.get("max_solver_seconds", 300)),
        phase2_max_solver_seconds=int(phase2_seconds) if phase2_seconds is not None else None,
    )

    return TestConfig(
        test_id=test_id,
        display_name=raw.get("display_name", test_id),
        project=raw["project"],
        target_flag=raw["target_flag"],
        target_pct=raw["target_pct"],
        sex_col=raw.get("sex_col", "sex"),
        female_value=raw.get("female_value", "F"),
        male_value=raw.get("male_value", "M"),
        demo_cols=list(raw["demo_cols"]),
        flag_cols=list(raw["flag_cols"]),
        age_source_cols=list(raw["age_source_cols"]),
        key_cols=list(raw["key_cols"]),
        valid_edl_values=set(str(v) for v in raw["valid_edl_values"]),
        pool_rules=pool_rules,
        target_cell_n=target_cell_n,
        osep_targets=osep_targets,
        vform_order=list(raw["vform_order"]),
        form_families=list(raw["form_families"]),
        demo_target_sheets=dict(raw["demo_target_sheets"]),
        demo_write_layout=demo_write_layout,
        default_paths=default_paths,
        output_names=output_names,
        engine_params=engine_params,
        export_cols=list(raw["export_cols"]),
        export_priority_cols=list(raw["export_priority_cols"]),
        final_total_n=raw.get("final_total_n"),
        set_project_column=bool(raw.get("set_project_column", False)),
        validate_target_capacity=bool(raw.get("validate_target_capacity", False)),
        rebuild_available_pool_from_universe=bool(raw.get("rebuild_available_pool_from_universe", False)),
        expected_scored_source=raw.get("expected_scored_source"),
        derive_agegrp_if_missing=bool(raw.get("derive_agegrp_if_missing", False)),
        flag_study_patterns=dict(raw.get("flag_study_patterns", {})),
        cid_groups={k: list(v) for k, v in raw.get("cid_groups", {}).items()},
        filter_master_edl=bool(raw.get("filter_master_edl", True)),
    )


def list_available_configs(configs_dir: Path | None = None) -> list:
    configs_dir = configs_dir or CONFIGS_DIR
    return sorted(p.stem for p in configs_dir.glob("*.yaml"))
