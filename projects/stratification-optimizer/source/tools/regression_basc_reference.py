"""Run the BASC first-run reference case through the shared SCOPE core.

This developer utility accepts a previously confirmed .scope/JSON project,
removes the current-sample and project-target overrides, and verifies that the
self-contained BASC profile still reaches its approved final N.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scope.canonical import canonical_target_records, prepare_canonical_inputs
from scope.configure import ProjectConfig
from scope.engine import run
from scope.profiles import load_profile


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the BASC self-contained-profile regression")
    parser.add_argument("baseline_project", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    config = ProjectConfig.load(args.baseline_project)
    profile = load_profile("BASC4_NB")
    config.profile_id = profile.id
    config.profile_version = profile.version
    config.project_dir = str(args.output)
    config.output_dir = str(args.output / "projects")
    config.project_name = "BASC4NB Reference Regression"
    config.project_file = ""
    config.master_file = ""
    config.use_existing_sample = False
    config.target_file = ""
    config.run_history = []
    args.output.mkdir(parents=True, exist_ok=True)

    # The checked-in regression fixture stores safe Parquet stand-ins beside
    # the original SAS filenames. Production projects continue to read SAS.
    original_read_sas = pd.read_sas
    def fixture_read_sas(path, *values, **options):
        source = Path(path)
        parquet = Path(str(source) + ".parquet")
        if not source.is_file() and parquet.is_file():
            return pd.read_parquet(parquet)
        return original_read_sas(path, *values, **options)
    pd.read_sas = fixture_read_sas

    prepared = prepare_canonical_inputs(config, profile)
    config.canonical_targets = canonical_target_records(prepared.targets)
    config.target_review_confirmed = True
    config.save()
    result = run("optimize", config, print)
    expected_n = int(profile.default_target_n or 0)
    actual_n = int(result.metrics.get("final_sample_n", -1))
    if actual_n != expected_n or result.metrics.get("hard_constraints") != "Passed":
        raise SystemExit(f"BASC regression failed: final N={actual_n}; expected {expected_n}; {result.metrics}")
    print(f"PASS: final N={actual_n}; hard constraints Passed")
    print(result.run_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
