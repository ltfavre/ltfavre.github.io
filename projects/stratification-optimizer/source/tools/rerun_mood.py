"""Refresh existing Mood membership and regenerate CP-SAT deliverables."""
import argparse
from datetime import datetime
import os
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    base = Path(r"C:\Users\UFAVRLA\PythonProjects")
    parser.add_argument("--scored", type=Path, default=base / "BASC4_NB_Scoring/qa/official_data/Mood/Mood_Scores")
    parser.add_argument("--sample", type=Path, default=base / "BASC4_Prenorming/outputs/deliverables/17sep26/BASC4NB_Mood_TEIDlist_post_stratification_17sep26.xlsx")
    parser.add_argument("--targets", type=Path, default=base / "BASC4_Prenorming/documentation/Demo and OSEP Targets Template.xlsx")
    parser.add_argument("--output", type=Path, default=base / "BASC4_Prenorming/outputs/mood_scoring_refresh" / datetime.now().strftime("%Y%m%d_%H%M%S"))
    args = parser.parse_args()
    for path in (args.scored, args.sample, args.targets):
        if not path.exists():
            parser.error(f"Input does not exist: {path}")
    args.output.mkdir(parents=True, exist_ok=True)
    env = {k: v for k, v in os.environ.items() if not k.startswith("OPT_ENGINE_")}
    env.update({
        "OPT_ENGINE_TEST_CONFIG": "basc4_mood",
        "OPT_ENGINE_PROJECT_DIR": str(base / "BASC4_Prenorming"),
        "OPT_ENGINE_MASTER_INPUT": str(args.sample.resolve()),
        "OPT_ENGINE_MASTER_SHEET": "selected_teids",
        "OPT_ENGINE_SCORED_DATA_DIR": str(args.scored.resolve()),
        "OPT_ENGINE_TARGET_WORKBOOK": str(args.targets.resolve()),
        "OPT_ENGINE_PART_A_OUTPUT_ROOT": str(args.output.resolve() / "prepare"),
        "OPT_ENGINE_PART_B_OUTPUT_ROOT": str(args.output.resolve() / "optimize"),
        "OPT_ENGINE_PART_A_FILE": str(args.output.resolve() / "prepare/Optimization_Engine_Part_A_prep_output_latest.xlsx"),
        "OPT_ENGINE_CLINICAL_SEEDING": "0",
    })
    legacy = Path(__file__).resolve().parents[1] / "scope/legacy"
    print(f"Outputs and run.log: {args.output.resolve()}", flush=True)
    with (args.output / "run.log").open("w", encoding="utf-8") as log:
        for part in ("a", "b"):
            print(f"Running Part {part.upper()}…", flush=True)
            result = subprocess.run([sys.executable, "-u", str(legacy / f"part_{part}.py")], env=env, stdout=log, stderr=subprocess.STDOUT)
            if result.returncode:
                raise SystemExit(f"Run failed. See {args.output / 'run.log'}")
    print(f"Complete. Deliverables: {args.output.resolve() / 'optimize/latest_deliverables'}", flush=True)


if __name__ == "__main__":
    main()
