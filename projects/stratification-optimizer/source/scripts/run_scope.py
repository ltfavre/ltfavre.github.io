"""Command-line SCOPE runners, including the strict BASC legacy smoke path."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scope.configure import ProjectConfig
from scope.engine import run
from scope.runner import run_scope as run_sample_optimization


def parse_swap(value: str) -> dict:
    parts = value.split(":", 2)
    if len(parts) < 2:
        raise argparse.ArgumentTypeError("Swap format is REMOVE_TEID:ADD_TEID[:reason]")
    return {"remove": parts[0], "add": parts[1], "reason": parts[2] if len(parts) == 3 else ""}


class StagePrinter:
    def __init__(self) -> None:
        self.started = False

    def __call__(self, stage: str) -> None:
        print((" → " if self.started else "") + stage, end="\n" if stage == "PASS" else "", flush=True)
        self.started = True


def _strict_inputs(args: argparse.Namespace) -> dict[str, object]:
    if args.smoke_test:
        fixture = ROOT / "tests" / "fixtures" / "basc_legacy"
        return {
            "initial_teid_source": str(fixture / "fr_teid_v1.xlsx"),
            "scored_data": str(fixture / "scored_data"),
            "allow_parquet_fixtures": True,
            "max_solver_seconds": args.max_solver_seconds or 300,
        }
    if args.inputs:
        return args.inputs
    values = {
        "initial_teid_source": args.initial_teid_source,
        "master_dataset": args.master_dataset,
        "master_sheet": args.master_sheet,
        "current_sample": args.current_sample,
        "scored_data": args.scored_data,
        "target_specification": args.target_specification,
        "case_review_data": args.case_review_data,
        "max_solver_seconds": args.max_solver_seconds,
    }
    return {key: value for key, value in values.items() if value not in (None, "")}


def _strict_output(args: argparse.Namespace) -> Path:
    if args.output_dir:
        return Path(args.output_dir)
    if args.smoke_test:
        stamp = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
        return ROOT / "scratch" / f"basc_smoke_{stamp}"
    raise ValueError("--output-dir is required (except when --smoke-test selects a timestamped scratch folder).")


def main() -> int:
    # The smoke-test contract uses a Unicode stage arrow. Explicit UTF-8 keeps
    # it printable from Windows shells whose inherited code page is cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Run SCOPE through the shared engine.")
    parser.add_argument("action", nargs="?", choices=("optimize", "prepare", "profile", "review_targets", "targeted_swap", "recruitment_plan"))
    parser.add_argument("--config", help="Path to scope_project.json (legacy shared-engine actions)")
    parser.add_argument("--sample", help="Optimized sample workbook (profile/targeted_swap)")
    parser.add_argument("--available", help="Available pool workbook (targeted_swap)")
    parser.add_argument("--swap", action="append", type=parse_swap, default=[], help="REMOVE:ADD[:reason]; repeatable")
    parser.add_argument("--confirm-targets", action="store_true", help="With review_targets, save the displayed canonical targets as confirmed")
    parser.add_argument("--profile", default="BASC4_NB", help="Strict runner profile (currently BASC4_NB)")
    parser.add_argument("--workflow", default="sample_optimization", help="Strict runner workflow")
    parser.add_argument("--inputs", help="Project-input JSON for strict Sample Optimization")
    parser.add_argument("--output-dir", help="Strict runner output directory")
    parser.add_argument("--initial-teid-source")
    parser.add_argument("--master-dataset")
    parser.add_argument("--master-sheet")
    parser.add_argument("--current-sample")
    parser.add_argument("--scored-data")
    parser.add_argument("--target-specification")
    parser.add_argument("--case-review-data")
    parser.add_argument("--max-solver-seconds", type=int)
    parser.add_argument("--smoke-test", action="store_true", help="Run the self-contained BASC legacy regression fixture")
    parser.add_argument("--verbose", action="store_true", help="Print legacy Part A/Part B detail during a strict run")
    args = parser.parse_args()
    strict = args.smoke_test or args.inputs or args.output_dir or args.initial_teid_source or args.master_dataset or args.current_sample or args.scored_data
    if strict:
        stages = StagePrinter()
        try:
            result = run_sample_optimization(
                profile=args.profile, workflow=args.workflow, project_inputs=_strict_inputs(args),
                output_dir=_strict_output(args), progress=stages,
                log=(lambda line: print(f"\n{line}", flush=True)) if args.verbose else None,
            )
        except Exception as exc:
            print(f"\nFAIL: {exc}", file=sys.stderr)
            return 1
        print(result.output_dir)
        return 0
    if not args.action or not args.config:
        parser.error("Supply an action with --config, or use the strict BASC options such as --inputs/--output-dir or --smoke-test.")
    config = ProjectConfig.load(args.config)
    kwargs = {"sample_file": args.sample} if args.action == "profile" else {}
    if args.action == "targeted_swap":
        sample = args.sample or config.swap_sample_file
        available = args.available or config.swap_available_file
        swaps = args.swap or config.targeted_swaps
        if not sample or not available or not swaps:
            parser.error("targeted_swap requires sample, available pool, and swaps in scope_project.json or command options")
        kwargs.update(sample_file=sample, available_file=available, swaps=swaps)
    if args.action == "review_targets": kwargs["confirm"] = args.confirm_targets
    result = run(args.action, config, progress=print, **kwargs)
    print(result.message)
    print(result.run_dir)
    for name, path in result.outputs.items(): print(f"{name}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
