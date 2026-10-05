"""End-to-end regression for the only workflow covered by the strict runner."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pandas as pd

from scope.runner import STAGES, run_scope


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "basc_legacy"


def _teids(path: Path, sheet: str) -> list[str]:
    frame = pd.read_excel(path, sheet_name=sheet, dtype={"teid": str})
    return sorted(frame["teid"].dropna().astype(str).str.strip().tolist())


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _teid_sha256(values: list[str]) -> str:
    return hashlib.sha256("\n".join(sorted(values)).encode("utf-8")).hexdigest()


def test_basc_scope_run_matches_raw_legacy_baseline(tmp_path):
    stages: list[str] = []
    result = run_scope(
        project_inputs={
            "initial_teid_source": FIXTURE / "fr_teid_v1.xlsx",
            "scored_data": FIXTURE / "scored_data",
            "allow_parquet_fixtures": True,
            "max_solver_seconds": 300,
        },
        output_dir=tmp_path / "run",
        progress=stages.append,
    )

    assert stages == list(STAGES)
    assert result.metrics["part_a_master_n"] == 2246
    assert result.metrics["part_a_available_n"] == 814
    assert result.metrics["candidate_pool_n"] == 3060
    assert result.metrics["final_n"] == 2400
    assert result.metrics["final_available_n"] == 660
    assert result.metrics["all_exact_n_constraints_met"] is True
    assert result.metrics["all_exact_sex_constraints_met"] is True

    legacy_part_a = result.legacy_outputs["part_a"]
    master_ids = _teids(legacy_part_a, "master_prepped")
    available_ids = _teids(legacy_part_a, "available_pool")
    candidate_ids = sorted(set(master_ids) | set(available_ids))
    assert _teid_sha256(master_ids) == "ccfba171beebb36c3c8e0dfe5dc6a687ed24985ed1c3026c0e4405cf634b2f63"
    assert _teid_sha256(available_ids) == "db9083144ec718829162637f90e5f9e0ecae6ec0e8b73e3356c24bb6a46bacf0"
    assert _teid_sha256(candidate_ids) == "958e689006a9cacea7d5186a9e0d9edb1a647282658fde59fcab054a7f81cc57"

    # SCOPE must package exactly what legacy Part A/Part B produced.
    assert _teids(result.outputs["selected_teids"], "final_all") == _teids(
        result.legacy_outputs["selected_teids"], "final_all"
    )
    assert _teids(result.outputs["available_pool"], "available_pool") == _teids(
        result.legacy_outputs["available_pool_and_qa"], "remaining_available_pool"
    )
    assert _sha256(result.outputs["stratification"]) == _sha256(result.legacy_outputs["stratification"])

    qa = pd.read_excel(result.legacy_outputs["available_pool_and_qa"], sheet_name=None)
    swaps = qa["part_b_osep_swaps"]
    assert not swaps.empty
    for metric in ("solver_status", "solver_objective", "phase1_solver_status", "phase1_osep_objective", "pre_osep_loss", "post_osep_loss"):
        assert result.metrics[metric] == swaps[metric].iloc[0]
    assert qa["sex_balance_qa"]["exact_n_match"].all()
    assert qa["sex_balance_qa"]["exact_pct_match"].all()

    lineage = pd.read_csv(result.outputs["target_lineage"])
    assert not lineage.empty
    assert set(("target", "category", "value", "source_file", "source_sheet", "source_config_key", "source_sheet_config_key", "hard_soft", "tolerance", "priority")) <= set(lineage)
    assert {"hard", "soft"} <= set(lineage["hard_soft"])
    assert not lineage["source_file"].fillna("").eq("").any()

    manifest = json.loads(result.outputs["run_manifest"].read_text(encoding="utf-8"))
    assert manifest["status"] == "PASS"
    assert manifest["profile"] == "BASC4_NB"


def test_first_run_error_does_not_request_current_sample(tmp_path):
    try:
        run_scope(project_inputs={"scored_data": FIXTURE / "scored_data", "allow_parquet_fixtures": True}, output_dir=tmp_path)
    except ValueError as exc:
        message = str(exc)
    else:
        raise AssertionError("Missing initial source should have failed")
    assert "initial_teid_source" in message
    assert "current_sample is optional" in message
