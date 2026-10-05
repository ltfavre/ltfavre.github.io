from pathlib import Path

import pandas as pd

from scope.recruitment import build_recruitment_plan


def test_recruitment_plan_reconciles_margins(tmp_path: Path) -> None:
    path = tmp_path / "targets.xlsx"
    pd.DataFrame([
        {"Dimension": "Age Band", "Category": "6", "Target_N": 6},
        {"Dimension": "Age Band", "Category": "7", "Target_N": 4},
        {"Dimension": "Sex", "Category": "Female", "Target_N": 5},
        {"Dimension": "Sex", "Category": "Male", "Target_N": 5},
    ]).to_excel(path, sheet_name="Recruitment Targets", index=False)
    result = build_recruitment_plan(path, 10)
    assert result["Recruitment Cell Targets"].target_n.sum() == 10
    assert (result["Marginal Targets"].difference == 0).all()
    assert list(result["SCOPE Targets"].columns) == ["Category", "Group", "Rule", "Target", "Unit", "Tolerance_N", "Priority", "Required"]
    assert result["SCOPE Targets"].Target.sum() == 20
