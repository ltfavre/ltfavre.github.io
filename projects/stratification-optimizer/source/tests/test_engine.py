from pathlib import Path
import sys
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scope.engine import apply_targeted_swaps
from scope.profiles import load_profile


def test_targeted_swap_preserves_membership_size(tmp_path):
    selected = pd.DataFrame({"teid": ["1"], "vform": ["TRS_C"], "agegrp": [3], "sex": ["F"], "eth": ["A"]})
    available = pd.DataFrame({"teid": ["2"], "vform": ["TRS_C"], "agegrp": [3], "sex": ["F"], "eth": ["B"]})
    selected_file, available_file = tmp_path / "selected.xlsx", tmp_path / "available.xlsx"
    selected.to_excel(selected_file, sheet_name="final_all", index=False); available.to_excel(available_file, sheet_name="available_pool", index=False)
    result = apply_targeted_swaps(selected_file, available_file, [{"remove": "1", "add": "2", "reason": "QA"}], tmp_path, load_profile("BASC4_NB"))
    assert result.outputs["sample"].is_file()
    updated = pd.read_excel(result.outputs["sample"], sheet_name="final_all", dtype={"teid": str})
    assert updated.teid.tolist() == ["2"]
