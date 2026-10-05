"""Lock the September Mood rounding rules for subsequent reruns."""
import ast
from pathlib import Path

import pandas as pd


def _rounding_functions():
    source = Path(__file__).resolve().parents[1] / "scope/legacy/part_b.py"
    tree = ast.parse(source.read_text(encoding="utf-8"))
    names = {"pct_to_count", "pct_table"}
    functions = ast.Module(body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names], type_ignores=[])
    namespace = {"pd": pd, "SCALE": 1000}
    exec(compile(functions, str(source), "exec"), namespace)
    return namespace


def test_mood_report_percentages_keep_september_precision():
    functions = _rounding_functions()
    frame = pd.DataFrame({"eth": ["WH", "HI", "HI", "HI", "HI", "HI"]})
    assert functions["pct_table"](frame, "eth", ["WH", "HI", "AS"]) == {"WH": 16.667, "HI": 83.333, "AS": 0.0}
    assert functions["pct_table"](frame.iloc[:0], "eth", ["WH"]) == {"WH": 0}


def test_mood_targets_keep_fractional_counts_and_september_tie_rule():
    convert = _rounding_functions()["pct_to_count"]
    assert convert(400, 0.7) == 2800  # 2.8 cases, not rounded to 3 before solving
    assert convert(200, 7.8) == 15600
    assert convert(1, 0.25) == 2  # Python ties-to-even, as in September
    assert convert(1, 0.75) == 8
