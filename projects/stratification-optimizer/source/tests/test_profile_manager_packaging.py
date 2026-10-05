from __future__ import annotations

from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]


def test_profile_module_import_does_not_require_pandas():
    script = """
import builtins
original = builtins.__import__
def guarded(name, *args, **kwargs):
    if name == 'pandas' or name.startswith('pandas.'):
        raise ModuleNotFoundError("pandas intentionally unavailable")
    return original(name, *args, **kwargs)
builtins.__import__ = guarded
from scope.profiles import discover_profile_catalog
assert callable(discover_profile_catalog)
"""
    result = subprocess.run([sys.executable, "-c", script], cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


def test_lazy_public_workflow_api_remains_available():
    import scope

    assert callable(scope.run)
    assert callable(scope.validate)
    assert callable(scope.create_template)


def test_profile_manager_import_does_not_require_pandas():
    """The standalone Profile Manager EXE excludes pandas; the GUI must still import."""
    script = """
import builtins
original = builtins.__import__
def guarded(name, *args, **kwargs):
    if name == 'pandas' or name.startswith('pandas.'):
        raise ModuleNotFoundError("pandas intentionally unavailable")
    return original(name, *args, **kwargs)
builtins.__import__ = guarded
from gui.profile_manager import ProfileManager, build_profile_draft, normalized_workflow_inputs
assert callable(build_profile_draft) and callable(normalized_workflow_inputs)
"""
    result = subprocess.run([sys.executable, "-c", script], cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


def test_standalone_build_script_bundles_openpyxl_and_excludes_pandas():
    """Keep the EXE build aligned with the Profile Manager's packaging constraints."""
    script = (ROOT / "misc" / "Build_Scope_Profile_Manager_EXE.ps1").read_text(encoding="utf-8")
    assert "'--exclude-module', 'pandas'" in script
    assert "'--hidden-import', 'openpyxl'" in script
    assert "gui\\profile_manager.py" in script


def test_new_profile_schema_helpers_work_without_pandas():
    """International-adaptation helpers must run in the pandas-free standalone build."""
    script = """
import builtins
original = builtins.__import__
def guarded(name, *args, **kwargs):
    if name == 'pandas' or name.startswith('pandas.'):
        raise ModuleNotFoundError("pandas intentionally unavailable")
    return original(name, *args, **kwargs)
builtins.__import__ = guarded
from scope.profiles import derive_marginal_target_rows, normalize_sample_filter, profile_optimization_rules, validate_profile_data
rows = derive_marginal_target_rows(
    [{"criteria": {"sex": "F", "region": "W"}, "target": 25.0, "unit": "%"},
     {"criteria": {"sex": "M", "region": "W"}, "target": 25.0, "unit": "%"},
     {"criteria": {"sex": "F", "region": "CE"}, "target": 30.0, "unit": "%"},
     {"criteria": {"sex": "M", "region": "CE"}, "target": 20.0, "unit": "%"}],
    ["region"], {"Region": {"tolerance_percent": 10}},
)
by_group = {row["group"]: row for row in rows}
assert by_group["W"]["rule"] == "Range" and by_group["W"]["target"] == "40–60"
assert by_group["CE"]["target"] == "40–60"
assert normalize_sample_filter({"field": "study_type", "values": "STDZ"})["values"] == ["STDZ"]
payload = {"id": "X", "display_name": "X", "version": "1"}
profile, _checks = validate_profile_data(payload)
assert profile is not None
"""
    result = subprocess.run([sys.executable, "-c", script], cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr

