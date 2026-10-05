"""Contract test: GUI and CLI are thin callers of the same SCOPE engine."""
from pathlib import Path


def test_gui_and_script_use_shared_engine():
    root = Path(__file__).resolve().parents[1]
    gui = (root / "gui" / "app.py").read_text(encoding="utf-8")
    cli = (root / "scripts" / "run_scope.py").read_text(encoding="utf-8")
    assert "from scope.engine import run" in gui
    assert "from scope.engine import run" in cli
    assert "Current sample / master workbook" not in gui
    assert "requirements_for(workflow)" in gui
    assert "Review canonical targets" in gui
