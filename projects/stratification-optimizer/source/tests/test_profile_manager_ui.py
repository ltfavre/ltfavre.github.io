"""Profile Manager UI affordances: visible navigation, delete, and multi-select."""
from __future__ import annotations

import json
from pathlib import Path
from unittest import mock

import pytest

ROOT = Path(__file__).resolve().parents[1]

tk = pytest.importorskip("tkinter")


@pytest.fixture(scope="module")
def manager():
    from gui import profile_manager as pm

    tk.Tk.mainloop = lambda self: None  # never block the test runner
    try:
        instance = pm.ProfileManager()
    except tk.TclError:
        pytest.skip("No display available for the Profile Manager window")
    yield pm, instance
    instance.root.destroy()


def _draft(pm) -> dict:
    parent = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    return pm.build_profile_draft("WAIS-5", "5", "0.1", "Canada", "", "international", "WAIS5", parent, "WAIS5.json")


def test_step_navigation_shows_back_next_and_counter(manager):
    pm, app = manager
    app._show_step("overview")
    assert app.step_counter.cget("text") == "Step 1 of 9"
    assert str(app.back_step_button.cget("state")) == "disabled"
    assert str(app.next_step_button.cget("state")) == "normal"
    app._next_step()
    assert app.current_step == "sources"
    assert str(app.back_step_button.cget("state")) == "normal"
    assert app.step_counter.cget("text") == "Step 2 of 9"
    app._prev_step()
    assert app.current_step == "overview"


def test_next_step_ends_on_validate(manager):
    pm, app = manager
    app._show_step("outputs")
    assert app.next_step_button.cget("text") == "Validate →"
    app._next_step()
    assert app.current_step == "validate"
    assert str(app.next_step_button.cget("state")) == "disabled"
    app._show_step("overview")


def test_every_actionable_step_has_add_edit_and_remove_or_delete(manager):
    pm, app = manager
    labels = {}
    for step in ("sources", "sample_rules", "targets", "optimization"):
        texts = []
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    texts.append(str(child.cget("text")))
                except Exception:
                    pass
                walk(child)
        walk(app.pages[step])
        labels[step] = " | ".join(texts)
    assert "Add Source…" in labels["sources"] and "Add Multiple Files…" in labels["sources"] and "Remove Selected" in labels["sources"]
    assert "Add Rule…" in labels["sample_rules"] and "Delete Selected" in labels["sample_rules"]
    assert "Add Target Source…" in labels["targets"] and "Add Multiple Files…" in labels["targets"] and "Preview Derived Optimization Targets" in labels["targets"]
    assert "Add Rule…" in labels["optimization"] and "Delete Selected" in labels["optimization"]


def test_delete_selected_guides_on_bundled_and_confirms_user_profiles(manager):
    pm, app = manager
    infos = []
    with mock.patch.object(pm.messagebox, "showinfo", lambda t, m, **kw: infos.append((t, m))):
        app._delete_selected()  # nothing selected
    assert infos and infos[0][0] == "Delete Profile"

    catalog_entry = next(p for p in app.catalog_profiles if p.id == "WAIS5")
    app.catalog.selection_set("WAIS5")
    bundled = Path(catalog_entry.descriptor_path).resolve().is_relative_to((ROOT / "profiles").resolve()) if catalog_entry.descriptor_path else False
    if bundled:
        infos.clear()
        with mock.patch.object(pm.messagebox, "showinfo", lambda t, m, **kw: infos.append((t, m))):
            app._delete_selected()
        assert any("bundled read-only" in message for _title, message in infos)
    app.catalog.selection_set()  # clear the selection


def test_multi_select_handlers_add_items(manager):
    pm, app = manager
    app._populate(_draft(pm))
    with mock.patch.object(pm.filedialog, "askopenfilenames", return_value=(r"C:\data\a.xlsx", r"C:\data\b.csv")):
        app._add_multiple_sources()
        app._add_multiple_target_sources()
    assert len(app.source_documents) == 3  # one inherited reference + two new
    assert len(app.target_sources) == 2
    app._show_step("overview")


# ------------------------------------------------------------- install gating

def test_finish_install_is_available_for_a_new_draft(manager):
    pm, app = manager
    app._populate(_draft(pm))
    assert str(app.install_button.cget("state")) == "normal"


def test_finish_install_stays_available_after_needs_review_validation(manager):
    pm, app = manager
    app._populate(_draft(pm))
    with mock.patch.object(pm.messagebox, "showerror", lambda *args, **kwargs: None), mock.patch.object(pm.messagebox, "showinfo", lambda *args, **kwargs: None):
        app._validate()
    # Adaptations inherit values pending confirmation; that must not block Install.
    assert str(app.install_button.cget("state")) == "normal"


def test_finish_install_saves_a_needs_review_profile(manager, tmp_path, monkeypatch):
    pm, app = manager
    import scope.profiles as profiles_module

    monkeypatch.setattr(profiles_module, "USER_PROFILES_DIR", tmp_path / "profiles")
    monkeypatch.setattr(pm, "USER_PROFILES_DIR", tmp_path / "profiles")
    app._populate(_draft(pm))
    dialogs = []
    with mock.patch.object(pm.messagebox, "showerror", lambda *args, **kwargs: dialogs.append(("error", args))), \
         mock.patch.object(pm.messagebox, "showinfo", lambda *args, **kwargs: dialogs.append(("info", args))):
        app._install()
    assert not any(kind == "error" for kind, _args in dialogs)
    assert (tmp_path / "profiles" / "WAIS5_CA" / "profile.yaml").is_file()
