"""Main SCOPE app alignment with Profile Manager profile features."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

from unittest import mock

from scope.profiles import (
    describe_profile_optimization_rules, describe_target_interpretation,
    profile_study_type_defaults, validate_profile_data, workflow_readiness,
)


def _wais_ca_profile():
    parent = json.loads((ROOT / "profiles" / "WAIS5.json").read_text(encoding="utf-8"))
    parent.update({
        "id": "WAIS5_CA", "display_name": "WAIS-5 Canada", "market": "Canada",
        "profile_kind": "international_adaptation", "status": "Validated",
        "schema_mapping": {"teid": "TEID", "study_type": "Study", "sex": "Sex", "region": "Region"},
        "demographics": ["sex", "eth", "edl", "region"],
        "target_families": ["Census"],
        "study_types": ["STDZ", "Clinical"],
        "capabilities": {"sample_optimization": True},
        "source_config": "generic_cp_sat",
        "sampling_specification": {
            "engine_config": "generic_cp_sat",
            "sample_filters": [{"field": "study_type", "values": ["STDZ"], "operator": "include", "origin": "manual"}],
            "target_tolerances": {
                "Region": {"tolerance_percent": 10},
                "Age Group": {"target_n": 100},
                "Age Group | 16:0-17:11": {"applies_to": "group", "group": "16:0-17:11", "target_n": 120},
            },
            "target_interpretation": {"source_format": "crossed_cells", "cells_as": "report_only", "derive_marginals": ["region", "sex", "ethnicity", "education"]},
        },
    })
    profile, _checks = validate_profile_data(parent)
    assert profile is not None
    return profile


def test_profile_study_type_defaults_come_from_sample_rules():
    profile = _wais_ca_profile()
    assert profile_study_type_defaults(profile) == ["STDZ"]


def test_describe_profile_optimization_rules_shows_tolerances_and_target_n():
    profile = _wais_ca_profile()
    text = describe_profile_optimization_rules(profile)
    assert "Region (all groups) ±10%" in text
    assert "Age Group (all groups) = 100 N" in text
    assert "Age Group (only 16:0-17:11) = 120 N" in text


def test_describe_target_interpretation_for_crossed_cells():
    profile = _wais_ca_profile()
    text = describe_target_interpretation(profile)
    assert "crossed" in text.casefold()
    assert "marginal" in text.casefold()
    assert "report-only" in text.casefold()
    assert "Region" in text


def test_describe_helpers_tolerate_plain_profiles():
    profile, _checks = validate_profile_data({"id": "PLAIN", "display_name": "Plain", "version": "1"})
    assert profile is not None
    assert describe_profile_optimization_rules(profile) == ""
    assert describe_target_interpretation(profile) == ""
    assert profile_study_type_defaults(profile) == []


tk = pytest.importorskip("tkinter")


@pytest.fixture(scope="module")
def app_instance():
    import gui.app as app_module

    tk.Tk.mainloop = lambda self: None  # never block the test runner
    try:
        app = app_module.SCOPEApp()
    except tk.TclError:
        pytest.skip("No display available for the SCOPE window")
    yield app_module, app
    app.root.destroy()


def _install_wais_ca(app_module, app):
    from scope.profiles import validate_profile_data

    profile = _wais_ca_profile()
    app.profiles = [existing for existing in app.profiles if existing.id != "WAIS5_CA"] + [profile]
    app._build_profile_cards()  # rebuild cards so the new profile is selectable
    app._select_profile("WAIS5_CA")
    app.workflow.set("sample_optimization")
    app._refresh_profile_controls()
    app._update_workflow_buttons()
    app.root.update_idletasks()


def test_app_defaults_study_filter_from_profile_sample_rules(app_instance):
    _module, app = app_instance
    _install_wais_ca(_module, app)
    values = list(app.study_picker.cget("values"))
    assert "" in values and "STDZ" in values and "Clinical" in values
    assert app.study_type.get() == "STDZ"
    rules_text = app.profile_rules_label.cget("text")
    assert "Study = STDZ" in rules_text and "Region" in rules_text


def test_app_target_review_explains_crossed_cell_interpretation(app_instance):
    _module, app = app_instance
    _install_wais_ca(_module, app)
    app.canonical_target_rows = [
        {"category": "Region", "group": "W", "rule": "Range", "target": "40–60", "unit": "%", "source": "Derived from WAIS 5 EN CDN Targets.xlsx"},
        {"category": "Sample", "group": "Overall", "rule": "Exact", "target": 2, "unit": "N", "source": "Project Configuration"},
    ]
    app._populate_target_review()
    summary = app.target_review_summary.get()
    assert "derived marginal target" in summary
    assert "Target interpretation" in summary and "report-only" in summary


def test_validated_profile_workflow_is_not_regated_by_inherited_review(app_instance):
    """A Validated profile with the workflow enabled must be selectable and continueable."""
    module, app = app_instance
    _install_wais_ca(module, app)
    profile = app._profile()
    # The installed profile carries leftover adaptation-review readiness from the
    # Profile Manager; it must still be usable.
    profile_readiness = workflow_readiness(profile)["sample_optimization"]["status"]
    assert profile.status.casefold() == "validated"
    assert profile.supports("sample_optimization")
    assert str(app.workflow_buttons["sample_optimization"].cget("state")) == "normal"
    infos = []
    with mock.patch.object(module.messagebox, "showinfo", lambda t, m, **kw: infos.append((t, m))):
        app._choose_workflow("sample_optimization", navigate=False)
    assert app.workflow.get() == "sample_optimization"
    assert not any("not yet available" in message for _title, message in infos)
    with mock.patch.object(module.messagebox, "showinfo", lambda t, m, **kw: infos.append((t, m))):
        app._workflow_continue()
    assert 0 in app.completed and app.unlocked_step >= 1
    assert not any("Choose what SCOPE should do" in message for _title, message in infos)
