from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scope.profiles import list_profiles, load_profile, validate_profile_data


def test_basc_profile_loads():
    profile = load_profile("BASC4_NB")
    assert profile.source_config == "basc4_adhd"
    assert "TRS_C" in profile.forms
    assert profile.status == "Validated"
    assert profile.supports("sample_optimization")
    assert profile.supports("recruitment_planning")
    requirements = {item.role: item for item in profile.requirements_for("sample_optimization")}
    assert requirements["master_dataset"].required
    assert not requirements["current_sample"].required
    assert profile.study_types == ["STDZ", "Clinical", "Validity"]
    assert profile.target_families == ["Census", "OSEP", "Clinical"]
    assert requirements["target_specification"].adapter == "basc_target_workbook"
    assert requirements["initial_teid_source"].status == "optional"


def test_celf6_profile_uses_validated_assessment_metadata():
    profile = load_profile("CELF6")
    assert profile.status == "Validated"
    assert not profile.supports("sample_optimization")
    assert profile.missing_for("sample_optimization")
    assert profile.supports("recruitment_planning")
    assert profile.supports("targeted_swap")
    assert profile.supports("targeted_adjustment")


def test_wisc6_profile_uses_validated_assessment_metadata():
    profile = load_profile("WISC6")
    assert profile.status == "Validated"
    assert profile.test_ids == ["6051"]
    assert profile.age_structure[0] == "6:0–6:11"
    assert profile.supports("recruitment_planning")
    assert profile.supports("targeted_swap")
    assert not profile.supports("sample_optimization")


def test_profiles_have_unique_ids():
    ids = [profile.id for profile in list_profiles()]
    assert len(ids) == len(set(ids))


def test_targeted_adjustment_capability_can_be_disabled_independently():
    profile, _checks = validate_profile_data({
        "id": "EXPLICIT_CAPABILITY",
        "display_name": "Explicit Capability",
        "version": "1.0",
        "runnable": True,
        "capabilities": {"targeted_swap": True, "targeted_adjustment": False},
    })
    assert profile is not None
    assert profile.supports("targeted_swap")
    assert not profile.supports("targeted_adjustment")
