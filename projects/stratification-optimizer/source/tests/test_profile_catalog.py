import shutil
from pathlib import Path

from scope.profiles import discover_profile_catalog


ROOT = Path(__file__).resolve().parents[1]
EXPECTED_USER_PROFILES = {"BASC4_NB", "WISC6", "CELF6", "WAIS5", "VINELAND3"}


def _assert_capability_matrix(catalog):
    profiles = {profile.id: profile for profile in catalog.user_profiles}
    assert set(profiles) == EXPECTED_USER_PROFILES

    assert profiles["BASC4_NB"].supports("sample_optimization")
    for profile_id in ("WISC6", "CELF6", "WAIS5", "VINELAND3"):
        assert not profiles[profile_id].supports("sample_optimization")

    for profile_id in ("BASC4_NB", "WISC6", "CELF6"):
        assert profiles[profile_id].supports("recruitment_planning")
        assert profiles[profile_id].supports("targeted_swap")


def test_source_profile_catalog_discovers_all_normal_user_assessments():
    catalog = discover_profile_catalog([ROOT / "profiles"])
    assert not catalog.errors
    _assert_capability_matrix(catalog)
    assert "OTHER_NEW" in {profile.id for profile in catalog.profiles}
    assert "OTHER_NEW" not in {profile.id for profile in catalog.user_profiles}


def test_packaged_resource_style_catalog_discovers_all_json_descriptors(tmp_path):
    packaged_profiles = tmp_path / "_internal" / "profiles"
    packaged_profiles.mkdir(parents=True)
    for source in (ROOT / "profiles").glob("*.json"):
        shutil.copy2(source, packaged_profiles / source.name)
    catalog = discover_profile_catalog([packaged_profiles])
    assert not catalog.errors
    _assert_capability_matrix(catalog)


def test_invalid_descriptor_is_reported_instead_of_silently_dropped(tmp_path):
    broken = tmp_path / "BROKEN_TEST.json"
    broken.write_text('{"display_name": "Broken Assessment", "id": ', encoding="utf-8")
    catalog = discover_profile_catalog([tmp_path])
    assert not catalog.profiles
    assert len(catalog.errors) == 1
    assert catalog.errors[0].assessment_name == "Broken Assessment"
    assert catalog.errors[0].path == broken
    assert catalog.errors[0].message


def test_first_profile_directory_can_override_a_bundled_profile(tmp_path):
    user_root = tmp_path / "user"
    package = user_root / "WISC6"
    package.mkdir(parents=True)
    payload = {
        "id": "WISC6", "display_name": "WISC-6 Updated", "version": "2.0",
        "status": "Needs Review", "market": "Canada", "capabilities": {},
    }
    (package / "profile.yaml").write_text(__import__("yaml").safe_dump(payload), encoding="utf-8")
    catalog = discover_profile_catalog([user_root, ROOT / "profiles"])
    wisc = next(profile for profile in catalog.profiles if profile.id == "WISC6")
    assert wisc.display_name == "WISC-6 Updated"
    assert wisc.market == "Canada"


def test_gui_catalog_is_not_filtered_by_selected_workflow():
    source = (ROOT / "gui" / "app.py").read_text(encoding="utf-8")
    assert "available_profiles = [profile for profile in self.profiles" not in source
    assert "profile.status == \"Validated\" and any(profile.supports" not in source
    assert "self.profile_catalog.user_profiles" in source
    assert "profile.supports(self.workflow.get())" in source
