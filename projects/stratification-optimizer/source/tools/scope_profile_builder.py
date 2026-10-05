"""Developer-only Test Profile administration for SCOPE.

This utility is intentionally separate from the coworker-facing executable.
It accepts a complete JSON specification so schema mappings, eligibility,
targets, reporting templates, and required tables can be peer-reviewed and
versioned before installation.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import yaml

from scope.profiles import _from_dict, list_profiles, save_user_profile, validate_profile_package


WORKFLOWS = ("recruitment_planning", "sample_optimization", "targeted_swap")


def validate_spec(data: dict, package_dir: str | Path = "") -> list[str]:
    profile = _from_dict(data, package_dir)
    errors: list[str] = []
    if not profile.schema_mapping or "teid" not in profile.schema_mapping:
        errors.append("schema_mapping must define teid")
    if not profile.eligibility_rules:
        errors.append("eligibility_rules must be documented")
    for workflow, enabled in (profile.capabilities or {}).items():
        if enabled and workflow in WORKFLOWS and not profile.requirements_for(workflow):
            errors.append(f"{workflow} is enabled but has no workflow_inputs")
    reporting = profile.reporting or {}
    required = reporting.get("required_tables") if isinstance(reporting, dict) else None
    if not isinstance(required, dict):
        errors.append("reporting.required_tables must be defined")
    for workflow in WORKFLOWS:
        if profile.supports(workflow) and not profile.report_template(workflow):
            errors.append(f"{workflow} is enabled but has no report template")
    for name, passed, detail in validate_profile_package(profile):
        if not passed:
            errors.append(f"{name}: {detail}")
    return errors


def draft_template(identifier: str, display_name: str) -> dict:
    return {
        "id": identifier, "display_name": display_name, "version": "0.1",
        "status": "Draft", "runnable": False, "source_config": "",
        "forms": [], "demographics": [], "target_flags": [],
        "capabilities": {workflow: False for workflow in WORKFLOWS},
        "schema_mapping": {"teid": "teid", "form": "", "age_group": "", "sex": "", "ethnicity": "", "education": "", "region": ""},
        "age_structure": [], "eligibility_rules": [], "languages": [], "study_types": [], "target_families": ["Census", "OSEP", "Clinical"],
        "workflow_inputs": {workflow: [] for workflow in WORKFLOWS},
        "reporting": {
            "recruitment_template": "builtin:field_research",
            "optimization_template": "builtin:soe_stratification",
            "targeted_swap_template": "builtin:targeted_swap",
            "required_tables": {"demographics": [], "demo_cross_tabs": [], "osep": [], "clinical": []},
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="SCOPE administrator Test Profile manager")
    sub = parser.add_subparsers(dest="command", required=True)
    create = sub.add_parser("create", help="Create a draft profile specification")
    create.add_argument("profile_id"); create.add_argument("display_name"); create.add_argument("output", type=Path)
    check = sub.add_parser("validate", help="Validate a profile specification")
    check.add_argument("spec", type=Path)
    install = sub.add_parser("install", help="Validate and install a profile specification")
    install.add_argument("spec", type=Path)
    sub.add_parser("list", help="List installed profiles")
    args = parser.parse_args()
    if args.command == "create":
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(draft_template(args.profile_id, args.display_name), indent=2), encoding="utf-8")
        print(args.output); return 0
    if args.command == "list":
        for profile in list_profiles(): print(f"{profile.id}\t{profile.version}\t{profile.status}\t{profile.display_name}")
        return 0
    raw = args.spec.read_text(encoding="utf-8")
    data = yaml.safe_load(raw) if args.spec.suffix.casefold() in {".yaml", ".yml"} else json.loads(raw)
    errors = validate_spec(data, args.spec.parent)
    if errors:
        print("Profile validation failed:", file=sys.stderr)
        for error in errors: print(f"- {error}", file=sys.stderr)
        return 2
    profile = _from_dict(data, args.spec.parent)
    print(f"{profile.display_name}\nProfile v{profile.version}\n[PASS] VALIDATED")
    for workflow in WORKFLOWS:
        print(f"{workflow.replace('_', ' ').title()} {'[PASS]' if profile.supports(workflow) else '[OFF]'}")
    if args.command == "install":
        saved = save_user_profile(data)
        print(f"Installed {saved.id}. Normal SCOPE will show it only when status is Validated.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
